' "SldWorks" is the main application object. "ModelDoc2" is your specific Part or Assembly file.
' "Feature" refers to items in your tree (like Cut-List folders).
' "Body2" is the actual 3D solid lump of metal.

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.Feature
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim swFolder As SldWorks.BodyFolder
    Dim vBodies As Variant
    Dim swBody As SldWorks.Body2
    
    Dim featNames() As String
    Dim zCoords() As Double
    Dim count As Integer: count = 0

    ' Connect to the active SolidWorks session
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    ' --- STEP 1: PREP ---
    ' RunCommand 2014 is a built-in SolidWorks ID for "Update/Sort Cut List". 
    ' This ensures the model is calculated before we read it.
    swModel.Extension.SelectByID2 "Update Cut Lists", "COMMAND", 0, 0, 0, False, 0, Nothing, 0
    
    ' --- STEP 2: SCAN THE TREE ---
    ' We start at the very first feature in your FeatureManager tree.
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        ' We only care about "CutListFolder" features.
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Set swFolder = swFeat.GetSpecificFeature2
            If Not swFolder Is Nothing Then
                vBodies = swFolder.GetBodies ' Get all solid bodies inside this folder
                If Not IsEmpty(vBodies) Then
                    Dim folderMaxZ As Double: folderMaxZ = -100000 
                    Dim bFound As Boolean: bFound = False
                    Dim k As Integer
                    
                    For k = 0 To UBound(vBodies)
                        Set swBody = vBodies(k)
                        Dim vBox As Variant
                        ' GetBodyBox gives us the [MinX, MinY, MinZ, MaxX, MaxY, MaxZ] coordinates.
                        vBox = swBody.GetBodyBox 
                        If Not IsEmpty(vBox) Then
                            ' Since your Front Plane is perp to Y, the "Up" height is MaxZ (Index 5).
                            If vBox(5) > folderMaxZ Then folderMaxZ = vBox(5)
                            bFound = True
                        End If
                    Next k
                    
                    ' Save the name and height for sorting later
                    If bFound Then
                        ReDim Preserve featNames(count)
                        ReDim Preserve zCoords(count)
                        featNames(count) = swFeat.Name
                        zCoords(count) = folderMaxZ 
                        count = count + 1
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature ' Move to the next item in the tree
    Loop

    ' --- STEP 3: BUBBLE SORT ---
    ' Traditional sorting logic to order items by their Z-Height (Top to Bottom)
    Dim i As Integer, j As Integer
    Dim tempZ As Double, tempName As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If zCoords(i) < zCoords(j) Then ' Change < to > if you want Bottom-to-Top
                tempZ = zCoords(i): zCoords(i) = zCoords(j): zCoords(j) = tempZ
                tempName = featNames(i): featNames(i) = featNames(j): featNames(j) = tempName
            End If
        Next j
    Next i

    ' --- STEP 4: DATA EXTRACTION & RENAMING ---
    Dim tableData As String
    tableData = "Order" & vbTab & "Description" & vbTab & "L" & vbTab & "W" & vbTab & "T" & vbTab & "Qty" & vbTab & "Total Perimeter (mm)" & vbTab & "Faces" & vbCrLf
    
    For i = 0 To count - 1
        Set swFeat = swModel.FeatureByName(featNames(i))
        Set swFolder = swFeat.GetSpecificFeature2
        Set swCustPropMgr = swFeat.CustomPropertyManager ' Access "File Properties" for this folder
        
        vBodies = swFolder.GetBodies
        Dim itemQty As Long: itemQty = swFolder.GetBodyCount ' Total count of parts in this folder
        Dim faceCount As Long: faceCount = 0
        Dim totalPerimeter As Double: totalPerimeter = 0
        
        If Not IsEmpty(vBodies) Then
            Set swBody = vBodies(0) ' Look at the first body in the folder for geometry
            faceCount = swBody.GetFaceCount ' Total count of connected surfaces (inner/outer)
            totalPerimeter = GetGeometricPerimeterFromLoops(swBody) ' Call our custom math function
        End If
        
        ' Resolve Custom Properties like "Description" or "Length"
        Dim strDesc As String, valOut As String, b As Boolean
        swCustPropMgr.Get6 "Description", False, valOut, strDesc, b, False
        
        Dim strL As String: strL = GetDeepProp(swCustPropMgr, Array("Length", "LENGTH", "Bounding Box Length"))
        Dim strW As String: strW = GetDeepProp(swCustPropMgr, Array("Width", "WIDTH", "Bounding Box Width"))
        Dim strT As String: strT = GetDeepProp(swCustPropMgr, Array("Thickness", "THICKNESS", "Sheet Metal Thickness"))

        ' If properties are blank, use RegEx to pull numbers out of the "Description" string
        If strL = "-" Or strW = "-" Then
            Dim dims As Variant: dims = ParseDimsFromDesc(strDesc)
            If IsArray(dims) Then
                strT = dims(0): strW = dims(1): strL = dims(2)
            End If
        End If

        ' PHYSICAL TREE SORTING: 
        ' We rename the folders with 01_, 02_ etc. SolidWorks will use these names to sort the tree.
        Dim finalName As String: finalName = Format(i + 1, "00") & "_ " & strDesc
        On Error Resume Next
        swFeat.Name = "SORT_" & i ' Temporary name to prevent "name already exists" errors
        swFeat.Name = finalName
        On Error GoTo 0
        
        ' Add data to our "Clipboard" string using TABS for Excel
        tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & strL & vbTab & strW & vbTab & strT & vbTab & itemQty & vbTab & Round(totalPerimeter, 2) & vbTab & faceCount & vbCrLf
    Next i

    ' --- STEP 5: FINAL TREE REORDER ---
    ' We select the "Cut-List" and tell SolidWorks to run its internal "Sort" command.
    ' Since we renamed them with numbers, they will move into the correct physical order.
    swModel.Extension.SelectByID2 "Cut-List", "SUBWELD", 0, 0, 0, False, 0, Nothing, 0
    swModel.Extension.RunCommand 2014, "" 

    ' Copy the result to the clipboard
    Dim DataObj As Object
    On Error Resume Next
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData
    DataObj.PutInClipboard

    swModel.ForceRebuild3 True
    MsgBox "Success! Tree sorted and data ready for Excel."
End Sub

' --- GEOMETRY HELPER ---
' This function navigates the "Topology" of the part.
Function GetGeometricPerimeterFromLoops(swBody As SldWorks.Body2) As Double
    Dim swFace As SldWorks.Face2
    Dim vNormal As Variant
    Dim vLoops As Variant
    Dim swLoop As SldWorks.Loop2
    Dim vEdges As Variant
    Dim swEdge As SldWorks.Edge
    Dim swCurve As SldWorks.Curve
    Dim vParams As Variant
    Dim i As Integer, j As Integer
    Dim totalLength As Double: totalLength = 0
    
    Set swFace = swBody.GetFirstFace ' Check the first surface of the solid
    Do While Not swFace Is Nothing
        vNormal = swFace.Normal ' Find which way the surface faces (X, Y, or Z)
        ' 0.99 check means the face points Up along the Z-axis.
        If vNormal(2) > 0.99 Then
            vLoops = swFace.GetLoops ' A face has 1 Outer loop and many Inner loops (holes)
            If Not IsEmpty(vLoops) Then
                For i = 0 To UBound(vLoops)
                    Set swLoop = vLoops(i)
                    vEdges = swLoop.GetEdges ' A loop is a chain of edges
                    If Not IsEmpty(vEdges) Then
                        For j = 0 To UBound(vEdges)
                            Set swEdge = vEdges(j)
                            Set swCurve = swEdge.GetCurve ' The "Curve" contains the math/length
                            vParams = swEdge.GetCurveParams2 ' Find where the edge starts and ends
                            ' Calculate the true length of this segment
                            totalLength = totalLength + swCurve.GetLength3(vParams(6), vParams(7))
                        Next j
                    End If
                Next i
                GetGeometricPerimeterFromLoops = totalLength * 1000 ' Convert Meters to Millimeters
                Exit Function ' Found the top face, we can stop looking!
            End If
        End If
        Set swFace = swFace.GetNextFace
    Loop
End Function
