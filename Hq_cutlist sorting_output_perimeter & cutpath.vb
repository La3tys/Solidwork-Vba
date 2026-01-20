' Ensure "Microsoft Forms 2.0 Object Library" is enabled in Tools > References
Sub main()

    ' ===== SolidWorks main application object =====
    Dim swApp As SldWorks.SldWorks
    
    ' ===== Active document (part / assembly / drawing) =====
    Dim swModel As SldWorks.ModelDoc2
    
    ' ===== Feature object (used to walk the feature tree) =====
    Dim swFeat As SldWorks.Feature
    
    ' ===== Custom Property Manager for cut-list properties =====
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    
    ' ===== Body folder (Cut List Folder) =====
    Dim swFolder As SldWorks.BodyFolder
    
    ' ===== Array holding bodies inside a cut list =====
    Dim vBodies As Variant
    
    ' ===== Individual body =====
    Dim swBody As SldWorks.Body2
    
    ' ===== Arrays to store cut-list names and Z-coordinates =====
    Dim featNames() As String
    Dim zCoords() As Double
    
    ' ===== Counter for number of cut-list items =====
    Dim count As Integer: count = 0

    ' Connect to SolidWorks
    Set swApp = Application.SldWorks
    
    ' Get currently active document
    Set swModel = swApp.ActiveDoc
    
    ' Exit if no file is open
    If swModel Is Nothing Then Exit Sub
    
    ' =========================================================
    ' 1. FORCE UPDATE CUT LIST
    ' =========================================================
    ' This ensures the cut list is rebuilt and up to date
    swModel.Extension.SelectByID2 "Update Cut Lists", "COMMAND", 0, 0, 0, False, 0, Nothing, 0
    
    ' =========================================================
    ' 2. SCAN FEATURE TREE AND GET Z-COORDINATES
    ' =========================================================
    
    ' Start at the first feature in the tree
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing
        
        ' Check if this feature is a Cut List Folder
        If swFeat.GetTypeName2 = "CutListFolder" Then
            
            ' Get the actual body folder object
            Set swFolder = swFeat.GetSpecificFeature2
            
            If Not swFolder Is Nothing Then
                
                ' Get all bodies inside this cut-list folder
                vBodies = swFolder.GetBodies
                
                If Not IsEmpty(vBodies) Then
                    
                    ' Initialize max Z value very low
                    Dim folderMaxZ As Double: folderMaxZ = -100000
                    Dim bFound As Boolean: bFound = False
                    Dim k As Integer
                    
                    ' Loop through all bodies in this folder
                    For k = 0 To UBound(vBodies)
                        
                        Set swBody = vBodies(k)
                        
                        ' Get bounding box of the body
                        Dim vBox As Variant
                        vBox = swBody.GetBodyBox
                        
                        ' vBox(5) is the MAX Z value
                        If Not IsEmpty(vBox) Then
                            If vBox(5) > folderMaxZ Then folderMaxZ = vBox(5)
                            bFound = True
                        End If
                    Next k
                    
                    ' Store the cut-list name and its max Z
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
        
        ' Move to next feature in the tree
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' =========================================================
    ' 3. SORT CUT LIST ITEMS BY Z (HIGHEST FIRST)
    ' =========================================================
    
    Dim i As Integer, j As Integer
    Dim tempZ As Double, tempName As String
    
    ' Simple bubble sort
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If zCoords(i) < zCoords(j) Then
                tempZ = zCoords(i): zCoords(i) = zCoords(j): zCoords(j) = tempZ
                tempName = featNames(i): featNames(i) = featNames(j): featNames(j) = tempName
            End If
        Next j
    Next i

    ' =========================================================
    ' 4. COLLECT DATA AND BUILD TABLE
    ' =========================================================
    
    ' Header row (tab-separated for Excel)
    Dim tableData As String
    tableData = "Order" & vbTab & "Description" & vbTab & "L" & vbTab & _
                "W" & vbTab & "T" & vbTab & "Qty" & vbTab & _
                "Total Perimeter (mm)" & vbTab & "Faces" & vbCrLf
    
    For i = 0 To count - 1
        
        ' Find cut-list feature by name
        Set swFeat = swModel.FeatureByName(featNames(i))
        Set swFolder = swFeat.GetSpecificFeature2
        
        ' Get property manager
        Set swCustPropMgr = swFeat.CustomPropertyManager
        
        ' Bodies and quantity
        vBodies = swFolder.GetBodies
        Dim itemQty As Long: itemQty = swFolder.GetBodyCount
        
        ' Default values
        Dim faceCount As Long: faceCount = 0
        Dim totalPerimeter As Double: totalPerimeter = 0
        
        If Not IsEmpty(vBodies) Then
            Set swBody = vBodies(0)
            faceCount = swBody.GetFaceCount
            
            ' Calculate true geometric perimeter (outer + inner loops)
            totalPerimeter = GetGeometricPerimeterFromLoops(swBody)
        End If
        
        ' Get Description property
        Dim strDesc As String, valOut As String, b As Boolean
        swCustPropMgr.Get6 "Description", False, valOut, strDesc, b, False
        
        ' Try multiple property names for dimensions
        Dim strL As String: strL = GetDeepProp(swCustPropMgr, Array("Length", "LENGTH", "Bounding Box Length"))
        Dim strW As String: strW = GetDeepProp(swCustPropMgr, Array("Width", "WIDTH", "Bounding Box Width"))
        Dim strT As String: strT = GetDeepProp(swCustPropMgr, Array("Thickness", "THICKNESS", "Sheet Metal Thickness"))

        ' Fallback: extract numbers from description text
        If strL = "-" Or strW = "-" Then
            Dim dims As Variant: dims = ParseDimsFromDesc(strDesc)
            If IsArray(dims) Then
                strT = dims(0): strW = dims(1): strL = dims(2)
            End If
        End If

        ' Rename cut-list item to enforce tree order
        Dim finalName As String: finalName = Format(i + 1, "00") & "_ " & strDesc
        On Error Resume Next
        swFeat.Name = "SORT_" & i
        swFeat.Name = finalName
        On Error GoTo 0
        
        ' Append row to output table
        tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & _
                    strL & vbTab & strW & vbTab & strT & vbTab & _
                    itemQty & vbTab & Round(totalPerimeter, 2) & vbTab & _
                    faceCount & vbCrLf
    Next i

    ' =========================================================
    ' 5. RE-SORT CUT LIST TREE AND COPY TO CLIPBOARD
    ' =========================================================
    
    ' Select Cut List folder
    swModel.Extension.SelectByID2 "Cut-List", "SUBWELD", 0, 0, 0, False, 0, Nothing, 0
    
    ' Force SolidWorks to reorder the tree
    swModel.Extension.RunCommand 2014, ""
    
    ' Copy data to clipboard
    Dim DataObj As Object
    On Error Resume Next
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData
    DataObj.PutInClipboard

    ' Final rebuild
    swModel.ForceRebuild3 True
    
    MsgBox "Success! Tree sorted and Loop-based perimeter copied."
End Sub


' UPDATED: Handles Inner and Outer Loops explicitly
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
    
    Set swFace = swBody.GetFirstFace
    Do While Not swFace Is Nothing
        vNormal = swFace.Normal
        ' Look for the face pointing Up (+Z)
        If vNormal(2) > 0.99 Then
            vLoops = swFace.GetLoops
            If Not IsEmpty(vLoops) Then
                For i = 0 To UBound(vLoops)
                    Set swLoop = vLoops(i)
                    vEdges = swLoop.GetEdges
                    If Not IsEmpty(vEdges) Then
                        For j = 0 To UBound(vEdges)
                            Set swEdge = vEdges(j)
                            Set swCurve = swEdge.GetCurve
                            vParams = swEdge.GetCurveParams2
                            ' Summing every segment of the loop
                            totalLength = totalLength + swCurve.GetLength3(vParams(6), vParams(7))
                        Next j
                    End If
                Next i
                GetGeometricPerimeterFromLoops = totalLength * 1000 ' Meters to mm
                Exit Function
            End If
        End If
        Set swFace = swFace.GetNextFace
    Loop
    GetGeometricPerimeterFromLoops = 0
End Function

Function GetDeepProp(mgr As SldWorks.CustomPropertyManager, names As Variant) As String
    Dim i As Integer, val As String, res As String, b As Boolean
    For i = LBound(names) To UBound(names)
        mgr.Get6 CStr(names(i)), False, val, res, b, False
        If res <> "" And Not res Like "*@*" Then
            GetDeepProp = Trim(Replace(res, "mm", ""))
            Exit Function
        End If
    Next i
    GetDeepProp = "-"
End Function

Function ParseDimsFromDesc(desc As String) As Variant
    On Error Resume Next
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object: Dim results(2) As String
    regEx.Global = True: regEx.Pattern = "[0-9.]+"
    Set matches = regEx.Execute(desc)
    If matches.count >= 3 Then
        results(0) = matches(0): results(1) = matches(1): results(2) = matches(2)
        ParseDimsFromDesc = results
    Else: ParseDimsFromDesc = ""
    End If
End Function




