' Ensure "Microsoft Forms 2.0 Object Library" is enabled in Tools > References
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

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    ' 1. FORCE UPDATE CUT LIST
    swModel.Extension.SelectByID2 "Update Cut Lists", "COMMAND", 0, 0, 0, False, 0, Nothing, 0
    
    ' 2. SCAN AND GET COORDINATES (Z-AXIS)
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Set swFolder = swFeat.GetSpecificFeature2
            If Not swFolder Is Nothing Then
                vBodies = swFolder.GetBodies
                If Not IsEmpty(vBodies) Then
                    Dim folderMaxZ As Double: folderMaxZ = -100000
                    Dim bFound As Boolean: bFound = False
                    Dim k As Integer
                    
                    For k = 0 To UBound(vBodies)
                        Set swBody = vBodies(k)
                        Dim vBox As Variant
                        vBox = swBody.GetBodyBox
                        If Not IsEmpty(vBox) Then
                            If vBox(5) > folderMaxZ Then folderMaxZ = vBox(5)
                            bFound = True
                        End If
                    Next k
                    
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
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' 3. SORTING (Highest Z first)
    Dim i As Integer, j As Integer
    Dim tempZ As Double, tempName As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If zCoords(i) < zCoords(j) Then
                tempZ = zCoords(i): zCoords(i) = zCoords(j): zCoords(j) = tempZ
                tempName = featNames(i): featNames(i) = featNames(j): featNames(j) = tempName
            End If
        Next j
    Next i

    ' 4. DATA COLLECTION
    Dim tableData As String
    tableData = "Order" & vbTab & "Description" & vbTab & "L" & vbTab & "W" & vbTab & "T" & vbTab & "Qty" & vbTab & "Total Perimeter (mm)" & vbTab & "Faces" & vbCrLf
    
    For i = 0 To count - 1
        Set swFeat = swModel.FeatureByName(featNames(i))
        Set swFolder = swFeat.GetSpecificFeature2
        Set swCustPropMgr = swFeat.CustomPropertyManager
        
        vBodies = swFolder.GetBodies
        Dim itemQty As Long: itemQty = swFolder.GetBodyCount
        Dim faceCount As Long: faceCount = 0
        Dim totalPerimeter As Double: totalPerimeter = 0
        
        If Not IsEmpty(vBodies) Then
            Set swBody = vBodies(0)
            faceCount = swBody.GetFaceCount
            ' This now uses the Inner/Outer Loop Logic
            totalPerimeter = GetGeometricPerimeterFromLoops(swBody)
        End If
        
        ' Get Properties
        Dim strDesc As String, valOut As String, b As Boolean
        swCustPropMgr.Get6 "Description", False, valOut, strDesc, b, False
        
        Dim strL As String: strL = GetDeepProp(swCustPropMgr, Array("Length", "LENGTH", "Bounding Box Length"))
        Dim strW As String: strW = GetDeepProp(swCustPropMgr, Array("Width", "WIDTH", "Bounding Box Width"))
        Dim strT As String: strT = GetDeepProp(swCustPropMgr, Array("Thickness", "THICKNESS", "Sheet Metal Thickness"))

        If strL = "-" Or strW = "-" Then
            Dim dims As Variant: dims = ParseDimsFromDesc(strDesc)
            If IsArray(dims) Then
                strT = dims(0): strW = dims(1): strL = dims(2)
            End If
        End If

        ' Rename for Tree Sorting
        Dim finalName As String: finalName = Format(i + 1, "00") & "_ " & strDesc
        On Error Resume Next
        swFeat.Name = "SORT_" & i
        swFeat.Name = finalName
        On Error GoTo 0
        
        tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & strL & vbTab & strW & vbTab & strT & vbTab & itemQty & vbTab & Round(totalPerimeter, 2) & vbTab & faceCount & vbCrLf
    Next i

    ' 5. RE-SORT TREE & CLIPBOARD
    swModel.Extension.SelectByID2 "Cut-List", "SUBWELD", 0, 0, 0, False, 0, Nothing, 0
    swModel.Extension.RunCommand 2014, ""

    Dim DataObj As Object
    On Error Resume Next
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData
    DataObj.PutInClipboard

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
