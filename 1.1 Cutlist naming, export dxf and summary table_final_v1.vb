' Ensure "Microsoft Forms 2.0 Object Library" is enabled in Tools > References
Sub main()
    Dim swApp As SldWorks.SldWorks: Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2: Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    ' 1. SORTING AXIS (Does not affect View Orientation anymore)
    Dim userDir As String: userDir = UCase(InputBox("Select axis for sorting sequence (X, Y, or Z):", "Sort Order", "Z"))
    If InStr("XYZ", userDir) = 0 Then Exit Sub
    
    ' 2. PREP OUTPUT
    Dim fullPath As String: fullPath = swModel.GetPathName
    If fullPath = "" Then MsgBox "Save part first.": Exit Sub
    
    Dim basePath As String: basePath = Left(fullPath, InStrRev(fullPath, "\"))
    Dim dxfPath As String: dxfPath = basePath & "DXF_Exports\"
    If Dir(dxfPath, vbDirectory) = "" Then MkDir dxfPath

    ' 3. GATHER BODIES
    swModel.Extension.SelectByID2 "Update Cut Lists", "COMMAND", 0, 0, 0, False, 0, Nothing, 0
    Dim swFeat As SldWorks.Feature: Set swFeat = swModel.FirstFeature
    Dim featNames() As String, sortCoords() As Double, count As Integer: count = 0
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Dim swFolder As SldWorks.BodyFolder: Set swFolder = swFeat.GetSpecificFeature2
            Dim vBodies As Variant: vBodies = swFolder.GetBodies
            If Not IsEmpty(vBodies) Then
                Dim swTempBody As SldWorks.Body2: Set swTempBody = vBodies(0)
                Dim bBox As Variant: bBox = swTempBody.GetBodyBox
                
                ' Filter small artifacts
                If (bBox(3) - bBox(0)) > 0.001 Then
                    ReDim Preserve featNames(count): ReDim Preserve sortCoords(count)
                    featNames(count) = swFeat.name
                    
                    ' Sorting Coordinate
                    Dim idx As Integer: If userDir = "X" Then idx = 3 Else If userDir = "Y" Then idx = 4 Else idx = 5
                    sortCoords(count) = bBox(idx)
                    count = count + 1
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    If count = 0 Then MsgBox "No valid bodies found.": Exit Sub

    ' 4. SORT LIST
    Dim i As Integer, j As Integer, tempC As Double, tempN As String
    For i = 0 To count - 2: For j = i + 1 To count - 1
        If sortCoords(i) < sortCoords(j) Then
            tempC = sortCoords(i): sortCoords(i) = sortCoords(j): sortCoords(j) = tempC
            tempN = featNames(i): featNames(i) = featNames(j): featNames(j) = tempN
        End If
    Next j: Next i

    ' 5. EXPORT LOOP
    Dim tableData As String
    tableData = "Order" & vbTab & "Description" & vbTab & "Length" & vbTab & "Width" & vbTab & "Thickness" & vbTab & "Qty" & vbTab & "Perimeter" & vbTab & "Faces" & vbTab & "Exported" & vbTab & "Location " & userDir & vbCrLf
    
    Dim swPart As SldWorks.PartDoc: Set swPart = swModel
    Dim successCount As Integer: successCount = 0
    swModel.ClearSelection2 True
    
    For i = 0 To count - 1
        Dim swCurrFeat As SldWorks.Feature: Set swCurrFeat = swModel.FeatureByName(featNames(i))
        Dim swFolderObj As SldWorks.BodyFolder: Set swFolderObj = swCurrFeat.GetSpecificFeature2
        Dim vBods As Variant: vBods = swFolderObj.GetBodies
        
        If Not IsEmpty(vBods) Then
            Dim swBody As SldWorks.Body2: Set swBody = vBods(0)
            
            ' A. CALCULATE DIMENSIONS (Smart Sort L > W > T)
            Dim bBox As Variant: bBox = swBody.GetBodyBox
            Dim dArr(2) As Double
            dArr(0) = bBox(3) - bBox(0) ' X Len
            dArr(1) = bBox(4) - bBox(1) ' Y Len
            dArr(2) = bBox(5) - bBox(2) ' Z Len
            Call SortArray(dArr) ' Result: dArr(0)=Thickness, dArr(1)=Width, dArr(2)=Length
            
            Dim valT As Double: valT = Round(dArr(0) * 1000, 2)
            Dim valW As Double: valW = Round(dArr(1) * 1000, 2)
            Dim valL As Double: valL = Round(dArr(2) * 1000, 2)
            
            ' B. GET PROPERTIES
            Dim swCustPropMgr As SldWorks.CustomPropertyManager: Set swCustPropMgr = swCurrFeat.CustomPropertyManager
            Dim strDesc As String: strDesc = GetDeepProp(swCustPropMgr, Array("Description", "DESCRIPTION"))
            If strDesc = "-" Then strDesc = swCurrFeat.name
            
            ' C. RENAME CUT-LIST FOLDER
            Dim cleanDesc As String: cleanDesc = CleanFileName(strDesc)
            Dim newName As String: newName = Format(i + 1, "00") & "_" & cleanDesc
            swCurrFeat.name = "TEMP_" & i
            swCurrFeat.name = newName
            
            ' D. SMART EXPORT (Surface Detection)
            swBody.HideBody False
            
            ' 1. Select the Largest Face (Length x Width)
            ' This function returns the Normal Axis of that face ("X", "Y", or "Z")
            Dim detectedAxis As String
            Dim bFaceFound As Boolean
            bFaceFound = SelectLargestFaceAndGetNormal(swBody, detectedAxis)
            
            Dim exportStatus As String: exportStatus = "No"
            If bFaceFound Then
                Dim fileName As String: fileName = newName & ".dxf"
                
                ' 2. Get Matrix matching that Face's Normal
                Dim vAlign As Variant
                vAlign = GetMatrixForAxis(detectedAxis)
                
                ' 3. Export
                Dim bRet As Boolean
                ' Using Option 1 (ExportSelectedFacesOrLoops) combined with the Matrix
                ' ensures we view the selected face "Head On"
                bRet = swPart.ExportToDWG2(dxfPath & fileName, fullPath, 2, True, vAlign, False, False, 0, Nothing)
                
                If bRet Then
                    successCount = successCount + 1
                    exportStatus = "Yes"
                End If
                
                ' E. PERIMETER (Calculate on the detected face axis)
                Dim normIdx As Integer
                If detectedAxis = "X" Then normIdx = 0 Else If detectedAxis = "Y" Then normIdx = 1 Else normIdx = 2
                Dim totalPerimeter As Double: totalPerimeter = GetGeometricPerimeter(swBody, normIdx)
                
                Dim locVal As Double: locVal = Round(sortCoords(i) * 1000, 2)
                tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & valL & vbTab & valW & vbTab & valT & vbTab & swFolderObj.GetBodyCount & vbTab & Round(totalPerimeter, 2) & vbTab & swBody.GetFaceCount & vbTab & exportStatus & vbTab & locVal & vbCrLf
            End If
        End If
        swModel.ClearSelection2 True
    Next i

    ' Clipboard
    Dim DataObj As Object: Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData: DataObj.PutInClipboard
    
    MsgBox "Completed! " & successCount & " DXFs exported." & vbCrLf & "Views aligned to largest surface."
    Shell "explorer.exe " & dxfPath, vbNormalFocus
End Sub

' --- HELPER FUNCTIONS ---

' BUBBLE SORT ARRAY (Smallest to Largest)
Sub SortArray(ByRef arr() As Double)
    Dim x As Integer, y As Integer, temp As Double
    For x = LBound(arr) To UBound(arr) - 1
        For y = x + 1 To UBound(arr)
            If arr(x) > arr(y) Then
                temp = arr(x): arr(x) = arr(y): arr(y) = temp
            End If
        Next y
    Next x
End Sub

' NEW: Finds the largest face regardless of axis, selects it, and returns which axis it faces
Function SelectLargestFaceAndGetNormal(body As SldWorks.Body2, ByRef axisOut As String) As Boolean
    Dim swFace As SldWorks.Face2: Set swFace = body.GetFirstFace
    Dim bestFace As SldWorks.Face2
    Dim maxArea As Double: maxArea = -1
    Dim bestNormal As Variant
    
    Do While Not swFace Is Nothing
        Dim area As Double: area = swFace.GetArea
        If area > maxArea Then
            maxArea = area
            Set bestFace = swFace
            bestNormal = swFace.Normal ' Store normal of largest face
        End If
        Set swFace = swFace.GetNextFace
    Loop
    
    If Not bestFace Is Nothing Then
        bestFace.Select4 False, Nothing
        
        ' Determine Axis from Normal Vector
        ' Normal is (X, Y, Z). If X is close to 1 or -1, it's X-Axis.
        If Abs(bestNormal(0)) > 0.9 Then
            axisOut = "X"
        ElseIf Abs(bestNormal(1)) > 0.9 Then
            axisOut = "Y"
        Else
            axisOut = "Z"
        End If
        
        SelectLargestFaceAndGetNormal = True
    Else
        SelectLargestFaceAndGetNormal = False
    End If
End Function

Function GetMatrixForAxis(axis As String) As Variant
    Dim vData(11) As Double
    ' The view matrix: X-vector, Y-vector, Z-vector (Normal)
    Select Case axis
        Case "X"
            vData(5) = 1: vData(7) = 1: vData(9) = 1 ' Normal = X
        Case "Y"
            vData(3) = 1: vData(8) = 1: vData(10) = 1 ' Normal = Y
        Case "Z"
            vData(3) = 1: vData(7) = 1: vData(11) = 1 ' Normal = Z
    End Select
    GetMatrixForAxis = vData
End Function

Function GetGeometricPerimeter(swBody As SldWorks.Body2, normalIdx As Integer) As Double
    Dim swFace As SldWorks.Face2: Set swFace = swBody.GetFirstFace
    Do While Not swFace Is Nothing
        Dim vNormal As Variant: vNormal = swFace.Normal
        If Abs(vNormal(normalIdx)) > 0.99 Then
            ' Found a face matching our export axis
            Dim vLoops As Variant: vLoops = swFace.GetLoops
            If Not IsEmpty(vLoops) Then
                Dim i As Integer, totalLength As Double: totalLength = 0
                For i = 0 To UBound(vLoops)
                    Dim swLoop As SldWorks.Loop2: Set swLoop = vLoops(i)
                    Dim vEdges As Variant: vEdges = swLoop.GetEdges
                    Dim j As Integer
                    For j = 0 To UBound(vEdges)
                        Dim swEdge As SldWorks.Edge: Set swEdge = vEdges(j)
                        Dim swCurve As SldWorks.Curve: Set swCurve = swEdge.GetCurve
                        Dim vParams As Variant: vParams = swEdge.GetCurveParams2
                        totalLength = totalLength + swCurve.GetLength3(vParams(6), vParams(7))
                    Next j
                Next i
                GetGeometricPerimeter = totalLength * 1000
                Exit Function
            End If
        End If
        Set swFace = swFace.GetNextFace
    Loop
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

Function CleanFileName(strIn As String) As String
    Dim invalidChars As String: invalidChars = "\/:*?""<>|"
    Dim i As Integer: CleanFileName = Trim(strIn)
    For i = 1 To Len(invalidChars)
        CleanFileName = Replace(CleanFileName, Mid(invalidChars, i, 1), "_")
    Next i
End Function
