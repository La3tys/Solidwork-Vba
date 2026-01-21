' Ensure "Microsoft Forms 2.0 Object Library" is enabled in Tools > References
Sub main()
    Dim swApp As SldWorks.SldWorks: Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2: Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    ' 1. SORTING AXIS
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
    
    ' Declare bBox once here (VBA scope is procedure-level)
    Dim bBox As Variant
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Dim swFolder As SldWorks.BodyFolder: Set swFolder = swFeat.GetSpecificFeature2
            Dim vBodies As Variant: vBodies = swFolder.GetBodies
            If Not IsEmpty(vBodies) Then
                Dim swTempBody As SldWorks.Body2: Set swTempBody = vBodies(0)
                bBox = swTempBody.GetBodyBox ' <--- Assignment 1
                
                If (bBox(3) - bBox(0)) > 0.001 Then
                    ReDim Preserve featNames(count): ReDim Preserve sortCoords(count)
                    featNames(count) = swFeat.name
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
        ' === EXPLICIT VARIABLE RESET START ===
        Dim valL As Double: valL = 0
        Dim valW As Double: valW = 0
        Dim valT As Double: valT = 0
        Dim totalPerimeter As Double: totalPerimeter = 0
        Dim faceCount As Long: faceCount = 0
        Dim exportStatus As String: exportStatus = "No"
        Dim strDesc As String: strDesc = ""
        ' === EXPLICIT VARIABLE RESET END ===

        Dim swCurrFeat As SldWorks.Feature: Set swCurrFeat = swModel.FeatureByName(featNames(i))
        Dim swFolderObj As SldWorks.BodyFolder: Set swFolderObj = swCurrFeat.GetSpecificFeature2
        Dim vBods As Variant: vBods = swFolderObj.GetBodies
        
        If Not IsEmpty(vBods) Then
            Dim swBody As SldWorks.Body2: Set swBody = vBods(0)
            faceCount = swBody.GetFaceCount
            
            ' A. CALCULATE DIMENSIONS
            ' === FIX IS HERE: Don't Re-Dim, just Assign ===
            bBox = swBody.GetBodyBox
            
            If Not IsEmpty(bBox) Then
                Dim dArr(2) As Double
                dArr(0) = bBox(3) - bBox(0) ' X Len
                dArr(1) = bBox(4) - bBox(1) ' Y Len
                dArr(2) = bBox(5) - bBox(2) ' Z Len
                
                ' Inline Bubble Sort (Smallest to Largest)
                Dim x As Integer, y As Integer, temp As Double
                For x = 0 To 1
                    For y = x + 1 To 2
                        If dArr(x) > dArr(y) Then
                            temp = dArr(x): dArr(x) = dArr(y): dArr(y) = temp
                        End If
                    Next y
                Next x
                
                valT = Round(dArr(0) * 1000, 2) ' Smallest
                valW = Round(dArr(1) * 1000, 2) ' Middle
                valL = Round(dArr(2) * 1000, 2) ' Largest
            End If
            
            ' B. GET PROPERTIES
            Dim swCustPropMgr As SldWorks.CustomPropertyManager: Set swCustPropMgr = swCurrFeat.CustomPropertyManager
            strDesc = GetDeepProp(swCustPropMgr, Array("Description", "DESCRIPTION"))
            If strDesc = "-" Then strDesc = swCurrFeat.name
            
            ' C. RENAME FOLDER
            Dim cleanDesc As String: cleanDesc = CleanFileName(strDesc)
            Dim newName As String: newName = Format(i + 1, "00") & "_" & cleanDesc
            
            If swCurrFeat.name <> newName Then
                swCurrFeat.name = "TEMP_" & i
                swCurrFeat.name = newName
            End If
            
            ' D. SMART EXPORT
            swBody.HideBody False
            
            Dim detectedAxis As String
            Dim bFaceFound As Boolean
            bFaceFound = SelectLargestFaceAndGetNormal(swBody, detectedAxis)
            
            If bFaceFound Then
                Dim fileName As String: fileName = newName & ".dxf"
                Dim vAlign As Variant: vAlign = GetMatrixForAxis(detectedAxis)
                
                Dim bRet As Boolean
                ' Export using Option 2 (as requested)
                bRet = swPart.ExportToDWG2(dxfPath & fileName, fullPath, 2, True, vAlign, False, False, 0, Nothing)
                
                If bRet Then
                    successCount = successCount + 1
                    exportStatus = "Yes"
                End If
                
                ' E. PERIMETER
                Dim normIdx As Integer
                If detectedAxis = "X" Then normIdx = 0 Else If detectedAxis = "Y" Then normIdx = 1 Else normIdx = 2
                totalPerimeter = GetGeometricPerimeter(swBody, normIdx)
            End If
            
            ' F. BUILD TABLE ROW
            Dim locVal As Double: locVal = Round(sortCoords(i) * 1000, 2)
            tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & valL & vbTab & valW & vbTab & valT & vbTab & swFolderObj.GetBodyCount & vbTab & Round(totalPerimeter, 2) & vbTab & faceCount & vbTab & exportStatus & vbTab & locVal & vbCrLf
        End If
        swModel.ClearSelection2 True
    Next i

    ' Clipboard
    Dim DataObj As Object: Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData: DataObj.PutInClipboard
    
    MsgBox "Completed! " & successCount & " DXFs exported."
    Shell "explorer.exe " & dxfPath, vbNormalFocus
End Sub

' --- HELPER FUNCTIONS ---

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
            bestNormal = swFace.Normal
        End If
        Set swFace = swFace.GetNextFace
    Loop
    
    If Not bestFace Is Nothing Then
        bestFace.Select4 False, Nothing
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
    Select Case axis
        Case "X": vData(5) = 1: vData(7) = 1: vData(9) = 1
        Case "Y": vData(3) = 1: vData(8) = 1: vData(10) = 1
        Case "Z": vData(3) = 1: vData(7) = 1: vData(11) = 1
    End Select
    GetMatrixForAxis = vData
End Function

Function GetGeometricPerimeter(swBody As SldWorks.Body2, normalIdx As Integer) As Double
    Dim swFace As SldWorks.Face2: Set swFace = swBody.GetFirstFace
    Do While Not swFace Is Nothing
        Dim vNormal As Variant: vNormal = swFace.Normal
        If Abs(vNormal(normalIdx)) > 0.99 Then
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
