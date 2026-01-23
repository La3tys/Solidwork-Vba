' Ensure "Microsoft Forms 2.0 Object Library" is enabled in Tools > References
Sub main()
    Dim swApp As SldWorks.SldWorks: Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2: Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    ' 1. AXIS SELECTION
    Dim userDir As String: userDir = UCase(InputBox("Select view axis (X, Y, or Z):", "DXF Export Axis", "Z"))
    If InStr("XYZ", userDir) = 0 Then Exit Sub
    
    ' Determine Index for Perimeter Function (X=0, Y=1, Z=2)
    Dim normalIndex As Integer
    If userDir = "X" Then normalIndex = 0 Else If userDir = "Y" Then normalIndex = 1 Else normalIndex = 2
    
    ' Define Alignment Matrix
    Dim vData(11) As Double
    Select Case userDir
        Case "X": vData(5) = 1: vData(7) = 1: vData(9) = 1
        Case "Y": vData(3) = 1: vData(8) = 1: vData(10) = 1
        Case "Z": vData(3) = 1: vData(7) = 1: vData(11) = 1
    End Select
    Dim vAlignment As Variant: vAlignment = vData

    ' 2. PREP FOLDER
    Dim fullPath As String: fullPath = swModel.GetPathName
    If fullPath = "" Then MsgBox "Save part first.": Exit Sub
    
    Dim basePath As String: basePath = Left(fullPath, InStrRev(fullPath, "\"))
    Dim dxfPath As String: dxfPath = basePath & "DXF_Exports\"
    If Dir(dxfPath, vbDirectory) = "" Then MkDir dxfPath

    ' 3. SCAN TREE & GET COORDINATES
    swModel.Extension.SelectByID2 "Update Cut Lists", "COMMAND", 0, 0, 0, False, 0, Nothing, 0
    Dim swFeat As SldWorks.Feature: Set swFeat = swModel.FirstFeature
    Dim featNames() As String, coords() As Double, count As Integer: count = 0
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Dim swFolder As SldWorks.BodyFolder: Set swFolder = swFeat.GetSpecificFeature2
            Dim vBodies As Variant: vBodies = swFolder.GetBodies
            If Not IsEmpty(vBodies) Then
                Dim swTempBody As SldWorks.Body2: Set swTempBody = vBodies(0)
                If swTempBody.GetBodyBox(3) - swTempBody.GetBodyBox(0) > 0.001 Then
                    ReDim Preserve featNames(count): ReDim Preserve coords(count)
                    featNames(count) = swFeat.name
                    
                    ' GET LOCATION COORDINATE
                    Dim idx As Integer: If userDir = "X" Then idx = 3 Else If userDir = "Y" Then idx = 4 Else idx = 5
                    coords(count) = swTempBody.GetBodyBox(idx)
                    
                    count = count + 1
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    If count = 0 Then MsgBox "No valid bodies found.": Exit Sub

    ' 4. SORT
    Dim i As Integer, j As Integer, tempC As Double, tempN As String
    For i = 0 To count - 2: For j = i + 1 To count - 1
        If coords(i) < coords(j) Then
            tempC = coords(i): coords(i) = coords(j): coords(j) = tempC
            tempN = featNames(i): featNames(i) = featNames(j): featNames(j) = tempN
        End If
    Next j: Next i

    ' 5. PROCESS: EXPORT, RENAME & REPORT
    Dim tableData As String
    tableData = "Order" & vbTab & "Description" & vbTab & "L" & vbTab & "W" & vbTab & "T" & vbTab & "Qty" & vbTab & "Perimeter (mm)" & vbTab & "Faces" & vbTab & "Exported?" & vbTab & "Location " & userDir & " (mm)" & vbCrLf
    
    Dim swPart As SldWorks.PartDoc: Set swPart = swModel
    Dim successCount As Integer: successCount = 0
    
    swModel.ClearSelection2 True
    
    For i = 0 To count - 1
        Dim swCurrFeat As SldWorks.Feature: Set swCurrFeat = swModel.FeatureByName(featNames(i))
        Dim swFolderObj As SldWorks.BodyFolder: Set swFolderObj = swCurrFeat.GetSpecificFeature2
        Dim vBods As Variant: vBods = swFolderObj.GetBodies
        
        If Not IsEmpty(vBods) Then
            Dim swBody As SldWorks.Body2: Set swBody = vBods(0)
            
            ' A. CALCULATE DATA
            Dim faceCount As Long: faceCount = swBody.GetFaceCount
            Dim totalPerimeter As Double: totalPerimeter = GetGeometricPerimeter(swBody, normalIndex)
            
            ' B. GET PROPERTIES
            Dim swCustPropMgr As SldWorks.CustomPropertyManager: Set swCustPropMgr = swCurrFeat.CustomPropertyManager
            Dim strDesc As String: strDesc = GetDeepProp(swCustPropMgr, Array("Description", "DESCRIPTION"))
            If strDesc = "-" Then strDesc = swCurrFeat.name
            
            Dim strL As String: strL = GetDeepProp(swCustPropMgr, Array("Length", "LENGTH", "Bounding Box Length"))
            Dim strW As String: strW = GetDeepProp(swCustPropMgr, Array("Width", "WIDTH", "Bounding Box Width"))
            Dim strT As String: strT = GetDeepProp(swCustPropMgr, Array("Thickness", "THICKNESS", "Sheet Metal Thickness"))

            Dim bBox As Variant: bBox = swBody.GetBodyBox
            If strL = "-" Then strL = Round((bBox(3) - bBox(0)) * 1000, 1)
            If strW = "-" Then strW = Round((bBox(4) - bBox(1)) * 1000, 1)
            If strT = "-" Then strT = Round((bBox(5) - bBox(2)) * 1000, 1)

            ' C. RENAME FOLDER (New Feature)
            ' Format: 01_Description, 02_Description
            Dim cleanDesc As String: cleanDesc = CleanFileName(strDesc)
            Dim newName As String: newName = Format(i + 1, "00") & "_" & cleanDesc
            
            ' Safety renaming to avoid conflicts
            swCurrFeat.name = "TEMP_RENAME_" & i
            swCurrFeat.name = newName

            ' D. EXPORT DXF
            swBody.HideBody False
            Dim bFaceSelected As Boolean
            bFaceSelected = SelectLargestFaceAligned(swBody, userDir)
            
            Dim exportStatus As String: exportStatus = "No"
            If bFaceSelected Then
                Dim fileName As String: fileName = newName & ".dxf"
                Dim bRet As Boolean
                bRet = swPart.ExportToDWG2(dxfPath & fileName, fullPath, 2, True, vAlignment, False, False, 0, Nothing)
                If bRet Then
                    successCount = successCount + 1
                    exportStatus = "Yes"
                End If
            End If
            
            ' E. BUILD TABLE ROW
            Dim locVal As Double: locVal = Round(coords(i) * 1000, 2)
            tableData = tableData & (i + 1) & vbTab & strDesc & vbTab & strL & vbTab & strW & vbTab & strT & vbTab & swFolderObj.GetBodyCount & vbTab & Round(totalPerimeter, 2) & vbTab & faceCount & vbTab & exportStatus & vbTab & locVal & vbCrLf
        End If
        swModel.ClearSelection2 True
    Next i

    ' Clipboard
    Dim DataObj As Object: Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText tableData: DataObj.PutInClipboard
    
    MsgBox "Completed! " & successCount & " DXFs exported." & vbCrLf & "Folders renamed and data copied to clipboard."
    Shell "explorer.exe " & dxfPath, vbNormalFocus
End Sub

' --- HELPER FUNCTIONS ---

Function SelectLargestFaceAligned(body As SldWorks.Body2, dirAxis As String) As Boolean
    Dim swFace As SldWorks.Face2: Set swFace = body.GetFirstFace
    Dim bestFace As SldWorks.Face2
    Dim maxArea As Double: maxArea = -1
    
    Do While Not swFace Is Nothing
        Dim swSurf As SldWorks.Surface: Set swSurf = swFace.GetSurface
        If swSurf.IsPlane Then
            Dim vParams As Variant: vParams = swSurf.PlaneParams
            Dim matches As Boolean: matches = False
            Select Case dirAxis
                Case "X": If Abs(vParams(0)) > 0.95 Then matches = True
                Case "Y": If Abs(vParams(1)) > 0.95 Then matches = True
                Case "Z": If Abs(vParams(2)) > 0.95 Then matches = True
            End Select
            
            If matches Then
                Dim area As Double: area = swFace.GetArea
                If area > maxArea Then
                    maxArea = area
                    Set bestFace = swFace
                End If
            End If
        End If
        Set swFace = swFace.GetNextFace
    Loop
    
    If Not bestFace Is Nothing Then
        bestFace.Select4 False, Nothing
        SelectLargestFaceAligned = True
    Else
        SelectLargestFaceAligned = False
    End If
End Function

Function GetGeometricPerimeter(swBody As SldWorks.Body2, normalIdx As Integer) As Double
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
        If Abs(vNormal(normalIdx)) > 0.99 Then
            vLoops = swFace.GetLoops
            If Not IsEmpty(vLoops) Then
                For i = 0 To UBound(vLoops)
                    Set swLoop = vLoops(i)
                    vEdges = swLoop.GetEdges
                    For j = 0 To UBound(vEdges)
                        Set swEdge = vEdges(j)
                        Set swCurve = swEdge.GetCurve
                        vParams = swEdge.GetCurveParams2
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
