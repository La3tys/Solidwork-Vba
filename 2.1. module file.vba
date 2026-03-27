' Global variables to pass data to the UserForm
Public featNames() As String
Public cleanDescs() As String
Public count As Integer
Public logData As String
Public copyData As String
Public swModel As SldWorks.ModelDoc2

Sub main()
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    
    Dim userDir As String
    userDir = UCase(InputBox("Select axis for sorting sequence (X, Y, or Z):", "Sort Order", "Z"))
    If userDir <> "X" And userDir <> "Y" And userDir <> "Z" Then Exit Sub
    
    Dim fullPath As String
    fullPath = swModel.GetPathName
    If fullPath = "" Then
        MsgBox "Save part first."
        Exit Sub
    End If
    
    Dim basePath As String
    basePath = Left(fullPath, InStrRev(fullPath, "\"))
    Dim partTitle As String
    partTitle = swModel.GetTitle
    If InStrRev(partTitle, ".") > 0 Then partTitle = Left(partTitle, InStrRev(partTitle, ".") - 1)
    
    Dim dxfPath As String
    dxfPath = basePath & partTitle & "_dxfexport\"
    If Dir(dxfPath, vbDirectory) = "" Then MkDir dxfPath

    ' =====================================================================
    ' FIX: Safe Rebuild. Syncs the API to your manual drag-and-drops
    ' without destroying your custom folders!
    swModel.EditRebuild3
    ' =====================================================================
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    count = 0
    
    Dim bBox As Variant
    Dim sortCoords() As Double
    
    ' 1. GATHER ALL CUT-LISTS
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "SolidBodyFolder" Then
            Dim swSubFeat As SldWorks.Feature
            Set swSubFeat = swFeat.GetFirstSubFeature
            
            Do While Not swSubFeat Is Nothing
                
                ' WILDCARD FIX: Grabs literally any folder inside the Solid Bodies list
                If swSubFeat.GetTypeName2 Like "*Folder*" Then
                    
                    If InStr(1, swSubFeat.Name, "Exclude", vbTextCompare) = 0 Then
                        
                        Dim swFolder As SldWorks.BodyFolder
                        Set swFolder = swSubFeat.GetSpecificFeature2
                        
                        If Not swFolder Is Nothing Then
                            Dim vBodies As Variant
                            vBodies = swFolder.GetBodies
                            
                            If Not IsEmpty(vBodies) Then
                                Dim swTempBody As SldWorks.Body2
                                Set swTempBody = vBodies(0)
                                bBox = swTempBody.GetBodyBox
                                
                                If Not IsEmpty(bBox) Then
                                    ReDim Preserve featNames(count)
                                    ReDim Preserve sortCoords(count)
                                    featNames(count) = swSubFeat.Name
                                    
                                    Dim centerVal As Double
                                    If userDir = "X" Then
                                        centerVal = (bBox(0) + bBox(3)) / 2
                                    ElseIf userDir = "Y" Then
                                        centerVal = (bBox(1) + bBox(4)) / 2
                                    Else
                                        centerVal = (bBox(2) + bBox(5)) / 2
                                    End If
                                    
                                    sortCoords(count) = centerVal
                                    count = count + 1
                                End If
                            End If
                        End If
                        
                    End If
                    
                End If
                Set swSubFeat = swSubFeat.GetNextFeature
            Loop
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    If count = 0 Then
        MsgBox "No valid bodies found. Check if they are grouped inside the Solid Bodies folder."
        Exit Sub
    End If

    ' 2. SORT THEM BY COORDINATES
    Dim i As Integer, j As Integer
    Dim tempC As Double, tempN As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            Dim doSwap As Boolean
            doSwap = False
            
            If userDir = "X" Then
                If sortCoords(i) > sortCoords(j) Then doSwap = True
            Else
                If sortCoords(i) < sortCoords(j) Then doSwap = True
            End If
            
            If doSwap Then
                tempC = sortCoords(i)
                sortCoords(i) = sortCoords(j)
                sortCoords(j) = tempC
                
                tempN = featNames(i)
                featNames(i) = featNames(j)
                featNames(j) = tempN
            End If
        Next j
    Next i

    ' 3. PREP DATA STRINGS
    copyData = "Item" & vbTab & "Qty" & vbTab & "Description" & vbTab & "Material" & vbTab & "Pos." & vbTab & "Length" & vbTab & "Width" & vbTab & "Height" & vbTab & "Perimeter" & vbTab & "Faces" & vbCrLf
    logData = ""
    ReDim cleanDescs(count - 1)
    
    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel
    swModel.ClearSelection2 True
    
    ' 4. EXPORT AND EXTRACT DATA
    For i = 0 To count - 1
        Dim valL As Double: valL = 0
        Dim valW As Double: valW = 0
        Dim valH As Double: valH = 0
        Dim totalPerimeter As Double: totalPerimeter = 0
        Dim faceCount As Long: faceCount = 0
        Dim strDesc As String: strDesc = ""
        Dim strMaterial As String: strMaterial = ""
        Dim strPos As String: strPos = ""

        Dim swCurrFeat As SldWorks.Feature
        Set swCurrFeat = swModel.FeatureByName(featNames(i))
        Dim swFolderObj As SldWorks.BodyFolder
        Set swFolderObj = swCurrFeat.GetSpecificFeature2
        vBodies = swFolderObj.GetBodies
        
        ' =====================================================================
        ' FIX: True Quantity Counting!
        ' Manually counts the physical bodies inside the folder instead of trusting the API
        Dim bodyQty As Integer
        If Not IsEmpty(vBodies) Then
            bodyQty = UBound(vBodies) - LBound(vBodies) + 1
        Else
            bodyQty = 0
        End If
        ' =====================================================================
        
        If Not IsEmpty(vBodies) Then
            Dim swExportBody As SldWorks.Body2
            Set swExportBody = vBodies(0)
            faceCount = swExportBody.GetFaceCount - 2
            
            bBox = swExportBody.GetBodyBox
            If Not IsEmpty(bBox) Then
                Dim dArr(2) As Double
                dArr(0) = bBox(3) - bBox(0)
                dArr(1) = bBox(4) - bBox(1)
                dArr(2) = bBox(5) - bBox(2)
                
                Dim x As Integer, y As Integer, temp As Double
                For x = 0 To 1
                    For y = x + 1 To 2
                        If dArr(x) > dArr(y) Then
                            temp = dArr(x)
                            dArr(x) = dArr(y)
                            dArr(y) = temp
                        End If
                    Next y
                Next x
                
                valH = Round(dArr(0) * 1000, 2)
                valW = Round(dArr(1) * 1000, 2)
                valL = Round(dArr(2) * 1000, 2)
            End If
            
            Dim folderName As String
            folderName = swCurrFeat.Name
            If InStr(folderName, "<") > 0 Then
                folderName = Trim(Left(folderName, InStr(folderName, "<") - 1))
            End If
            
            strDesc = folderName
            
            Dim descParts() As String
            descParts = Split(folderName, ",")
            
            If UBound(descParts) > 0 Then
                strPos = Trim(descParts(0))
                cleanDescs(i) = Trim(descParts(UBound(descParts)))
            Else
                strPos = "-"
                cleanDescs(i) = folderName
            End If
            
            Dim swCustPropMgr As SldWorks.CustomPropertyManager
            Set swCustPropMgr = swCurrFeat.CustomPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                strMaterial = GetDeepProp(swCustPropMgr, Array("Material", "MATERIAL"))
                If strMaterial = "-" Then strMaterial = "Unknown"
            Else
                strMaterial = "Unknown"
            End If
            
            swExportBody.HideBody False
            Dim detectedAxis As String
            Dim bFaceFound As Boolean
            bFaceFound = SelectLargestFaceAndGetNormal(swExportBody, detectedAxis)
            
            If bFaceFound Then
                Dim fileName As String
                fileName = CleanFileName(strDesc) & ".dxf"
                Dim vAlign As Variant
                vAlign = GetMatrixForAxis(detectedAxis)
                
                swPart.ExportToDWG2 dxfPath & fileName, fullPath, 2, True, vAlign, False, False, 0, Nothing
                
                Dim normIdx As Integer
                If detectedAxis = "X" Then
                    normIdx = 0
                ElseIf detectedAxis = "Y" Then
                    normIdx = 1
                Else
                    normIdx = 2
                End If
                
                totalPerimeter = GetGeometricPerimeter(swExportBody, normIdx)
            End If
            
            Dim rowStr As String
            rowStr = (i + 1) & vbTab & bodyQty & vbTab & strDesc & vbTab & strMaterial & vbTab & strPos & vbTab & valL & vbTab & valW & vbTab & valH & vbTab & Round(totalPerimeter, 2) & vbTab & faceCount
            logData = logData & rowStr & vbCrLf
            copyData = copyData & rowStr & vbCrLf
        End If
        swModel.ClearSelection2 True
    Next i
    
    Shell "explorer.exe """ & dxfPath & """", vbNormalFocus
    UserForm1.Show
End Sub

' --- HELPER FUNCTIONS ---
Function SelectLargestFaceAndGetNormal(body As SldWorks.Body2, ByRef axisOut As String) As Boolean
    Dim swFace As SldWorks.Face2
    Set swFace = body.GetFirstFace
    Dim bestFace As SldWorks.Face2
    Dim maxArea As Double
    maxArea = -1
    Dim bestNormal As Variant
    
    Do While Not swFace Is Nothing
        Dim area As Double
        area = swFace.GetArea
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
        Case "X"
            vData(5) = 1: vData(7) = 1: vData(9) = 1
        Case "Y"
            vData(3) = 1: vData(8) = 1: vData(10) = 1
        Case "Z"
            vData(3) = 1: vData(7) = 1: vData(11) = 1
    End Select
    GetMatrixForAxis = vData
End Function

Function GetGeometricPerimeter(swBody As SldWorks.Body2, normalIdx As Integer) As Double
    Dim swFace As SldWorks.Face2
    Set swFace = swBody.GetFirstFace
    Do While Not swFace Is Nothing
        Dim vNormal As Variant
        vNormal = swFace.Normal
        If Abs(vNormal(normalIdx)) > 0.99 Then
            Dim vLoops As Variant
            vLoops = swFace.GetLoops
            If Not IsEmpty(vLoops) Then
                Dim i As Integer
                Dim totalLength As Double
                totalLength = 0
                For i = 0 To UBound(vLoops)
                    Dim swLoop As SldWorks.Loop2
                    Set swLoop = vLoops(i)
                    Dim vEdges As Variant
                    vEdges = swLoop.GetEdges
                    Dim j As Integer
                    For j = 0 To UBound(vEdges)
                        Dim swEdge As SldWorks.Edge
                        Set swEdge = vEdges(j)
                        Dim swCurve As SldWorks.Curve
                        Set swCurve = swEdge.GetCurve
                        Dim vParams As Variant
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
    Dim i As Integer
    Dim val As String, res As String
    Dim b As Boolean
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
    Dim invalidChars As String
    invalidChars = "\/:*?""<>|"
    Dim i As Integer
    CleanFileName = Trim(strIn)
    For i = 1 To Len(invalidChars)
        CleanFileName = Replace(CleanFileName, Mid(invalidChars, i, 1), "_")
    Next i
End Function
