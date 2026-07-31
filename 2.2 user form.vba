Option Explicit

Private isInitializing As Boolean

Private Sub UserForm_Initialize()
    Dim rows() As String, cols() As String
    Dim i As Long, r As Long

    isInitializing = True
    Me.Caption = "Cut-List Extraction Log"
    lstLog.ColumnCount = 10
    lstLog.ColumnWidths = "30 pt;150 pt;80 pt;40 pt;30 pt;60 pt;60 pt;60 pt;70 pt;80 pt"
    SetLogHeaders

    rows = Split(logData, vbCrLf)
    r = 1
    For i = 0 To UBound(rows)
        If Trim$(rows(i)) <> "" Then
            cols = Split(rows(i), vbTab)
            If UBound(cols) >= 9 Then
                AddLogRow cols, r
                r = r + 1
            End If
        End If
    Next i

    cboStyle.AddItem "1, 2, 3..."
    cboStyle.AddItem "A, B, C..."
    cboStyle.AddItem "I, II, III..."
    cboStyle.ListIndex = 0
    cboDirection.AddItem "X"
    cboDirection.AddItem "Y"
    cboDirection.AddItem "Z"
    cboDirection.ListIndex = 2
    selectedAxis = "Z"
    isInitializing = False
End Sub

Private Sub cboDirection_Change()
    If isInitializing Then Exit Sub
    If cboDirection.ListIndex < 0 Then Exit Sub

    selectedAxis = cboDirection.Value
    RemoveExcludedCutListFeatures
    SortCutListByAxis selectedAxis
    RefreshLogFromCleanDescs
End Sub

Private Sub cmdCopy_Click()
    Dim DataObj As Object
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText copyData
    DataObj.PutInClipboard
    MsgBox "Data copied to clipboard!", vbInformation, "Copied"
End Sub

Private Sub cmdRename_Click()
    Dim style As String, prefix As String, i As Long, renameCount As Long
    Dim indexNum As Long, sequenceStr As String, finalName As String
    Dim swRenameFeat As SldWorks.Feature

    RemoveExcludedCutListFeatures
    SortCutListByAxis cboDirection.Value
    style = cboStyle.Text
    prefix = txtPrefix.Text

    For i = 0 To count - 1
        indexNum = i + 1
        Select Case style
            Case "A, B, C...": sequenceStr = NumberToLetter(indexNum)
            Case "I, II, III...": sequenceStr = DecimalToRoman(indexNum)
            Case Else: sequenceStr = CStr(indexNum)
        End Select

        finalName = prefix & sequenceStr & ", " & cleanDescs(i)
        Set swRenameFeat = swModel.FeatureByName(featNames(i))

        If IsIncludedCutListFeature(swRenameFeat) Then
            If swRenameFeat.Name <> finalName Then
                swRenameFeat.Name = "TEMP_RENAME_" & CStr(i)
                swRenameFeat.Name = finalName
                featNames(i) = finalName
                renameCount = renameCount + 1
            End If
        End If
    Next i

    RefreshLogFromCleanDescs
    MsgBox renameCount & " folders renamed to style: " & style, vbInformation, "Rename Complete"
End Sub

Private Sub cmdExportDXF_Click()
    Dim swApp As SldWorks.SldWorks
    Dim activeModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    Dim fullPath As String, basePath As String, partTitle As String, dxfPath As String
    Dim logRow As Long, sourceIndex As Long, bodyQty As Long
    Dim swCurrFeat As SldWorks.Feature
    Dim swFolderObj As SldWorks.BodyFolder
    Dim vBodies As Variant, swExportBody As SldWorks.Body2, bBox As Variant
    Dim valL As Double, valW As Double, valH As Double
    Dim detectedAxis As String, totalPerimeter As Double, contourSegments As Long
    Dim bFaceFound As Boolean, exportSucceeded As Boolean
    Dim exportedCount As Long, skippedCount As Long
    Dim fileName As String, fullFeatName As String, descForFile As String, posForFile As String
    Dim featParts() As String, vAlign As Variant

    Set swApp = Application.SldWorks
    Set activeModel = swApp.ActiveDoc
    If activeModel Is Nothing Then
        MsgBox "No active document.", vbExclamation
        Exit Sub
    End If

    fullPath = activeModel.GetPathName
    If fullPath = "" Then
        MsgBox "Save part first.", vbExclamation
        Exit Sub
    End If

    basePath = Left$(fullPath, InStrRev(fullPath, "\"))
    partTitle = activeModel.GetTitle
    If InStrRev(partTitle, ".") > 0 Then partTitle = Left$(partTitle, InStrRev(partTitle, ".") - 1)
    dxfPath = basePath & partTitle & "_dxfexport\"
    If Dir(dxfPath, vbDirectory) = "" Then MkDir dxfPath

    Set swModel = activeModel
    Set swPart = activeModel

    ' The list is the authoritative export set. Its Item values retain the
    ' source index when excluded folders are hidden (for example: 1, 6, 8, 9).
    For logRow = 1 To lstLog.ListCount - 1
        If IsNumeric(lstLog.List(logRow, 0)) Then
            sourceIndex = CLng(lstLog.List(logRow, 0)) - 1

            If sourceIndex >= LBound(featNames) And sourceIndex <= UBound(featNames) Then
                Set swCurrFeat = activeModel.FeatureByName(featNames(sourceIndex))

                If IsIncludedCutListFeature(swCurrFeat) Then
                    Set swFolderObj = swCurrFeat.GetSpecificFeature2
                    vBodies = swFolderObj.GetBodies

                    If Not IsEmpty(vBodies) Then
                        bodyQty = UBound(vBodies) - LBound(vBodies) + 1
                        Set swExportBody = vBodies(LBound(vBodies))
                        GetSortedBodyDimensions swExportBody, valL, valW, valH

                        swExportBody.HideBody False
                        activeModel.ClearSelection2 True
                        bFaceFound = SelectLargestFaceAndGetNormal(swExportBody, detectedAxis)

                        If bFaceFound Then
                            fullFeatName = swCurrFeat.Name
                            featParts = Split(fullFeatName, ",")
                            If UBound(featParts) > 0 Then
                                posForFile = Trim$(featParts(0))
                                descForFile = Trim$(featParts(UBound(featParts)))
                            Else
                                posForFile = "-"
                                descForFile = Trim$(featParts(0))
                            End If

                            fileName = FormatDXFFileName(posForFile, descForFile, valL, valW, valH, bodyQty) & ".dxf"
                            vAlign = GetMatrixForAxis(detectedAxis)
                            exportSucceeded = swPart.ExportToDWG2(dxfPath & fileName, fullPath, 2, True, vAlign, False, False, 0, Nothing)

                            If exportSucceeded Then
                                exportedCount = exportedCount + 1
                            Else
                                skippedCount = skippedCount + 1
                            End If
                        Else
                            skippedCount = skippedCount + 1
                        End If
                    Else
                        skippedCount = skippedCount + 1
                    End If
                Else
                    skippedCount = skippedCount + 1
                End If
            Else
                skippedCount = skippedCount + 1
            End If
        End If

        activeModel.ClearSelection2 True
    Next logRow

    MsgBox exportedCount & " DXF file(s) exported." & vbCrLf & _
           skippedCount & " folder(s) failed or were skipped." & vbCrLf & _
           "Folder: " & dxfPath, vbInformation, "DXF Export"
End Sub

Public Sub RefreshLogFromCleanDescs()
    Dim i As Long

    RemoveExcludedCutListFeatures
    logData = ""
    For i = 0 To count - 1
        AppendCutListLogRow i
    Next i

    copyData = "Item" & vbTab & "Description" & vbTab & "Material" & vbTab & "Pos." & vbTab & "Qty" & vbTab & "Length" & vbTab & "Width" & vbTab & "Height" & vbTab & "Cut-length(Perimeter)" & vbTab & "Transition(Faces)" & vbCrLf & logData
    PopulateLog
End Sub

Private Sub PopulateLog()
    Dim rows() As String, cols() As String
    Dim i As Long, r As Long

    lstLog.Clear
    SetLogHeaders
    rows = Split(logData, vbCrLf)
    r = 1

    For i = 0 To UBound(rows)
        If Trim$(rows(i)) <> "" Then
            cols = Split(rows(i), vbTab)
            If UBound(cols) >= 9 Then
                AddLogRow cols, r
                r = r + 1
            End If
        End If
    Next i
End Sub

Private Sub SetLogHeaders()
    lstLog.AddItem "Item"
    lstLog.List(0, 1) = "Description"
    lstLog.List(0, 2) = "Material"
    lstLog.List(0, 3) = "Pos."
    lstLog.List(0, 4) = "Qty"
    lstLog.List(0, 5) = "Length"
    lstLog.List(0, 6) = "Width"
    lstLog.List(0, 7) = "Height"
    lstLog.List(0, 8) = "Cut-Length"
    lstLog.List(0, 9) = "Transition"
End Sub

Private Sub AddLogRow(ByRef cols() As String, ByVal rowIndex As Long)
    lstLog.AddItem cols(0)
    lstLog.List(rowIndex, 1) = cols(1)
    lstLog.List(rowIndex, 2) = cols(2)
    lstLog.List(rowIndex, 3) = cols(3)
    lstLog.List(rowIndex, 4) = cols(4)
    lstLog.List(rowIndex, 5) = cols(5)
    lstLog.List(rowIndex, 6) = cols(6)
    lstLog.List(rowIndex, 7) = cols(7)
    lstLog.List(rowIndex, 8) = cols(8)
    lstLog.List(rowIndex, 9) = cols(9)
End Sub

Private Sub GetSortedBodyDimensions(ByVal swBody As SldWorks.Body2, ByRef valL As Double, ByRef valW As Double, ByRef valH As Double)
    Dim bBox As Variant, dArr(2) As Double
    Dim x As Integer, y As Integer, temp As Double

    bBox = swBody.GetBodyBox
    If IsEmpty(bBox) Then Exit Sub

    dArr(0) = bBox(3) - bBox(0)
    dArr(1) = bBox(4) - bBox(1)
    dArr(2) = bBox(5) - bBox(2)

    For x = 0 To 1
        For y = x + 1 To 2
            If dArr(x) > dArr(y) Then
                temp = dArr(x): dArr(x) = dArr(y): dArr(y) = temp
            End If
        Next y
    Next x

    valH = Round(dArr(0) * 1000#, 2)
    valW = Round(dArr(1) * 1000#, 2)
    valL = Round(dArr(2) * 1000#, 2)
End Sub
