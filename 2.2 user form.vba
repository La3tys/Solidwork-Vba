Private Sub UserForm_Initialize()
    Me.Caption = "Cut-List Extraction Log"
    
    ' Set up ListBox for 10 Columns now
    lstLog.ColumnCount = 10
    ' Added a 40pt width slot for the Pos. column
    lstLog.ColumnWidths = "30 pt;30 pt;150 pt;80 pt;40 pt;50 pt;50 pt;50 pt;60 pt;40 pt"
    
    ' Add Headers
    lstLog.AddItem "Item"
    lstLog.List(0, 1) = "Qty"
    lstLog.List(0, 2) = "Description"
    lstLog.List(0, 3) = "Material"
    lstLog.List(0, 4) = "Pos."         '<-- NEW COLUMN
    lstLog.List(0, 5) = "Length"
    lstLog.List(0, 6) = "Width"
    lstLog.List(0, 7) = "Height"
    lstLog.List(0, 8) = "Perimeter"
    lstLog.List(0, 9) = "Faces"
    
    ' Populate Grid
    Dim rows() As String, cols() As String
    Dim i As Integer, r As Integer
    rows = Split(logData, vbCrLf)
    r = 1
    For i = 0 To UBound(rows)
        If Trim(rows(i)) <> "" Then
            cols = Split(rows(i), vbTab)
            ' Check for 9 or more tabs (10 columns)
            If UBound(cols) >= 9 Then
                lstLog.AddItem cols(0)
                lstLog.List(r, 1) = cols(1): lstLog.List(r, 2) = cols(2)
                lstLog.List(r, 3) = cols(3): lstLog.List(r, 4) = cols(4)
                lstLog.List(r, 5) = cols(5): lstLog.List(r, 6) = cols(6)
                lstLog.List(r, 7) = cols(7): lstLog.List(r, 8) = cols(8)
                lstLog.List(r, 9) = cols(9)
                r = r + 1
            End If
        End If
    Next i
    
    ' Set up Naming Styles Dropdown
    cboStyle.AddItem "1, 2, 3..."
    cboStyle.AddItem "A, B, C..."
    cboStyle.AddItem "I, II, III..."
    cboStyle.ListIndex = 0 ' Default to numbers
End Sub

Private Sub cmdCopy_Click()
    Dim DataObj As Object
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText copyData
    DataObj.PutInClipboard
    MsgBox "Data copied to clipboard!", vbInformation, "Copied"
End Sub

Private Sub cmdRename_Click()
    Dim style As String: style = cboStyle.Text
    Dim prefix As String: prefix = txtPrefix.Text
    Dim i As Integer
    Dim renameCount As Integer: renameCount = 0
    
    For i = 0 To count - 1
        Dim indexNum As Integer: indexNum = i + 1
        Dim sequenceStr As String
        
        ' 1. Determine the sequence format
        Select Case style
            Case "A, B, C...": sequenceStr = NumberToLetter(indexNum)
            Case "I, II, III...": sequenceStr = DecimalToRoman(indexNum)
            Case Else: sequenceStr = CStr(indexNum)
        End Select
        
        ' 2. Combine Prefix + Sequence + Clean Description
        Dim finalName As String
        finalName = prefix & sequenceStr & ", " & cleanDescs(i)
        
        ' 3. Apply to SolidWorks Feature Tree
        Dim swRenameFeat As Object
        Set swRenameFeat = swModel.FeatureByName(featNames(i))
        
        If Not swRenameFeat Is Nothing Then
            If swRenameFeat.Name <> finalName Then
                swRenameFeat.Name = "TEMP_RENAME_" & i
                swRenameFeat.Name = finalName
                renameCount = renameCount + 1
            End If
        End If
    Next i
    
    MsgBox renameCount & " folders renamed to style: " & style, vbInformation, "Rename Complete"
End Sub

' --- FORMATTING HELPER FUNCTIONS ---
Function DecimalToRoman(ByVal n As Integer) As String
    Dim ro As String: ro = ""
    Dim vals As Variant: vals = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    Dim roms As Variant: roms = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    Dim i As Integer
    For i = 0 To UBound(vals)
        Do While n >= vals(i)
            n = n - vals(i)
            ro = ro & roms(i)
        Loop
    Next i
    DecimalToRoman = ro
End Function

Function NumberToLetter(ByVal n As Integer) As String
    Dim s As String: s = ""
    Do While n > 0
        Dim remVal As Integer
        remVal = (n - 1) Mod 26
        s = Chr(65 + remVal) & s
        n = Int((n - remVal) / 26)
    Loop
    NumberToLetter = s
End Function
