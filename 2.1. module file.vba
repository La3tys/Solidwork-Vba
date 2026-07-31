Option Explicit

Public featNames() As String
Public cleanDescs() As String
Public count As Long
Public logData As String
Public copyData As String
Public swModel As SldWorks.ModelDoc2
Public selectedAxis As String

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim fullPath As String
    Dim swFeat As SldWorks.Feature
    Dim swSubFeat As SldWorks.Feature
    Dim swFolder As SldWorks.BodyFolder
    Dim vBodies As Variant
    Dim swTempBody As SldWorks.Body2
    Dim bBox As Variant
    Dim sortCoords() As Double
    Dim i As Long, j As Long
    Dim tempC As Double, tempN As String

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocPART Then
        MsgBox "Open a part document first.", vbExclamation
        Exit Sub
    End If

    selectedAxis = "Z"
    fullPath = swModel.GetPathName
    If fullPath = "" Then
        MsgBox "Save part first.", vbExclamation
        Exit Sub
    End If

    swModel.EditRebuild3
    count = 0
    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "SolidBodyFolder" Then
            Set swSubFeat = swFeat.GetFirstSubFeature

            Do While Not swSubFeat Is Nothing
                If IsIncludedCutListFeature(swSubFeat) Then
                    Set swFolder = swSubFeat.GetSpecificFeature2

                    If Not swFolder Is Nothing Then
                        vBodies = swFolder.GetBodies

                        If Not IsEmpty(vBodies) Then
                            Set swTempBody = vBodies(LBound(vBodies))
                            bBox = swTempBody.GetBodyBox

                            If Not IsEmpty(bBox) Then
                                ReDim Preserve featNames(count)
                                ReDim Preserve sortCoords(count)
                                featNames(count) = swSubFeat.Name
                                sortCoords(count) = (bBox(2) + bBox(5)) / 2#
                                count = count + 1
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
        MsgBox "No included cut-list folders were found.", vbExclamation
        Exit Sub
    End If

    ' Initial display order is Z. The UserForm can later sort by X, Y, or Z.
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If sortCoords(i) < sortCoords(j) Then
                tempC = sortCoords(i): sortCoords(i) = sortCoords(j): sortCoords(j) = tempC
                tempN = featNames(i): featNames(i) = featNames(j): featNames(j) = tempN
            End If
        Next j
    Next i

    ReDim cleanDescs(0 To count - 1)
    copyData = "Item" & vbTab & "Description" & vbTab & "Material" & vbTab & "Pos." & vbTab & "Qty" & vbTab & "Length" & vbTab & "Width" & vbTab & "Height" & vbTab & "Cut-length(Perimeter)" & vbTab & "Transition(Faces)" & vbCrLf
    logData = ""

    For i = 0 To count - 1
        AppendCutListLogRow i
    Next i

    copyData = copyData & logData
    UserForm1.Show
End Sub

Public Function IsIncludedCutListFeature(ByVal swFeat As SldWorks.Feature) As Boolean
    If swFeat Is Nothing Then Exit Function
    IsIncludedCutListFeature = (swFeat.GetTypeName2 = "CutListFolder" And Not swFeat.ExcludeFromCutList)
End Function

Public Function AxisIndex(ByVal axis As String) As Integer
    Select Case UCase$(axis)
        Case "X": AxisIndex = 0
        Case "Y": AxisIndex = 1
        Case Else: AxisIndex = 2
    End Select
End Function

Public Sub RemoveExcludedCutListFeatures()
    Dim i As Long, writeIndex As Long
    Dim swFeat As SldWorks.Feature

    writeIndex = 0
    For i = 0 To count - 1
        Set swFeat = swModel.FeatureByName(featNames(i))

        If IsIncludedCutListFeature(swFeat) Then
            If writeIndex <> i Then
                featNames(writeIndex) = featNames(i)
                cleanDescs(writeIndex) = cleanDescs(i)
            End If
            writeIndex = writeIndex + 1
        End If
    Next i

    count = writeIndex
    If count = 0 Then
        Erase featNames
        Erase cleanDescs
    Else
        ReDim Preserve featNames(0 To count - 1)
        ReDim Preserve cleanDescs(0 To count - 1)
    End If
End Sub

Public Sub SortCutListByAxis(ByVal axis As String)
    Dim coords() As Double
    Dim i As Long, j As Long, axisIdx As Integer
    Dim swFeat As SldWorks.Feature
    Dim swFolder As SldWorks.BodyFolder
    Dim vBodies As Variant, bBox As Variant
    Dim tempCoord As Double, tempName As String, tempDesc As String

    If count < 2 Then Exit Sub
    axisIdx = AxisIndex(axis)
    ReDim coords(0 To count - 1)

    For i = 0 To count - 1
        Set swFeat = swModel.FeatureByName(featNames(i))
        If IsIncludedCutListFeature(swFeat) Then
            Set swFolder = swFeat.GetSpecificFeature2
            vBodies = swFolder.GetBodies
            If Not IsEmpty(vBodies) Then
                bBox = vBodies(LBound(vBodies)).GetBodyBox
                coords(i) = (bBox(axisIdx) + bBox(axisIdx + 3)) / 2#
            End If
        End If
    Next i

    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If coords(i) < coords(j) Then
                tempCoord = coords(i): coords(i) = coords(j): coords(j) = tempCoord
                tempName = featNames(i): featNames(i) = featNames(j): featNames(j) = tempName
                tempDesc = cleanDescs(i): cleanDescs(i) = cleanDescs(j): cleanDescs(j) = tempDesc
            End If
        Next j
    Next i
End Sub

Public Sub AppendCutListLogRow(ByVal itemIndex As Long)
    Dim swFeat As SldWorks.Feature
    Dim swFolder As SldWorks.BodyFolder
    Dim vBodies As Variant, bBox As Variant
    Dim swBody As SldWorks.Body2
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim bodyQty As Long, valL As Double, valW As Double, valH As Double
    Dim totalPerimeter As Double, contourSegments As Long, detectedAxis As String
    Dim dArr(2) As Double, x As Integer, y As Integer, temp As Double
    Dim folderName As String, strDesc As String, strMaterial As String, strPos As String
    Dim descParts() As String, rowStr As String

    Set swFeat = swModel.FeatureByName(featNames(itemIndex))
    If Not IsIncludedCutListFeature(swFeat) Then Exit Sub

    Set swFolder = swFeat.GetSpecificFeature2
    vBodies = swFolder.GetBodies
    If IsEmpty(vBodies) Then Exit Sub

    bodyQty = UBound(vBodies) - LBound(vBodies) + 1
    Set swBody = vBodies(LBound(vBodies))
    bBox = swBody.GetBodyBox

    If Not IsEmpty(bBox) Then
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
    End If

    folderName = swFeat.Name
    If InStr(folderName, "<") > 0 Then folderName = Trim$(Left$(folderName, InStr(folderName, "<") - 1))
    strDesc = folderName
    descParts = Split(folderName, ",")
    If UBound(descParts) > 0 Then
        strPos = Trim$(descParts(0))
        cleanDescs(itemIndex) = Trim$(descParts(UBound(descParts)))
    Else
        strPos = "-"
        cleanDescs(itemIndex) = folderName
    End If

    Set swCustPropMgr = swFeat.CustomPropertyManager
    If Not swCustPropMgr Is Nothing Then
        strMaterial = GetDeepProp(swCustPropMgr, Array("Material", "MATERIAL"))
        If strMaterial = "-" Then strMaterial = "Unknown"
    Else
        strMaterial = "Unknown"
    End If

    GetLargestFaceMetrics swBody, detectedAxis, totalPerimeter, contourSegments

    rowStr = (itemIndex + 1) & vbTab & strDesc & vbTab & strMaterial & vbTab & strPos & vbTab & bodyQty & vbTab & valL & vbTab & valW & vbTab & valH & vbTab & Round(totalPerimeter, 2) & vbTab & contourSegments
    logData = logData & rowStr & vbCrLf
End Sub

Public Function GetLargestFaceMetrics(ByVal swBody As SldWorks.Body2, ByRef axisOut As String, ByRef perimeterOut As Double, ByRef segmentCountOut As Long) As Boolean
    Dim swFace As SldWorks.Face2, swLargestFace As SldWorks.Face2
    Dim swLoop As SldWorks.Loop2, swEdge As SldWorks.Edge, swCurve As SldWorks.Curve
    Dim vLoops As Variant, vEdges As Variant, vParams As Variant, vNormal As Variant
    Dim maxArea As Double, i As Long, j As Long

    perimeterOut = 0
    segmentCountOut = 0
    maxArea = -1
    Set swFace = swBody.GetFirstFace

    Do While Not swFace Is Nothing
        If swFace.GetArea > maxArea Then
            maxArea = swFace.GetArea
            Set swLargestFace = swFace
        End If
        Set swFace = swFace.GetNextFace
    Loop

    If swLargestFace Is Nothing Then Exit Function
    vNormal = swLargestFace.Normal
    If Abs(vNormal(0)) > 0.9 Then
        axisOut = "X"
    ElseIf Abs(vNormal(1)) > 0.9 Then
        axisOut = "Y"
    Else
        axisOut = "Z"
    End If

    vLoops = swLargestFace.GetLoops
    If IsEmpty(vLoops) Then Exit Function
    For i = LBound(vLoops) To UBound(vLoops)
        Set swLoop = vLoops(i)
        vEdges = swLoop.GetEdges
        If Not IsEmpty(vEdges) Then
            For j = LBound(vEdges) To UBound(vEdges)
                Set swEdge = vEdges(j)
                Set swCurve = swEdge.GetCurve
                vParams = swEdge.GetCurveParams2
                perimeterOut = perimeterOut + swCurve.GetLength3(vParams(6), vParams(7)) * 1000#
                segmentCountOut = segmentCountOut + 1
            Next j
        End If
    Next i

    GetLargestFaceMetrics = True
End Function

Public Function SelectLargestFaceAndGetNormal(body As SldWorks.Body2, ByRef axisOut As String) As Boolean
    Dim swFace As SldWorks.Face2
    Dim bestFace As SldWorks.Face2
    Dim maxArea As Double
    Dim bestNormal As Variant

    maxArea = -1
    Set swFace = body.GetFirstFace

    Do While Not swFace Is Nothing
        If swFace.GetArea > maxArea Then
            maxArea = swFace.GetArea
            Set bestFace = swFace
            bestNormal = swFace.Normal
        End If
        Set swFace = swFace.GetNextFace
    Loop

    If bestFace Is Nothing Then Exit Function

    ' ExportToDWG2 uses the selected broad sheet face as its export reference.
    bestFace.Select4 False, Nothing

    If Abs(bestNormal(0)) > 0.9 Then
        axisOut = "X"
    ElseIf Abs(bestNormal(1)) > 0.9 Then
        axisOut = "Y"
    Else
        axisOut = "Z"
    End If

    SelectLargestFaceAndGetNormal = True
End Function

Public Function GetMatrixForAxis(axis As String) As Variant
    Dim vData(11) As Double
    Select Case axis
        Case "X": vData(5) = 1: vData(7) = 1: vData(9) = 1
        Case "Y": vData(3) = 1: vData(8) = 1: vData(10) = 1
        Case "Z": vData(3) = 1: vData(7) = 1: vData(11) = 1
    End Select
    GetMatrixForAxis = vData
End Function

Public Function GetDeepProp(mgr As SldWorks.CustomPropertyManager, names As Variant) As String
    Dim i As Long, val As String, res As String, b As Boolean
    For i = LBound(names) To UBound(names)
        mgr.Get6 CStr(names(i)), False, val, res, b, False
        If res <> "" And Not res Like "*@*" Then
            GetDeepProp = Trim$(Replace(res, "mm", ""))
            Exit Function
        End If
    Next i
    GetDeepProp = "-"
End Function

Public Function CleanFileName(strIn As String) As String
    Dim invalidChars As String, i As Long
    invalidChars = "\/:*?""<>|"
    CleanFileName = Trim$(strIn)
    For i = 1 To Len(invalidChars)
        CleanFileName = Replace(CleanFileName, Mid$(invalidChars, i, 1), "_")
    Next i
End Function

Public Function FormatDXFFileName(strPos As String, fullDesc As String, valL As Double, valW As Double, valH As Double, bodyQty As Long) As String
    Dim dimStr As String, qtyStr As String, paddedDim As String
    If strPos <> "-" And strPos <> "" Then
        dimStr = strPos & " - " & fullDesc
    Else
        dimStr = fullDesc
    End If
    If InStr(1, fullDesc, " x ", vbTextCompare) = 0 Then dimStr = dimStr & " - " & valL & " x " & valW & " x " & valH
    qtyStr = "(" & bodyQty & "x)"
    If Len(dimStr) < 30 Then paddedDim = dimStr & Space$(30 - Len(dimStr)) Else paddedDim = dimStr
    FormatDXFFileName = CleanFileName(paddedDim) & " " & qtyStr
End Function

Public Function DecimalToRoman(ByVal n As Long) As String
    Dim ro As String, vals As Variant, roms As Variant, i As Long
    vals = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    roms = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    For i = 0 To UBound(vals)
        Do While n >= vals(i)
            n = n - vals(i): ro = ro & roms(i)
        Loop
    Next i
    DecimalToRoman = ro
End Function

Public Function NumberToLetter(ByVal n As Long) As String
    Dim s As String, remVal As Long
    Do While n > 0
        remVal = (n - 1) Mod 26
        s = Chr$(65 + remVal) & s
        n = Int((n - remVal) / 26)
    Loop
    NumberToLetter = s
End Function
