Option Explicit

Public Sub RefreshZoneColours()
    Dim progressSheet As Worksheet
    Dim planSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set progressSheet = ThisWorkbook.Worksheets("Progress")
    lastRow = progressSheet.Cells(progressSheet.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        Dim roomId As String
        Dim planSheetName As String
        Dim shapeName As String
        Dim labelShapeName As String
        Dim percentValue As Double
        Dim opacityValue As Double

        roomId = CStr(progressSheet.Cells(rowIndex, "A").Value)
        planSheetName = CStr(progressSheet.Cells(rowIndex, "B").Value)
        shapeName = CStr(progressSheet.Cells(rowIndex, "C").Value)
        percentValue = CDbl(Val(Replace(CStr(progressSheet.Cells(rowIndex, "E").Value), "%", "")))
        opacityValue = CDbl(Val(Replace(CStr(progressSheet.Cells(rowIndex, "F").Value), "%", ""))) / 100
        labelShapeName = CStr(progressSheet.Cells(rowIndex, "H").Value)

        If Len(planSheetName) = 0 Then planSheetName = "Plan"
        Set planSheet = Nothing
        On Error Resume Next
        Set planSheet = ThisWorkbook.Worksheets(planSheetName)
        On Error GoTo 0
        If planSheet Is Nothing Then GoTo NextProgressRow

        If opacityValue < 0 Then opacityValue = 0
        If opacityValue > 1 Then opacityValue = 1

        If Len(shapeName) > 0 Then
            On Error Resume Next
            With planSheet.Shapes(shapeName)
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = ProgressColour(percentValue)
                .Fill.Transparency = 1 - opacityValue
            End With
            On Error GoTo 0
        End If

        If Len(labelShapeName) > 0 Then
            On Error Resume Next
            With planSheet.Shapes(labelShapeName)
                UpdateZoneLabel .TextFrame2, .Width, .Height, roomId, percentValue
            End With
            On Error GoTo 0
        End If

NextProgressRow:
    Next rowIndex
End Sub

Private Sub UpdateZoneLabel(ByVal frame As TextFrame2, ByVal shapeWidth As Double, ByVal shapeHeight As Double, ByVal roomId As String, ByVal percentValue As Double)
    Dim percentText As String
    Dim labelText As String
    Dim percentFontSize As Double
    Dim idFontSize As Double
    Dim idStart As Long

    percentText = CStr(Round(percentValue, 0)) & "%"
    labelText = percentText & vbCrLf & roomId
    percentFontSize = LabelLineFontSize(shapeWidth, shapeHeight, percentText, 18)
    idFontSize = LabelLineFontSize(shapeWidth, shapeHeight, roomId, 14)
    If percentFontSize * 0.78 < idFontSize Then idFontSize = percentFontSize * 0.78
    idStart = Len(percentText) + Len(vbCrLf) + 1

    With frame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .TextRange.Text = labelText
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Size = idFontSize
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        If Len(percentText) > 0 Then .TextRange.Characters(1, Len(percentText)).Font.Size = percentFontSize
        If Len(roomId) > 0 Then .TextRange.Characters(idStart, Len(roomId)).Font.Size = idFontSize
    End With
End Sub

Private Function LabelLineFontSize(ByVal shapeWidth As Double, ByVal shapeHeight As Double, ByVal labelText As String, ByVal maxSize As Double) As Double
    Dim sizeValue As Double
    Dim widthSize As Double
    Dim heightSize As Double
    Dim textLength As Long

    textLength = Len(labelText)
    If textLength < 1 Then textLength = 1
    widthSize = shapeWidth / (textLength * 0.68)
    heightSize = shapeHeight / 3.1
    sizeValue = widthSize
    If heightSize < sizeValue Then sizeValue = heightSize
    If sizeValue < 5 Then sizeValue = 5
    If sizeValue > maxSize Then sizeValue = maxSize

    LabelLineFontSize = sizeValue
End Function

Private Function ProgressColour(ByVal percentValue As Double) As Long
    Dim pct As Double
    Dim amount As Double
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    pct = percentValue
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    If pct <= 50 Then
        amount = pct / 50
        redValue = MixChannel(216, 217, amount)
        greenValue = MixChannel(66, 163, amount)
        blueValue = MixChannel(47, 33, amount)
    Else
        amount = (pct - 50) / 50
        redValue = MixChannel(217, 8, amount)
        greenValue = MixChannel(163, 88, amount)
        blueValue = MixChannel(33, 43, amount)
    End If

    ProgressColour = RGB(redValue, greenValue, blueValue)
End Function

Private Function MixChannel(ByVal startValue As Long, ByVal endValue As Long, ByVal amount As Double) As Long
    MixChannel = CLng(startValue + ((endValue - startValue) * amount))
End Function
