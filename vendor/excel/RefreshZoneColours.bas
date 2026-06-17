Option Explicit

Public Sub RefreshZoneColours()
    Dim progressSheet As Worksheet
    Dim planSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set progressSheet = ThisWorkbook.Worksheets("Progress")
    Set planSheet = ThisWorkbook.Worksheets("Plan")
    lastRow = progressSheet.Cells(progressSheet.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        Dim roomId As String
        Dim shapeName As String
        Dim labelShapeName As String
        Dim percentValue As Double
        Dim opacityValue As Double

        roomId = CStr(progressSheet.Cells(rowIndex, "A").Value)
        shapeName = CStr(progressSheet.Cells(rowIndex, "B").Value)
        percentValue = CDbl(Val(progressSheet.Cells(rowIndex, "D").Value))
        opacityValue = CDbl(Val(Replace(CStr(progressSheet.Cells(rowIndex, "E").Value), "%", ""))) / 100
        labelShapeName = CStr(progressSheet.Cells(rowIndex, "G").Value)

        If Len(shapeName) > 0 Then
            On Error Resume Next
            With planSheet.Shapes(shapeName)
                .Fill.ForeColor.RGB = ProgressColour(percentValue)
                .Fill.Transparency = 1 - opacityValue
            End With
            On Error GoTo 0
        End If

        If Len(labelShapeName) > 0 Then
            On Error Resume Next
            With planSheet.Shapes(labelShapeName)
                .TextFrame2.TextRange.Text = CStr(Round(percentValue, 0)) & "%" & vbCrLf & roomId
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .TextFrame2.TextRange.Font.Bold = msoTrue
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
            On Error GoTo 0
        End If
    Next rowIndex
End Sub

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
