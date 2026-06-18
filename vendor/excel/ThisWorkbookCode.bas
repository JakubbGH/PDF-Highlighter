Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    RefreshZoneColours
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If Sh.Name <> "Progress" Then Exit Sub
    If Intersect(Target, Sh.Range("A:A,D:E")) Is Nothing Then Exit Sub

    On Error GoTo CleanExit
    Application.EnableEvents = False
    RefreshZoneColours

CleanExit:
    Application.EnableEvents = True
End Sub

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
        percentValue = CDbl(Val(Replace(CStr(progressSheet.Cells(rowIndex, "D").Value), "%", "")))
        opacityValue = CDbl(Val(Replace(CStr(progressSheet.Cells(rowIndex, "E").Value), "%", ""))) / 100
        labelShapeName = CStr(progressSheet.Cells(rowIndex, "G").Value)

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
    Next rowIndex
End Sub

Private Sub UpdateZoneLabel(ByVal frame As TextFrame2, ByVal shapeWidth As Double, ByVal shapeHeight As Double, ByVal roomId As String, ByVal percentValue As Double)
    Dim percentText As String
    Dim labelText As String
    Dim percentFontSize As Double

    percentText = CStr(Round(percentValue, 0)) & "%"
    labelText = percentText & vbCrLf & roomId
    percentFontSize = LabelPercentFontSize(shapeWidth, shapeHeight)

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
        .TextRange.Font.Size = percentFontSize
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

Private Function LabelPercentFontSize(ByVal shapeWidth As Double, ByVal shapeHeight As Double) As Double
    Dim sizeValue As Double

    sizeValue = shapeWidth / 4
    If shapeHeight / 2.4 < sizeValue Then sizeValue = shapeHeight / 2.4
    If sizeValue < 9 Then sizeValue = 9
    If sizeValue > 19 Then sizeValue = 19

    LabelPercentFontSize = sizeValue
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
