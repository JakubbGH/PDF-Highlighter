Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("A:B,E:F")) Is Nothing Then Exit Sub

    On Error GoTo CleanExit
    Application.EnableEvents = False
    RefreshZoneColours

CleanExit:
    Application.EnableEvents = True
End Sub
