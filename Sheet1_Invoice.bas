Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = Me
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    ' Only react to edits inside the line-item block, columns A-D
    Dim itemBlock As Range
    Set itemBlock = ws.Range("A15:D" & (subtotalRow - 1))
    If Intersect(Target, itemBlock) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo CleanExit

    ' Re-guard the two hourly labor rates back to 80 / 50 if changed
    Dim r As Long
    For r = 15 To subtotalRow - 1
        If ws.Cells(r, 2).Value = "Labor" Then
            Select Case ws.Cells(r, 3).Value
                Case "Repair Labor @ $80.00/hr": ws.Cells(r, 4).Value = 80
                Case "Install Tire Labor @ $50.00/hr": ws.Cells(r, 4).Value = 50
            End Select
        End If
    Next r

    UpdateLineAmounts
    UpdateFormulas

CleanExit:
    Application.EnableEvents = True
End Sub
