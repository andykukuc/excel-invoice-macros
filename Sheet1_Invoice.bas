Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = Me
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    ' Auto-fill A8, C7, C8 when A7 is entered
    If Not Intersect(Target, ws.Range("A7")) Is Nothing Then
        Application.EnableEvents = False
        ws.Range("A8").Value = ws.Range("A7").Value
        ws.Range("C7").Value = ws.Range("A7").Value
        ws.Range("C8").Value = ws.Range("A7").Value
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Auto-fill B11 when A11 is entered
    If Not Intersect(Target, ws.Range("A11")) Is Nothing Then
        Application.EnableEvents = False
        ws.Range("B11").Value = ws.Range("A11").Value
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Only react to edits inside the line-item block, columns A-D
    Dim itemBlock As Range
    Set itemBlock = ws.Range("A15:D" & (subtotalRow - 1))
    If Intersect(Target, itemBlock) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo CleanExit

    ShowAllBuckets          ' un-hide any rows hidden by HideEmptyBuckets before editing

    ' Re-guard the hourly labor rates back to 80 / 50 if changed.
    ' Keys off the DESCRIPTION (col C) so it protects the tire rate
    ' even though that row is now tagged "Tires", not "Labor".
    Dim r As Long
    For r = 15 To subtotalRow - 1
        Select Case ws.Cells(r, 3).Value
            Case "Repair Labor @ $80.00/hr": ws.Cells(r, 4).Value = 80
            Case "Install Tire Labor @ $50.00/hr": ws.Cells(r, 4).Value = 50
        End Select
    Next r

    UpdateLineAmounts
    UpdateFormulas

    ' --- Auto-insert a blank line ABOVE the first Labor row, so there's ---
    ' --- always an empty parts line ready. Fires no matter which column ---
    ' --- (QTY / Item / Description / Price) you filled in.               ---
    subtotalRow = FindSubtotalRow(ws)          ' refresh (UpdateFormulas may have shifted rows)
    Dim laborRow As Long: laborRow = 0
    For r = 15 To subtotalRow - 1
        If ws.Cells(r, 2).Value = "Labor" Then laborRow = r: Exit For
    Next r
    If laborRow >= 16 Then
        Dim prev As Long: prev = laborRow - 1
        ' If the row just above the Labor block is fully populated (Item + Desc),
        ' insert a fresh blank line there so the user can keep typing.
        If Trim(CStr(ws.Cells(prev, 2).Value)) <> "" _
           And Trim(CStr(ws.Cells(prev, 3).Value)) <> "" Then
            ws.Rows(laborRow).Insert Shift:=xlDown
            ' Clear any inherited formatting/content on the new blank row
            ws.Rows(laborRow).ClearContents
        End If
    End If

CleanExit:
    Application.EnableEvents = True
End Sub
