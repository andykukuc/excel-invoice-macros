Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Me

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    Dim qtyRange   As Range
    Dim priceRange As Range
    Set qtyRange   = ws.Range("A15:A" & (subtotalRow - 1))
    Set priceRange = ws.Range("D15:D" & (subtotalRow - 1))

    Dim hitQty   As Range
    Dim hitPrice As Range
    Set hitQty   = Intersect(Target, qtyRange)
    Set hitPrice = Intersect(Target, priceRange)

    ' Guard labor unit prices — revert if anyone changes col D on a labor row
    Dim priceGuard As Range
    Set priceGuard = Intersect(Target, ws.Columns("D"))
    If Not priceGuard Is Nothing Then
        Dim gc As Range
        For Each gc In priceGuard
            If ws.Cells(gc.Row, 2).Value = "Labor" Then
                Application.EnableEvents = False
                Select Case ws.Cells(gc.Row, 3).Value
                    Case "Repair Labor @ $80.00/hr"
                        ws.Cells(gc.Row, 4).Value = 80
                    Case "Install Tire Labor @ $50.00/hr"
                        ws.Cells(gc.Row, 4).Value = 50
                End Select
                Application.EnableEvents = True
            End If
        Next gc
    End If

    If hitQty Is Nothing And hitPrice Is Nothing Then Exit Sub

    Application.EnableEvents = False

    Dim cell As Range

    If Not hitQty Is Nothing Then
        For Each cell In hitQty
            If IsNumeric(cell.Value) And cell.Value <> "" Then
                ws.Cells(cell.Row, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
            Else
                ws.Cells(cell.Row, 5).ClearContents
            End If
        Next cell
    End If

    If Not hitPrice Is Nothing Then
        For Each cell In hitPrice
            Dim qty As Variant
            qty = ws.Cells(cell.Row, 1).Value
            If IsNumeric(qty) And qty <> "" Then
                ws.Cells(cell.Row, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
            End If
        Next cell
    End If

    ' Find labor start row
    Dim laborCell As Range
    Set laborCell = ws.Columns("B").Find(What:="Labor", LookIn:=xlValues, LookAt:=xlWhole)
    If Not laborCell Is Nothing Then
        Dim lastVarRow As Long
        lastVarRow = laborCell.Row - 1

        ' Auto-expand: if the last variable row is filled, insert a new blank row
        If ws.Cells(lastVarRow, 1).Value <> "" Or ws.Cells(lastVarRow, 2).Value <> "" Then
            ws.Rows(lastVarRow + 1).Insert Shift:=xlDown
        End If

        ' Auto-collapse: delete empty variable rows, but always keep row 15
        Dim r As Long
        For r = lastVarRow To 16 Step -1   ' never delete row 15
            If ws.Cells(r, 1).Value = "" And ws.Cells(r, 2).Value = "" And _
               ws.Cells(r, 4).Value = "" And ws.Cells(r, 5).Value = "" Then
                ws.Rows(r).Delete Shift:=xlUp
            End If
        Next r
    End If

    Application.EnableEvents = True
End Sub
