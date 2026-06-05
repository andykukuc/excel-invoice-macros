Option Explicit

Private Sub Workbook_Open()
    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
        ResetTemplate skipConfirm:=True
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If bSaveAsTemplate Then Exit Sub
    On Error GoTo BeforeSaveError
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    AlignLineItems
    UpdateLineAmounts
    UpdateFormulas
    FormatInvoice          ' make it pretty every save

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
        Cancel = True
        SaveInvoice
    End If
    Exit Sub

BeforeSaveError:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Cancel = True
    MsgBox "Error preparing invoice for save:" & Chr(13) & _
           Err.Description & Chr(13) & Chr(13) & "Save cancelled.", _
           vbCritical, "Save Error"
End Sub
