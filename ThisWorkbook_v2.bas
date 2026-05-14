Option Explicit

Private Sub Workbook_Open()
    ' Always reset when opening the master template
    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
        ResetTemplate skipConfirm:=True
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Allow SaveTemplate macro to bypass this hook
    If bSaveAsTemplate Then Exit Sub

    On Error GoTo BeforeSaveError

    Application.ScreenUpdating = False
    AlignLineItems
    UpdateLineAmounts
    UpdateFormulas
    Application.ScreenUpdating = True

    ' Template: always block the native save and route through SaveInvoice
    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
        Cancel = True
        If Application.OperatingSystem Like "*Mac*" Then
            Application.OnTime Now, "SaveInvoice"
        Else
            SaveInvoice
        End If
    End If
    Exit Sub

BeforeSaveError:
    Application.ScreenUpdating = True
    Cancel = True
    MsgBox "Error preparing invoice for save:" & Chr(13) & _
           Err.Description & Chr(13) & Chr(13) & "Save cancelled.", _
           vbCritical, "Save Error"
End Sub
