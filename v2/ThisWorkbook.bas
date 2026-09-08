Attribute VB_Name = "ThisWorkbook"
Option Explicit

' Copy these procedures into the actual ThisWorkbook code module for V2.
Private Sub Workbook_Open()
    InitializeInvoiceMakerV2
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If bV2InternalSave Or bV2WorkflowBusy Then Exit Sub

    On Error GoTo BeforeSaveError
    If SaveAsUI Then
        Cancel = True
        MsgBox "Use SAVE INVOICE to save the editable invoice and its PDF, " & _
               "or use PRINT to send it to the printer.", _
               vbInformation, "Use the Invoice Buttons"
        Exit Sub
    End If

    If InStr(1, ThisWorkbook.Name, "Invoice_Template_Formatted_V2", vbTextCompare) > 0 Then
        Cancel = True
        SaveInvoiceV2
    Else
        bV2WorkflowBusy = True
        Application.EnableEvents = False
        PrepareForEditingSaveV2 ThisWorkbook.Worksheets("Invoice")
        Application.EnableEvents = True
        bV2WorkflowBusy = False
    End If
    Exit Sub

BeforeSaveError:
    bV2WorkflowBusy = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Cancel = True
    MsgBox "Excel could not prepare this invoice for saving." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Save Cancelled"
End Sub
