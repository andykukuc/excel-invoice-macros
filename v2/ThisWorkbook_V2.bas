Attribute VB_Name = "ThisWorkbook_V2"
Option Explicit

' Copy these procedures into the actual ThisWorkbook code module for V2.
Private Sub Workbook_Open()
    InitializeInvoiceMakerV2
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If bV2InternalSave Then Exit Sub

    On Error GoTo BeforeSaveError
    If InStr(1, ThisWorkbook.Name, "Invoice_Template_Formatted_V2", vbTextCompare) > 0 Then
        Cancel = True
        SaveInvoiceV2
    Else
        PrepareForEditingSaveV2 ThisWorkbook.Worksheets("Invoice")
    End If
    Exit Sub

BeforeSaveError:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Cancel = True
    MsgBox "Excel could not prepare this invoice for saving." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Save Cancelled"
End Sub
