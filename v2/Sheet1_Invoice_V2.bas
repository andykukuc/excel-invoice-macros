Attribute VB_Name = "Sheet1_Invoice_V2"
Option Explicit

' Copy this procedure into the Invoice worksheet code module for V2.
Private Sub Worksheet_Change(ByVal Target As Range)
    HandleInvoiceChangeV2 Me, Target
End Sub
