Attribute VB_Name = "V2BuildInstaller"
Option Explicit

' One-time build helper. It deliberately does not save while modifying the
' running VBA project; the original SaveTemplate macro is run after it returns.
Public Sub FinalizeV2BuildWithoutSaving()
    Dim projectObject As Object
    Dim workbookComponent As Object
    Dim worksheetComponent As Object
    Dim workbookCode As String
    Dim worksheetCode As String

    On Error GoTo InstallError
    Set projectObject = ThisWorkbook.VBProject
    Set workbookComponent = projectObject.VBComponents("ThisWorkbook")
    Set worksheetComponent = projectObject.VBComponents( _
        ThisWorkbook.Worksheets("Invoice").CodeName)

    workbookCode = "Option Explicit" & vbCrLf & vbCrLf & _
        "Private Sub Workbook_Open()" & vbCrLf & _
        "    InitializeInvoiceMakerV2" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)" & vbCrLf & _
        "    If bV2InternalSave Then Exit Sub" & vbCrLf & _
        "    On Error GoTo BeforeSaveError" & vbCrLf & _
        "    If InStr(1, ThisWorkbook.Name, ""Invoice_Template_Formatted_V2"", vbTextCompare) > 0 Then" & vbCrLf & _
        "        Cancel = True" & vbCrLf & _
        "        SaveInvoiceV2" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        PrepareForEditingSaveV2 ThisWorkbook.Worksheets(""Invoice"")" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "BeforeSaveError:" & vbCrLf & _
        "    Application.EnableEvents = True" & vbCrLf & _
        "    Application.ScreenUpdating = True" & vbCrLf & _
        "    Cancel = True" & vbCrLf & _
        "    MsgBox ""Excel could not prepare this invoice for saving."" & vbCrLf & vbCrLf & Err.Description, vbCritical, ""Save Cancelled""" & vbCrLf & _
        "End Sub"

    worksheetCode = "Option Explicit" & vbCrLf & vbCrLf & _
        "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    HandleInvoiceChangeV2 Me, Target" & vbCrLf & _
        "End Sub"

    ReplaceAllCode workbookComponent.CodeModule, workbookCode
    ReplaceAllCode worksheetComponent.CodeModule, worksheetCode
    InstallButtonsV2

    MsgBox "V2 event code and buttons are ready. Run SaveTemplate next.", _
           vbInformation, "V2 Build Ready"
    Exit Sub

InstallError:
    MsgBox "V2 build finalization failed:" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "V2 Build Failed"
End Sub

Private Sub ReplaceAllCode(ByVal codeModule As Object, ByVal newCode As String)
    If codeModule.CountOfLines > 0 Then codeModule.DeleteLines 1, codeModule.CountOfLines
    codeModule.AddFromString newCode
End Sub
