' ============================================================
'  PASTE THIS INTO: Module1
' ============================================================

Option Explicit

Public bSaveAsTemplate As Boolean
Public bSaveInProgress As Boolean   ' prevents BeforeSave re-entry during SaveAs

' ============================================================
'  SAVE TEMPLATE
' ============================================================

Public Sub SaveTemplate()
    bSaveAsTemplate = True
    Application.EnableEvents = False
    On Error GoTo TemplateError
    ThisWorkbook.Save
    Application.EnableEvents = True
    bSaveAsTemplate = False
    MsgBox "Template saved.", vbInformation, "Template Saved"
    Exit Sub
TemplateError:
    Application.EnableEvents = True
    bSaveAsTemplate = False
    MsgBox "Save failed: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
'  SHARED HELPER
' ============================================================

Public Function FindSubtotalRow(ws As Worksheet) As Long
    Dim fc As Range
    Set fc = ws.Columns("D").Find(What:="Subtotal", LookIn:=xlValues, LookAt:=xlPart)
    If fc Is Nothing Then
        Set fc = ws.Cells.Find(What:="Subtotal", LookIn:=xlValues, LookAt:=xlPart)
    End If
    FindSubtotalRow = IIf(fc Is Nothing, 0, fc.Row)
End Function

' ============================================================
'  ALIGN LINE ITEMS
' ============================================================

Public Sub AlignLineItems()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Or subtotalRow <= LINEITEM_START Then Exit Sub

    With ws.Range("A" & LINEITEM_START & ":A" & (subtotalRow - 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("B" & LINEITEM_START & ":C" & (subtotalRow - 1))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("D" & LINEITEM_START & ":E" & (subtotalRow - 1))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
End Sub

' ============================================================
'  UPDATE LINE AMOUNTS
' ============================================================

Public Sub UpdateLineAmounts()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Or subtotalRow <= LINEITEM_START Then Exit Sub

    Dim i As Long
    For i = LINEITEM_START To subtotalRow - 1
        If IsNumeric(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value <> "" Then
            ws.Cells(i, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
        End If
    Next i
End Sub

' ============================================================
'  UPDATE SUMMARY FORMULAS
' ============================================================

Public Sub UpdateFormulas()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    Dim fc As Range

    ws.Cells(subtotalRow, 5).Formula = "=SUM(E15:E" & (subtotalRow - 1) & ")"

    Set fc = ws.Columns("D").Find(What:="Parts Total", LookIn:=xlValues, LookAt:=xlPart)
    If Not fc Is Nothing Then
        fc.Offset(0, 1).Formula = "=SUMIF(B15:B" & (subtotalRow - 1) & ",""<>Labor"",E15:E" & (subtotalRow - 1) & ")"
        fc.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    Dim taxCell As Range
    Dim totalCell As Range
    Set taxCell = ws.Columns("D").Find(What:="Sales Tax", LookIn:=xlValues, LookAt:=xlPart)
    Set totalCell = ws.Columns("D").Find(What:="Total Invoice Amount", LookIn:=xlValues, LookAt:=xlPart)
    If Not totalCell Is Nothing And Not taxCell Is Nothing Then
        totalCell.Offset(0, 1).Formula = "=" & ws.Cells(subtotalRow, 5).Address & "+" & taxCell.Offset(0, 1).Address
    End If

    Dim dueCell As Range
    Dim creditCell As Range
    Set dueCell = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
    Set creditCell = ws.Columns("D").Find(What:="Payment/Credit", LookIn:=xlValues, LookAt:=xlPart)
    If Not dueCell Is Nothing And Not totalCell Is Nothing Then
        Dim totalAddr  As String
        Dim creditAddr As String
        totalAddr = totalCell.Offset(0, 1).Address
        If Not creditCell Is Nothing Then
            creditAddr = creditCell.Offset(0, 1).Address
            dueCell.Offset(0, 1).Formula = "=" & totalAddr & "-IF(ISNUMBER(" & creditAddr & ")," & creditAddr & ",0)"
        Else
            dueCell.Offset(0, 1).Formula = "=" & totalAddr
        End If
    End If

    ws.ResetAllPageBreaks
    Dim hpb As HPageBreak
    Dim splitFound As Boolean
    splitFound = False
    For Each hpb In ws.HPageBreaks
        If hpb.Location.Row > subtotalRow And hpb.Location.Row <= subtotalRow + 6 Then
            splitFound = True
            Exit For
        End If
    Next hpb
    If splitFound Then
        ws.HPageBreaks.Add Before:=ws.Rows(subtotalRow)
    End If
End Sub

' ============================================================
'  SAVE INVOICE
' ============================================================

Private Function InvoiceFieldsFilled(ws As Worksheet) As Boolean
    If Trim(CStr(ws.Cells(7, 1).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(7, 3).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(11, 1).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(11, 2).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(13, 3).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(13, 4).Value)) = "" Then Exit Function
    If Trim(CStr(ws.Cells(13, 5).Value)) = "" Then Exit Function
    InvoiceFieldsFilled = True
End Function

Public Sub SaveInvoice()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    If Not InvoiceFieldsFilled(ws) Then
        MsgBox "Please fill in all required fields before saving:" & Chr(13) & Chr(13) & _
               "  - Bill To / Ship To" & Chr(13) & _
               "  - Customer ID and Customer PO" & Chr(13) & _
               "  - Truck/Trailer, Fleet No, and Model", _
               vbExclamation, "Fields Required"
        Exit Sub
    End If

    Dim fleetNo As String, model As String
    Dim invoiceNo As String, dateStr As String, newName As String

    fleetNo = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 4).Value))
    model = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 5).Value))
    invoiceNo = Trim(CStr(ws.Cells(2, 5).Value))
    dateStr = Format(Date, "YYYYMMDD")
    newName = dateStr & "_" & invoiceNo & "_" & model & "_" & fleetNo & ".xlsm"

    ' --- Windows ---
    If Not Application.OperatingSystem Like "*Mac*" Then
        Dim defaultPath As String
        defaultPath = Environ("USERPROFILE") & "\Documents\" & newName

        Dim chosenPath As Variant
        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=defaultPath, _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save Invoice As")

        If chosenPath = False Then Exit Sub

        bSaveInProgress = True
        Application.EnableEvents = False
        On Error GoTo WinSaveError
        ThisWorkbook.SaveAs Filename:=CStr(chosenPath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.EnableEvents = True
        bSaveInProgress = False
        MsgBox "Invoice saved to:" & Chr(13) & CStr(chosenPath), vbInformation, "Invoice Saved"
        Exit Sub

WinSaveError:
        Application.EnableEvents = True
        bSaveInProgress = False
        MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
        Exit Sub
    End If

    ' --- Mac: dialog grants sandbox access to chosen folder ---
    Dim macChosenPath As Variant
    macChosenPath = Application.GetSaveAsFilename( _
        InitialFileName:=newName, _
        Title:="Save Invoice As (.xlsm) — navigate to your invoices folder")

    If macChosenPath = False Then Exit Sub

    ' Strip whatever extension the dialog set and enforce .xlsm
    Dim finalPath As String
    finalPath = CStr(macChosenPath)
    Dim dotPos As Long
    dotPos = InStrRev(finalPath, ".")
    If dotPos > 0 Then finalPath = Left(finalPath, dotPos - 1)
    finalPath = finalPath & ".xlsm"

    bSaveInProgress = True
    Application.EnableEvents = False
    On Error GoTo MacSaveError
    ThisWorkbook.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.EnableEvents = True
    bSaveInProgress = False
    MsgBox "Invoice saved to:" & Chr(13) & finalPath, vbInformation, "Invoice Saved"
    Exit Sub

MacSaveError:
    Application.EnableEvents = True
    bSaveInProgress = False
    MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
End Sub

' ============================================================
'  RESET TEMPLATE
' ============================================================

Public Sub ResetTemplate(Optional skipConfirm As Boolean = False)
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    If Not skipConfirm Then
        If MsgBox("Clear all invoice data and reset the template?", _
                  vbYesNo + vbExclamation, "Reset Template") <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ResetError

    On Error Resume Next
    ws.Unprotect
    On Error GoTo ResetError

    ws.Range("A7:B7").ClearContents
    ws.Range("A8:B8").ClearContents
    ws.Range("C7:E7").ClearContents
    ws.Range("C8:E8").ClearContents

    ws.Cells(11, 1).ClearContents
    ws.Cells(11, 2).ClearContents

    ws.Cells(13, 3).ClearContents
    ws.Cells(13, 4).ClearContents
    ws.Cells(13, 5).ClearContents

    ws.Range("A15:E16").ClearContents

    ws.Cells(17, 1).Value = 0
    ws.Cells(17, 2).Value = "Labor"
    ws.Cells(17, 3).Value = "Repair Labor @ $80.00/hr"
    ws.Cells(17, 4).Value = 80
    ws.Cells(17, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"

    ws.Cells(18, 1).Value = 0
    ws.Cells(18, 2).Value = "Labor"
    ws.Cells(18, 3).Value = "Install Tire Labor @ $50.00/hr"
    ws.Cells(18, 4).Value = 50
    ws.Cells(18, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"

    Dim clearCell As Range
    Set clearCell = ws.Columns("D").Find(What:="Sales Tax", LookIn:=xlValues, LookAt:=xlPart)
    If Not clearCell Is Nothing Then clearCell.Offset(0, 1).ClearContents

    Set clearCell = ws.Columns("D").Find(What:="Payment/Credit", LookIn:=xlValues, LookAt:=xlPart)
    If Not clearCell Is Nothing Then clearCell.Offset(0, 1).ClearContents

    Randomize Timer
    ws.Cells(2, 5).Value = Format(Int(Rnd() * 100000), "00000")
    ws.Cells(3, 5).Value = Date
    ws.Cells(4, 5).Value = 1

    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ResetError:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Reset failed: " & Err.Description, vbCritical, "Reset Error"
End Sub

Public Sub ResetTemplateManual()
    ResetTemplate skipConfirm:=False
End Sub

' ============================================================
'  HELPER: STRIP ILLEGAL FILENAME CHARACTERS
' ============================================================

Private Function CleanFileName(s As String) As String
    Dim illegal As String, i As Integer, c As String, result As String
    illegal = "/\:*?""<>|"
    result = s
    For i = 1 To Len(illegal)
        c = Mid(illegal, i, 1)
        result = Join(Split(result, c), "_")
    Next i
    Do While InStr(result, "__") > 0
        result = Join(Split(result, "__"), "_")
    Loop
    CleanFileName = Trim(result)
End Function


' ============================================================
'  PASTE THIS INTO: ThisWorkbook
' ============================================================

'Option Explicit
'
'Private Sub Workbook_Open()
'    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
'        ResetTemplate skipConfirm:=True
'    End If
'End Sub
'
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    If bSaveAsTemplate Then Exit Sub
'    If bSaveInProgress Then Exit Sub    ' <-- ADD THIS LINE (prevents re-entry loop)
'
'    On Error GoTo BeforeSaveError
'
'    Application.ScreenUpdating = False
'    AlignLineItems
'    UpdateLineAmounts
'    UpdateFormulas
'    Application.ScreenUpdating = True
'
'    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
'        Cancel = True
'        If Application.OperatingSystem Like "*Mac*" Then
'            Application.OnTime Now, "SaveInvoice"
'        Else
'            SaveInvoice
'        End If
'    End If
'    Exit Sub
'
'BeforeSaveError:
'    Application.ScreenUpdating = True
'    Cancel = True
'    MsgBox "Error preparing invoice for save:" & Chr(13) & _
'           Err.Description & Chr(13) & Chr(13) & "Save cancelled.", _
'           vbCritical, "Save Error"
'End Sub


' ============================================================
'  Sheet1 (Invoice) — NO CHANGES NEEDED
' ============================================================
