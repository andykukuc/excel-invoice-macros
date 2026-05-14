' ============================================================
'  FILE: ThisWorkbook  (paste into the ThisWorkbook module)
' ============================================================
Option Explicit

Private Sub Workbook_Open()
    With Worksheets("Invoice").Range("E2")
        If .HasFormula Or .Value = "" Then
            ' Use Timer as seed for better uniqueness on fast machines
            Randomize Timer
            .Value = Format(Int(Rnd() * 100000), "00000")
        End If
    End With
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo BeforeSaveError

    ' Suppress screen flicker while formatting subs run
    Application.ScreenUpdating = False

    ' Run formatting/formula updates on every save
    AlignLineItems
    UpdateLineAmounts
    UpdateFormulas

    Application.ScreenUpdating = True

    ' If this is still the template, cancel the direct save and
    ' defer to SaveInvoice via OnTime. Calling SaveAs directly inside
    ' BeforeSave causes a re-entrant crash on Mac Excel.
    If InStr(1, ThisWorkbook.Name, "Invoice_Template", vbTextCompare) > 0 Then
        Cancel = True
        Application.OnTime Now, "SaveInvoice"
    End If
    Exit Sub

BeforeSaveError:
    Application.ScreenUpdating = True
    Cancel = True   ' Don't save in a potentially broken state
    MsgBox "An error occurred while preparing the invoice for save:" & Chr(13) & _
           Err.Description & Chr(13) & Chr(13) & "Save cancelled.", vbCritical, "Save Error"
End Sub


' ============================================================
'  FILE: Module1  (paste into a standard module)
' ============================================================
Option Explicit

' ============================================================
'  SHARED HELPER
' ============================================================

Private Function FindSubtotalRow(ws As Worksheet) As Long
    Dim fc As Range
    Set fc = ws.Columns("D").Find(What:="Subtotal", LookIn:=xlValues, LookAt:=xlPart)
    If fc Is Nothing Then
        Set fc = ws.Cells.Find(What:="Subtotal", LookIn:=xlValues, LookAt:=xlPart)
    End If
    If fc Is Nothing Then
        FindSubtotalRow = 0
    Else
        FindSubtotalRow = fc.Row
    End If
End Function

' ============================================================
'  ALIGN LINE ITEMS
' ============================================================

Public Sub AlignLineItems()
    Const LINEITEM_START_ROW As Long = 15

    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    ' Guard: subtotal must be below the first line-item row
    If subtotalRow <= LINEITEM_START_ROW Then
        MsgBox "Layout error: Subtotal row (" & subtotalRow & ") is at or above line-item start (" & LINEITEM_START_ROW & ").", vbCritical, "AlignLineItems"
        Exit Sub
    End If

    With ws.Range("A" & LINEITEM_START_ROW & ":A" & (subtotalRow - 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

    With ws.Range("B" & LINEITEM_START_ROW & ":C" & (subtotalRow - 1))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With

    With ws.Range("D" & LINEITEM_START_ROW & ":E" & (subtotalRow - 1))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
End Sub

' ============================================================
'  UPDATE LINE AMOUNTS
' ============================================================

Public Sub UpdateLineAmounts()
    Const LINEITEM_START_ROW As Long = 15

    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    If subtotalRow <= LINEITEM_START_ROW Then
        MsgBox "Layout error: Subtotal row (" & subtotalRow & ") is at or above line-item start (" & LINEITEM_START_ROW & ").", vbCritical, "UpdateLineAmounts"
        Exit Sub
    End If

    Dim i As Long
    For i = LINEITEM_START_ROW To subtotalRow - 1
        ' Only write formula when column A contains a numeric quantity
        If IsNumeric(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value <> "" Then
            ' RC[-4] = col A (qty), RC[-1] = col D (unit price), relative to col E
            ' FormulaR1C1 is locale-safe (works on European Excel with semicolon separators)
            ws.Cells(i, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
        End If
    Next i
End Sub

' ============================================================
'  UPDATE SUMMARY FORMULAS
' ============================================================

Public Sub UpdateFormulas()
    ' Expected layout below Subtotal row:
    '   subtotalRow + 0 : Subtotal       (E = SUM of line amounts)
    '   subtotalRow + 1 : Tax            (E = tax amount, entered manually)
    '   subtotalRow + 2 : Total          (E = Subtotal + Tax)
    '   subtotalRow + 3 : Discount       (E = discount amount, entered manually)
    '   subtotalRow + 4 : Grand Total    (E = Total - Discount)
    Const TAX_OFFSET      As Long = 1
    Const TOTAL_OFFSET    As Long = 2
    Const DISC_OFFSET     As Long = 3
    Const GRANDTOT_OFFSET As Long = 4

    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    ' Bounds check: ensure offsets stay within the sheet
    If subtotalRow + GRANDTOT_OFFSET > ws.Rows.Count Then
        MsgBox "Invoice layout error: Subtotal row too close to sheet boundary.", vbCritical, "UpdateFormulas"
        Exit Sub
    End If

    ws.Cells(subtotalRow, 5).Formula = "=SUM(E15:E" & (subtotalRow - 1) & ")"

    ws.Cells(subtotalRow + TOTAL_OFFSET, 5).Formula = _
        "=E" & subtotalRow & "+E" & (subtotalRow + TAX_OFFSET)

    ws.Cells(subtotalRow + GRANDTOT_OFFSET, 5).Formula = _
        "=E" & (subtotalRow + TOTAL_OFFSET) & _
        "-IF(ISNUMBER(E" & (subtotalRow + DISC_OFFSET) & "),E" & (subtotalRow + DISC_OFFSET) & ",0)"
End Sub

' ============================================================
'  SAVE INVOICE
' ============================================================

Public Sub SaveInvoice()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    ' --- Validate required cells before building the filename ---
    If Trim(CStr(ws.Range("E2").Value)) = "" Then
        MsgBox "Invoice number (E2) is empty. Cannot save.", vbCritical, "Save Cancelled"
        Exit Sub
    End If
    If Trim(CStr(ws.Range("D13").Value)) = "" Then
        MsgBox "Fleet number (D13) is empty. Cannot save.", vbCritical, "Save Cancelled"
        Exit Sub
    End If
    If Trim(CStr(ws.Range("E13").Value)) = "" Then
        MsgBox "Model (E13) is empty. Cannot save.", vbCritical, "Save Cancelled"
        Exit Sub
    End If

    Dim fleetNo As String, model As String
    Dim dateStr As String, invoiceNo As String
    Dim newName As String, sharePath As String, localPath As String

    fleetNo   = ws.Range("D13").Value
    model     = ws.Range("E13").Value
    invoiceNo = ws.Range("E2").Value
    dateStr   = Format(Date, "YYYYMMDD")

    model   = Application.WorksheetFunction.Clean(model)
    fleetNo = Application.WorksheetFunction.Clean(fleetNo)
    model   = CleanFileName(model)
    fleetNo = CleanFileName(fleetNo)

    newName = dateStr & "_" & invoiceNo & "_" & model & "_" & fleetNo & ".xlsm"

    ' --- OS gate: Windows uses Save As dialog, Mac uses folder picker ---
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' Default the dialog to the same folder the template lives in
        Dim defaultPath As String
        defaultPath = ThisWorkbook.Path & "\" & newName

        Dim chosenPath As Variant
        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=defaultPath, _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save Invoice As")

        ' User cancelled the dialog
        If chosenPath = False Then Exit Sub

        Application.EnableEvents = False
        On Error GoTo WinSaveFailed
        ThisWorkbook.SaveAs Filename:=CStr(chosenPath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.EnableEvents = True
        MsgBox "Saved to:" & Chr(13) & CStr(chosenPath), vbInformation, "Invoice Saved"
        Exit Sub

WinSaveFailed:
        Application.EnableEvents = True
        MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
        Exit Sub
    End If

    ' --- Mac: folder picker against the Samba share ---
    Dim macHome As String
    macHome = Environ("HOME")
    If Len(macHome) = 0 Then
        MsgBox "Cannot determine home directory. Save cancelled.", vbCritical, "Save Error"
        Exit Sub
    End If

    Dim baseDir As String
    baseDir = "/Volumes/invoices/"

    ' Level 1: enumerate top-level subfolders
    ' Note: Excel's shared Find/Dir state is modified here.
    Dim folders() As String, folderCount As Integer, folderList As String, entry As String
    folderCount = 0
    folderList = ""
    entry = Dir(baseDir, vbDirectory)
    Do While entry <> ""
        If entry <> "." And entry <> ".." And InStr(entry, ".") = 0 Then
            folderCount = folderCount + 1
            ReDim Preserve folders(1 To folderCount)
            folders(folderCount) = entry
            folderList = folderList & folderCount & ")  " & entry & Chr(13)
        End If
        entry = Dir()
    Loop

    Dim saveDir As String

    If folderCount = 0 Then
        ' No subfolders found - save to root
        MsgBox "No subfolders found in " & baseDir & ". Saving to root.", vbInformation
        saveDir = baseDir
        GoTo DoSave
    End If

    Dim choice As Variant
    choice = InputBox("Choose a folder to save into:" & Chr(13) & Chr(13) & _
        folderList & Chr(13) & "0)  (root of invoices)", "Save Location", "")
    If choice = "" Then Exit Sub

    ' Validate numeric input
    If Not IsNumeric(choice) Then
        MsgBox "Invalid selection. Please enter a number.", vbExclamation, "Save Cancelled"
        Exit Sub
    End If

    Dim choiceNum As Integer
    choiceNum = CInt(choice)

    If choiceNum >= 1 And choiceNum <= folderCount Then
        saveDir = baseDir & folders(choiceNum) & "/"

        ' Level 2: subfolders of chosen folder (e.g. Excel / PDF / Word)
        Dim subFolders() As String, subCount As Integer, subList As String, subEntry As String
        subCount = 0
        subList = ""
        subEntry = Dir(saveDir, vbDirectory)
        Do While subEntry <> ""
            If subEntry <> "." And subEntry <> ".." And InStr(subEntry, ".") = 0 Then
                subCount = subCount + 1
                ReDim Preserve subFolders(1 To subCount)
                subFolders(subCount) = subEntry
                subList = subList & subCount & ")  " & subEntry & Chr(13)
            End If
            subEntry = Dir()
        Loop

        If subCount > 0 Then
            Dim subChoice As Variant
            subChoice = InputBox("Choose subfolder within " & folders(choiceNum) & ":" & Chr(13) & Chr(13) & _
                subList & Chr(13) & "0)  (save directly in " & folders(choiceNum) & ")", _
                "Save Location", "")
            If subChoice = "" Then Exit Sub

            If Not IsNumeric(subChoice) Then
                MsgBox "Invalid selection. Please enter a number.", vbExclamation, "Save Cancelled"
                Exit Sub
            End If

            Dim subChoiceNum As Integer
            subChoiceNum = CInt(subChoice)
            If subChoiceNum >= 1 And subChoiceNum <= subCount Then
                saveDir = saveDir & subFolders(subChoiceNum) & "/"
            End If
        End If
    Else
        saveDir = baseDir
    End If

DoSave:
    sharePath = saveDir & newName
    localPath = macHome & "/Documents/" & newName

    Application.EnableEvents = False

    On Error GoTo SaveLocally
    ThisWorkbook.SaveAs Filename:=sharePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.EnableEvents = True
    MsgBox "Saved to:" & Chr(13) & sharePath, vbInformation, "Invoice Saved"
    Exit Sub

SaveLocally:
    On Error GoTo SaveFailed
    ThisWorkbook.SaveAs Filename:=localPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.EnableEvents = True
    MsgBox "Share not writable - saved locally instead:" & Chr(13) & localPath, vbExclamation, "Invoice Saved"
    Exit Sub

SaveFailed:
    Application.EnableEvents = True
    MsgBox "Save failed entirely: " & Err.Description, vbCritical, "Save Error"
End Sub

' ============================================================
'  HELPER: CLEAN ILLEGAL CHARACTERS FROM FILENAME
' ============================================================

Private Function CleanFileName(s As String) As String
    Dim illegal As String, i As Integer, c As String, result As String
    illegal = "/\:*?""<>|"
    result = s
    For i = 1 To Len(illegal)
        c = Mid(illegal, i, 1)
        result = Join(Split(result, c), "_")
    Next i
    ' Collapse consecutive underscores from multiple replacements (e.g. "A:/B" -> "A__B" -> "A_B")
    Do While InStr(result, "__") > 0
        result = Join(Split(result, "__"), "_")
    Loop
    CleanFileName = Trim(result)
End Function

' ============================================================
'  RESET TEMPLATE
' ============================================================

Public Sub ResetTemplate()
    ' Prompts user then wipes all data-entry cells back to blank,
    ' and regenerates a fresh invoice number.

    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    Dim answer As VbMsgBoxResult
    answer = MsgBox("This will clear all invoice data and reset the template." & Chr(13) & Chr(13) & _
                    "Are you sure?", vbYesNo + vbExclamation, "Reset Template")
    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ResetError

    ' Unprotect sheet if protected (re-protected at end)
    Dim wasProtected As Boolean
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    ' --- Bill To / Ship To (two rows each) ---
    ' Uses .Cells() instead of .Range() to safely handle merged cells
    ws.Cells(7, 1).ClearContents      ' Bill To - company name
    ws.Cells(8, 1).ClearContents      ' Bill To - address
    ws.Cells(7, 3).ClearContents      ' Ship To - company name
    ws.Cells(8, 3).ClearContents      ' Ship To - address

    ' --- Customer fields ---
    ws.Cells(11, 1).ClearContents     ' Customer ID value
    ws.Cells(11, 2).ClearContents     ' Customer PO value

    ' --- Vehicle fields ---
    ws.Cells(13, 4).ClearContents     ' Fleet No
    ws.Cells(13, 5).ClearContents     ' Model

    ' --- Line items (rows 15 to Subtotal - 1) ---
    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow > 15 Then
        ws.Range("A15:E" & (subtotalRow - 1)).ClearContents
    End If

    ' --- Summary rows below Subtotal (Tax and Discount only — formulas stay) ---
    If subtotalRow > 0 Then
        ws.Cells(subtotalRow + 1, 5).ClearContents   ' Tax amount
        ws.Cells(subtotalRow + 3, 5).ClearContents   ' Discount amount
    End If

    ' --- Regenerate invoice number ---
    Randomize Timer
    ws.Cells(2, 5).Value = Format(Int(Rnd() * 100000), "00000")

    ' Re-protect if it was protected before
    If wasProtected Then ws.Protect

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Template has been reset. New invoice number: " & ws.Cells(2, 5).Value, vbInformation, "Reset Complete"
    Exit Sub

ResetError:
    If wasProtected Then ws.Protect
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Reset failed: " & Err.Description & Chr(13) & Chr(13) & _
           "If the sheet is password-protected, add the password to the Unprotect call.", vbCritical, "Reset Error"
End Sub
