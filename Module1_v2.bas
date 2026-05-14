Option Explicit

' Flag used by SaveTemplate to bypass BeforeSave hook
Public bSaveAsTemplate As Boolean

' ============================================================
'  SAVE TEMPLATE  (use this to save changes to the template itself)
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
'  SHARED HELPER — finds the Subtotal row dynamically
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

    ' Subtotal — sum all line items (parts + labor)
    ws.Cells(subtotalRow, 5).Formula = "=SUM(E15:E" & (subtotalRow - 1) & ")"

    ' Parts Total — sum only non-labor rows
    Set fc = ws.Columns("D").Find(What:="Parts Total", LookIn:=xlValues, LookAt:=xlPart)
    If Not fc Is Nothing Then
        fc.Offset(0, 1).Formula = "=SUMIF(B15:B" & (subtotalRow - 1) & ",""<>Labor"",E15:E" & (subtotalRow - 1) & ")"
        fc.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    ' Total Invoice Amount = Subtotal + Sales Tax
    Dim taxCell As Range
    Dim totalCell As Range
    Set taxCell  = ws.Columns("D").Find(What:="Sales Tax", LookIn:=xlValues, LookAt:=xlPart)
    Set totalCell = ws.Columns("D").Find(What:="Total Invoice Amount", LookIn:=xlValues, LookAt:=xlPart)
    If Not totalCell Is Nothing And Not taxCell Is Nothing Then
        totalCell.Offset(0, 1).Formula = "=" & ws.Cells(subtotalRow, 5).Address & "+" & taxCell.Offset(0, 1).Address
    End If

    ' Total Due = Total Invoice Amount - Payment/Credit Applied
    Dim dueCell As Range
    Dim creditCell As Range
    Set dueCell    = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
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

    ' Keep totals block together — if a page break falls inside it, move it to before Subtotal
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
'  Only saves when all required fields are filled.
'  Never overwrites the template — always saves as a new file.
' ============================================================

Private Function InvoiceFieldsFilled(ws As Worksheet) As Boolean
    If Trim(CStr(ws.Cells(7, 1).Value)) = "" Then Exit Function   ' Bill To
    If Trim(CStr(ws.Cells(7, 3).Value)) = "" Then Exit Function   ' Ship To
    If Trim(CStr(ws.Cells(11, 1).Value)) = "" Then Exit Function  ' Customer ID
    If Trim(CStr(ws.Cells(11, 2).Value)) = "" Then Exit Function  ' Customer PO
    If Trim(CStr(ws.Cells(13, 3).Value)) = "" Then Exit Function  ' Truck/Trailer
    If Trim(CStr(ws.Cells(13, 4).Value)) = "" Then Exit Function  ' Fleet No
    If Trim(CStr(ws.Cells(13, 5).Value)) = "" Then Exit Function  ' Model
    InvoiceFieldsFilled = True
End Function

Public Sub SaveInvoice()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    ' Block save if required fields are not filled — do NOT save the template
    If Not InvoiceFieldsFilled(ws) Then
        MsgBox "Please fill in all required fields before saving:" & Chr(13) & Chr(13) & _
               "  - Bill To / Ship To" & Chr(13) & _
               "  - Customer ID and Customer PO" & Chr(13) & _
               "  - Truck/Trailer, Fleet No, and Model", _
               vbExclamation, "Fields Required"
        Exit Sub
    End If

    ' Build filename
    Dim fleetNo As String, model As String
    Dim invoiceNo As String, dateStr As String, newName As String

    fleetNo   = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 4).Value))
    model     = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 5).Value))
    invoiceNo = Trim(CStr(ws.Cells(2, 5).Value))
    dateStr   = Format(Date, "YYYYMMDD")
    newName   = dateStr & "_" & invoiceNo & "_" & model & "_" & fleetNo & ".xlsm"

    ' --- Windows: show Save As dialog pre-filled with generated name ---
    If Not Application.OperatingSystem Like "*Mac*" Then
        Dim defaultPath As String
        defaultPath = Environ("USERPROFILE") & "\Documents\" & newName

        Dim chosenPath As Variant
        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=defaultPath, _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save Invoice As")

        If chosenPath = False Then Exit Sub   ' user cancelled

        Application.EnableEvents = False
        On Error GoTo WinSaveError
        ThisWorkbook.SaveAs Filename:=CStr(chosenPath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.EnableEvents = True
        MsgBox "Invoice saved to:" & Chr(13) & CStr(chosenPath), vbInformation, "Invoice Saved"
        Exit Sub

WinSaveError:
        Application.EnableEvents = True
        MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
        Exit Sub
    End If

    ' --- Mac: folder picker against the Samba share ---
    Dim baseDir As String
    baseDir = "/Volumes/invoices/"

    Dim folders() As String, folderCount As Integer, folderList As String, entry As String
    folderCount = 0
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
        saveDir = baseDir
        GoTo DoMacSave
    End If

    Dim choice As Variant
    choice = InputBox("Choose a folder:" & Chr(13) & Chr(13) & _
        folderList & Chr(13) & "0)  (root of invoices)", "Save Location", "")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub

    Dim choiceNum As Integer
    choiceNum = CInt(choice)

    If choiceNum >= 1 And choiceNum <= folderCount Then
        saveDir = baseDir & folders(choiceNum) & "/"

        Dim subFolders() As String, subCount As Integer, subList As String, subEntry As String
        subCount = 0
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
            If Not IsNumeric(subChoice) Then Exit Sub
            Dim subChoiceNum As Integer
            subChoiceNum = CInt(subChoice)
            If subChoiceNum >= 1 And subChoiceNum <= subCount Then
                saveDir = saveDir & subFolders(subChoiceNum) & "/"
            End If
        End If
    Else
        saveDir = baseDir
    End If

DoMacSave:
    Dim sharePath As String
    Dim localPath As String
    sharePath = saveDir & newName
    localPath = Environ("HOME") & "/Documents/" & newName

    Application.EnableEvents = False
    On Error GoTo MacFallback
    ThisWorkbook.SaveAs Filename:=sharePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.EnableEvents = True
    MsgBox "Invoice saved to:" & Chr(13) & sharePath, vbInformation, "Invoice Saved"
    Exit Sub

MacFallback:
    On Error GoTo MacSaveError
    ThisWorkbook.SaveAs Filename:=localPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.EnableEvents = True
    MsgBox "Share unavailable - saved locally:" & Chr(13) & localPath, vbExclamation, "Invoice Saved"
    Exit Sub

MacSaveError:
    Application.EnableEvents = True
    MsgBox "Save failed entirely: " & Err.Description, vbCritical, "Save Error"
End Sub

' ============================================================
'  RESET TEMPLATE  (runs automatically on open)
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

    ' Bill To / Ship To
    ws.Range("A7:B7").ClearContents
    ws.Range("A8:B8").ClearContents
    ws.Range("C7:E7").ClearContents
    ws.Range("C8:E8").ClearContents

    ' Customer fields
    ws.Cells(11, 1).ClearContents   ' Customer ID
    ws.Cells(11, 2).ClearContents   ' Customer PO

    ' Vehicle fields
    ws.Cells(13, 3).ClearContents   ' Truck/Trailer
    ws.Cells(13, 4).ClearContents   ' Fleet No
    ws.Cells(13, 5).ClearContents   ' Model

    ' Line items rows 15-16 (variable)
    ws.Range("A15:E16").ClearContents

    ' Restore fixed labor rows 17-18
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

    ' Clear manually-entered totals fields by label
    Dim clearCell As Range
    Set clearCell = ws.Columns("D").Find(What:="Sales Tax", LookIn:=xlValues, LookAt:=xlPart)
    If Not clearCell Is Nothing Then clearCell.Offset(0, 1).ClearContents

    Set clearCell = ws.Columns("D").Find(What:="Payment/Credit", LookIn:=xlValues, LookAt:=xlPart)
    If Not clearCell Is Nothing Then clearCell.Offset(0, 1).ClearContents

    ' Fresh invoice number and today's date
    Randomize Timer
    ws.Cells(2, 5).Value = Format(Int(Rnd() * 100000), "00000")
    ws.Cells(3, 5).Value = Date
    ws.Cells(4, 5).Value = 1   ' Page number

    ' No sheet protection — causes more issues than it solves

    ' Page setup: fit 1 page wide, unlimited pages tall — keeps text readable
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

' Wrapper so it appears in Alt+F8 macro list
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
