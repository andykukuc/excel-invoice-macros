Option Explicit

Public bSaveAsTemplate As Boolean

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
        Dim chosenPath As Variant
        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=Environ("USERPROFILE") & "\Documents\" & newName, _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save Invoice As")
        If chosenPath = False Then Exit Sub
        On Error GoTo WinSaveError
        ThisWorkbook.SaveCopyAs Filename:=CStr(chosenPath)
        MsgBox "Invoice saved to:" & Chr(13) & CStr(chosenPath), vbInformation, "Invoice Saved"
        Exit Sub
WinSaveError:
        MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
        Exit Sub
    End If

    ' --- Mac: InputBox folder picker ---
    Dim baseDir As String
    baseDir = "/Users/andykukuc/mnt/invoices/"

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
    sharePath = saveDir & newName

    On Error GoTo MacSaveError
    ThisWorkbook.SaveCopyAs Filename:=sharePath
    MsgBox "Invoice saved to:" & Chr(13) & sharePath, vbInformation, "Invoice Saved"
    Exit Sub

MacSaveError:
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
'  FIT INVOICE TO ONE PRINTED PAGE
' ============================================================

Public Sub FitToOnePage()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")

    ' Find the bottom of the invoice (Total Due row + a small buffer)
    Dim lastRow As Long
    Dim fc As Range
    Set fc = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
    If Not fc Is Nothing Then
        lastRow = fc.Row + 1
    Else
        ' Fall back to the bottom of the used range
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    End If

    ' Hide empty line item rows to reclaim vertical space
    Const LINEITEM_START As Long = 15
    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow > LINEITEM_START Then
        Dim i As Long
        For i = LINEITEM_START To subtotalRow - 1
            Dim qty As Variant
            qty = ws.Cells(i, 1).Value
            ws.Rows(i).Hidden = (Trim(CStr(qty)) = "" Or CStr(qty) = "0")
        Next i
    End If

    ' Set print area to columns A:E
    ws.PageSetup.PrintArea = "A1:E" & lastRow

    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .CenterHorizontally = True
        .PrintGridlines = False
        .PrintHeadings = False
    End With

    ws.PrintPreview
End Sub

Public Sub UnhideLineItems()
    Dim ws As Worksheet
    Set ws = Worksheets("Invoice")
    Const LINEITEM_START As Long = 15
    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow > LINEITEM_START Then
        ws.Rows(LINEITEM_START & ":" & (subtotalRow - 1)).Hidden = False
    End If
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
