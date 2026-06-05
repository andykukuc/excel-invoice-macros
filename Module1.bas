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
'  SHARED HELPERS
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
'  SORT LINE ITEMS  ->  Parts, then Labor, then Tires
'  Rate-bucket rows ("Repair Labor @ $80.00/hr" and
'  "Install Tire Labor @ $50.00/hr") are pinned to the BOTTOM
'  of their section. Merged rows / blanks are ignored.
'  Reads to an array, sorts in memory, writes back (crash-safe).
' ============================================================
Public Sub SortLineItems()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Or subtotalRow <= LINEITEM_START + 1 Then Exit Sub
    Dim lastItemRow As Long: lastItemRow = subtotalRow - 1

    ' Collect non-empty, non-merged item rows into an array of (A..E)
    Dim data() As Variant
    Dim n As Long: n = 0
    ReDim data(1 To (lastItemRow - LINEITEM_START + 1), 1 To 6)  ' col6 = sort key
    Dim r As Long, c As Long
    For r = LINEITEM_START To lastItemRow
        If Not IsMergedRow(ws, r) Then
            Dim tag As String: tag = Trim(CStr(ws.Cells(r, 2).Value))
            Dim desc As String: desc = Trim(CStr(ws.Cells(r, 3).Value))
            ' keep only rows that actually have content
            If tag <> "" Or desc <> "" Or ws.Cells(r, 1).Value <> "" Then
                n = n + 1
                For c = 1 To 5
                    data(n, c) = ws.Cells(r, c).Value
                Next c
                data(n, 6) = SortKey(tag, desc)
            End If
        End If
    Next r
    If n <= 1 Then Exit Sub

    ' Simple stable insertion sort on the numeric key (col 6)
    Dim i As Long, j As Long, keyVal As Double
    Dim tmp(1 To 6) As Variant
    For i = 2 To n
        For c = 1 To 6: tmp(c) = data(i, c): Next c
        keyVal = data(i, 6)
        j = i - 1
        Do While j >= 1
            If data(j, 6) > keyVal Then
                For c = 1 To 6: data(j + 1, c) = data(j, c): Next c
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        For c = 1 To 6: data(j + 1, c) = tmp(c): Next c
    Next i

    ' Write back (values only; amount formulas re-applied by UpdateLineAmounts)
    Dim outRow As Long: outRow = LINEITEM_START
    For i = 1 To n
        For c = 1 To 5
            ws.Cells(outRow, c).Value = data(i, c)
        Next c
        outRow = outRow + 1
    Next i
    ' Clear any leftover rows below the written block, up to lastItemRow
    Do While outRow <= lastItemRow
        If Not IsMergedRow(ws, outRow) Then ws.Range("A" & outRow & ":E" & outRow).ClearContents
        outRow = outRow + 1
    Loop
End Sub

' Sort key: Parts=100, Labor=200, Tires=300.
' Rate-bucket rows get +90 so they sink to the bottom of their section.
Private Function SortKey(tag As String, desc As String) As Double
    Dim base As Double
    Select Case LCase(tag)
        Case "parts": base = 100
        Case "labor", "install": base = 200
        Case "tires": base = 300
        Case Else: base = 400          ' unknown tags sort last
    End Select
    If desc = "Repair Labor @ $80.00/hr" Then base = base + 90
    If desc = "Install Tire Labor @ $50.00/hr" Then base = base + 90
    SortKey = base
End Function

' ============================================================
'  ALIGN LINE ITEMS
' ============================================================
Public Sub AlignLineItems()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
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
'  UPDATE LINE AMOUNTS  (SPLIT: Tires stays "Tires")
'  Tags (col B): Parts | Labor | Install | Tires
'    - Labor / Install -> hours roll into $80/hr repair line (tag -> "Labor")
'    - Tires           -> hours roll into $50/hr tire line  (tag stays "Tires")
'  Skips merged rows (tax note) so it never errors 1004.
' ============================================================
Public Sub UpdateLineAmounts()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Or subtotalRow <= LINEITEM_START Then Exit Sub

    Dim lastItemRow As Long: lastItemRow = subtotalRow - 1

    ' Normalize "Install" -> "Labor"
    Dim r As Long
    For r = LINEITEM_START To lastItemRow
        If StrComp(Trim(CStr(ws.Cells(r, 2).Value)), "Install", vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = "Labor"
        End If
    Next r

    ' Locate the two priced bucket rows by description
    Dim repairRow As Long: repairRow = 0
    Dim tireRow As Long: tireRow = 0
    For r = LINEITEM_START To lastItemRow
        Select Case ws.Cells(r, 3).Value
            Case "Repair Labor @ $80.00/hr": repairRow = r
            Case "Install Tire Labor @ $50.00/hr": tireRow = r
        End Select
    Next r

    ' --- Roll "Labor"-tagged hours into the $80/hr row ---
    If repairRow > 0 Then
        Dim laborHrs As Double: laborHrs = 0
        For r = LINEITEM_START To lastItemRow
            If r <> repairRow And ws.Cells(r, 2).Value = "Labor" _
               And Trim(CStr(ws.Cells(r, 3).Value)) <> "" Then
                If IsNumeric(ws.Cells(r, 1).Value) Then laborHrs = laborHrs + ws.Cells(r, 1).Value
            End If
        Next r
        ws.Cells(repairRow, 1).Value = laborHrs
        ws.Cells(repairRow, 4).Value = 80
        ws.Cells(repairRow, 2).Value = "Labor"
    End If

    ' --- Roll "Tires"-tagged hours into the $50/hr row (tag STAYS "Tires") ---
    If tireRow > 0 Then
        Dim tireHrs As Double: tireHrs = 0
        For r = LINEITEM_START To lastItemRow
            If r <> tireRow And ws.Cells(r, 2).Value = "Tires" _
               And Trim(CStr(ws.Cells(r, 3).Value)) <> "" Then
                If IsNumeric(ws.Cells(r, 1).Value) Then tireHrs = tireHrs + ws.Cells(r, 1).Value
            End If
        Next r
        ws.Cells(tireRow, 1).Value = tireHrs
        ws.Cells(tireRow, 4).Value = 50
        ws.Cells(tireRow, 2).Value = "Tires"
    End If

    ' --- Amount = QTY * Unit Price for every numeric-QTY row ---
    Dim i As Long
    For i = LINEITEM_START To lastItemRow
        If Not IsMergedRow(ws, i) Then
            If IsNumeric(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value <> "" Then
                ws.Cells(i, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
            Else
                ws.Cells(i, 5).ClearContents
            End If
        End If
    Next i
End Sub

' Returns True if any cell in A:E of the given row is part of a merged area
Private Function IsMergedRow(ws As Worksheet, rowNum As Long) As Boolean
    IsMergedRow = ws.Range("A" & rowNum & ":E" & rowNum).MergeCells = True
    ' Note: MergeCells returns Null if only some cells are merged; treat that as merged too
    If IsNull(ws.Range("A" & rowNum & ":E" & rowNum).MergeCells) Then IsMergedRow = True
End Function

' ============================================================
'  UPDATE SUMMARY FORMULAS  (SPLIT: Parts / Labor / Tire Labor)
'  Uses xlWhole matching so "Labor Total" doesn't mis-match inside
'  "Tire Labor Total". Inserts only A:E rows (never touches the
'  merged tax note). Tire Labor Total = SUMIF "Tires".
' ============================================================
Public Sub UpdateFormulas()
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub
    Dim firstRow As String: firstRow = "15"
    Dim lastRow As String: lastRow = CStr(subtotalRow - 1)
    Dim rng As String: rng = "E" & firstRow & ":E" & lastRow
    Dim tagRng As String: tagRng = "B" & firstRow & ":B" & lastRow

    ' Subtotal = everything
    ws.Cells(subtotalRow, 5).Formula = "=SUM(" & rng & ")"
    ws.Cells(subtotalRow, 5).NumberFormat = "$#,##0.00"

    ' Parts Total = explicit "Parts"
    Dim fc As Range
    Set fc = ws.Columns("D").Find(What:="Parts Total", LookIn:=xlValues, LookAt:=xlWhole)
    If Not fc Is Nothing Then
        fc.Offset(0, 1).Formula = "=SUMIF(" & tagRng & ",""Parts""," & rng & ")"
        fc.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    ' Labor Total = explicit "Labor" (insert A:E row if missing)
    Dim lc As Range
    Set lc = ws.Columns("D").Find(What:="Labor Total", LookIn:=xlValues, LookAt:=xlWhole)
    If lc Is Nothing And Not fc Is Nothing Then
        ws.Range("A" & (fc.Row + 1) & ":E" & (fc.Row + 1)).Insert Shift:=xlDown
        ws.Cells(fc.Row + 1, 4).Value = "Labor Total"
        ws.Cells(fc.Row + 1, 4).Font.Bold = True
        Set lc = ws.Cells(fc.Row + 1, 4)
    End If
    If Not lc Is Nothing Then
        lc.Offset(0, 1).Formula = "=SUMIF(" & tagRng & ",""Labor""," & rng & ")"
        lc.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    ' Tire Labor Total = explicit "Tires" (insert A:E row if missing)
    Dim tc As Range
    Set tc = ws.Columns("D").Find(What:="Tire Labor Total", LookIn:=xlValues, LookAt:=xlWhole)
    If tc Is Nothing And Not lc Is Nothing Then
        ws.Range("A" & (lc.Row + 1) & ":E" & (lc.Row + 1)).Insert Shift:=xlDown
        ws.Cells(lc.Row + 1, 4).Value = "Tire Labor Total"
        ws.Cells(lc.Row + 1, 4).Font.Bold = True
        Set tc = ws.Cells(lc.Row + 1, 4)
    End If
    If Not tc Is Nothing Then
        tc.Offset(0, 1).Formula = "=SUMIF(" & tagRng & ",""Tires""," & rng & ")"
        tc.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    ' Total Invoice = Subtotal + Sales Tax  (refresh subtotal addr after inserts)
    Dim subAddr As String: subAddr = ws.Cells(FindSubtotalRow(ws), 5).Address
    Dim taxCell As Range, totalCell As Range
    Set taxCell = ws.Columns("D").Find(What:="Sales Tax", LookIn:=xlValues, LookAt:=xlPart)
    Set totalCell = ws.Columns("D").Find(What:="Total Invoice Amount", LookIn:=xlValues, LookAt:=xlPart)
    If Not totalCell Is Nothing And Not taxCell Is Nothing Then
        totalCell.Offset(0, 1).Formula = "=" & subAddr & "+" & taxCell.Offset(0, 1).Address
        totalCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If

    ' Total Due = Total Invoice - Payment/Credit
    Dim dueCell As Range, creditCell As Range
    Set dueCell = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
    Set creditCell = ws.Columns("D").Find(What:="Payment/Credit", LookIn:=xlValues, LookAt:=xlPart)
    If Not dueCell Is Nothing And Not totalCell Is Nothing Then
        Dim tA As String: tA = totalCell.Offset(0, 1).Address
        If Not creditCell Is Nothing Then
            Dim cA As String: cA = creditCell.Offset(0, 1).Address
            dueCell.Offset(0, 1).Formula = "=" & tA & "-IF(ISNUMBER(" & cA & ")," & cA & ",0)"
        Else
            dueCell.Offset(0, 1).Formula = "=" & tA
        End If
        dueCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    End If
End Sub

' ============================================================
'  FORMAT INVOICE  (consistent look for screen + PDF)
'  QTY column uses 0.## so half-hours (0.5, 1.5) show correctly.
' ============================================================
Public Sub FormatInvoice()
    Const LINEITEM_START As Long = 15
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub
    Dim lastItemRow As Long: lastItemRow = subtotalRow - 1

    ws.Range("D" & LINEITEM_START & ":E" & lastItemRow).NumberFormat = "$#,##0.00"
    ws.Range("A" & LINEITEM_START & ":A" & lastItemRow).NumberFormat = "0.##"

    With ws.Range("A14:E" & lastItemRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    Dim r As Long
    For r = LINEITEM_START To lastItemRow
        If (r - LINEITEM_START) Mod 2 = 0 Then
            ws.Range("A" & r & ":E" & r).Interior.Color = RGB(255, 255, 255)
        Else
            ws.Range("A" & r & ":E" & r).Interior.Color = RGB(244, 247, 252)
        End If
    Next r

    AlignLineItems
End Sub

' ============================================================
'  SAVE INVOICE  (validates, names file, exports PDF + xlsm)
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
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    If Not InvoiceFieldsFilled(ws) Then
        MsgBox "Please fill in all required fields before saving:" & Chr(13) & Chr(13) & _
               "  - Bill To / Ship To" & Chr(13) & _
               "  - Customer ID and Customer PO" & Chr(13) & _
               "  - Truck/Trailer, Fleet No, and Model", _
               vbExclamation, "Fields Required"
        Exit Sub
    End If

    AlignLineItems
    UpdateLineAmounts
    UpdateFormulas
    FormatInvoice

    Dim fleetNo As String, model As String, invoiceNo As String
    Dim dateStr As String, baseName As String
    fleetNo = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 4).Value))
    model = CleanFileName(Application.WorksheetFunction.Clean(ws.Cells(13, 5).Value))
    invoiceNo = Trim(CStr(ws.Cells(2, 5).Value))
    dateStr = Format(Date, "YYYYMMDD")
    baseName = dateStr & "_" & invoiceNo & "_" & model & "_" & fleetNo

    ' --- Windows ---
    If Not Application.OperatingSystem Like "*Mac*" Then
        Dim chosenPath As Variant
        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=Environ("USERPROFILE") & "\Documents\" & baseName & ".xlsm", _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save Invoice As")
        If chosenPath = False Then Exit Sub
        Dim winXlsm As String: winXlsm = CStr(chosenPath)
        Dim winPdf As String:  winPdf = Left(winXlsm, InStrRev(winXlsm, ".") - 1) & ".pdf"
        On Error GoTo WinSaveError
        ThisWorkbook.SaveCopyAs Filename:=winXlsm
        ExportPDF ws, winPdf
        OpenSavedAndCloseTemplate winXlsm     ' <-- open the new invoice, close template
        Exit Sub
WinSaveError:
        MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
        Exit Sub
    End If

    ' --- Mac: folder picker ---
    Dim baseDir As String: baseDir = "/Users/andykukuc/mnt/invoices/"
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
    Else
        Dim choice As Variant
        choice = InputBox("Choose a folder:" & Chr(13) & Chr(13) & _
            folderList & Chr(13) & "0)  (root of invoices)", "Save Location", "")
        If choice = "" Then Exit Sub
        If Not IsNumeric(choice) Then Exit Sub
        Dim choiceNum As Integer: choiceNum = CInt(choice)
        If choiceNum >= 1 And choiceNum <= folderCount Then
            saveDir = baseDir & folders(choiceNum) & "/"
        Else
            saveDir = baseDir
        End If
    End If

    Dim macXlsm As String: macXlsm = saveDir & baseName & ".xlsm"
    Dim macPdf As String:  macPdf = saveDir & baseName & ".pdf"
    On Error GoTo MacSaveError
    ThisWorkbook.SaveCopyAs Filename:=macXlsm
    ExportPDF ws, macPdf
    OpenSavedAndCloseTemplate macXlsm         ' <-- open the new invoice, close template
    Exit Sub
MacSaveError:
    MsgBox "Save failed: " & Err.Description, vbCritical, "Save Error"
End Sub

' ============================================================
'  Open the just-saved invoice, then close the template workbook
' ============================================================
Private Sub OpenSavedAndCloseTemplate(savedPath As String)
    Dim tmpl As Workbook: Set tmpl = ThisWorkbook
    Application.EnableEvents = False        ' don't let Workbook_Open auto-reset fire
    Dim wbNew As Workbook
    On Error Resume Next
    Set wbNew = Workbooks.Open(savedPath)
    On Error GoTo 0
    Application.EnableEvents = True

    If wbNew Is Nothing Then
        MsgBox "Saved, but could not reopen:" & Chr(13) & savedPath, vbExclamation, "Saved"
        Exit Sub
    End If

    wbNew.Activate
    ' Close the template WITHOUT saving (keeps template pristine for next time)
    Application.DisplayAlerts = False
    tmpl.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub

' ============================================================
'  EXPORT PDF  (full size, flows to page 2 if needed)
' ============================================================
Public Sub ExportPDF(ws As Worksheet, pdfPath As String)
    Dim subtotalRow As Long: subtotalRow = FindSubtotalRow(ws)
    Dim dueRow As Long
    Dim fc As Range
    Set fc = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
    dueRow = IIf(fc Is Nothing, subtotalRow + 6, fc.Row + 1)

    With ws.PageSetup
        .PrintArea = "A1:E" & dueRow
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .CenterHorizontally = True
        .PrintGridlines = False
        .PrintHeadings = False
    End With

    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0
End Sub

' ============================================================
'  RESET TEMPLATE
' ============================================================
Public Sub ResetTemplate(Optional skipConfirm As Boolean = False)
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
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
    ws.Cells(18, 2).Value = "Tires"        ' tag stays "Tires" for Tire Labor Total split
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

    UpdateFormulas
    FormatInvoice

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
'  PREVIEW / FIT TO ONE PAGE  (full size, flows to page 2 if needed)
' ============================================================
Public Sub FitToOnePage()
    Dim ws As Worksheet: Set ws = Worksheets("Invoice")
    UpdateLineAmounts
    UpdateFormulas
    FormatInvoice
    Dim fc As Range, lastRow As Long
    Set fc = ws.Columns("D").Find(What:="Total Due", LookIn:=xlValues, LookAt:=xlPart)
    lastRow = IIf(fc Is Nothing, ws.UsedRange.Rows.Count, fc.Row + 1)
    With ws.PageSetup
        .PrintArea = "A1:E" & lastRow
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .CenterHorizontally = True
        .PrintGridlines = False
    End With
    ws.PrintPreview
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
