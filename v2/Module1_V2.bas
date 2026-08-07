Attribute VB_Name = "Module1_V2"
Option Explicit

' Invoice Maker V2
' This module intentionally uses V2 names so it can be installed beside the
' original Module1 during testing. The original source files are unchanged.

Public bV2InternalSave As Boolean

' Item-row count that FormatInvoiceV2 last applied banding/validation for.
' Purely a repaint cache: it is re-derived from the sheet on every call, and
' the worst case if it ever went stale is cosmetic banding, never wrong data.
Private lV2FormattedItemRows As Long

Private Const INVOICE_SHEET As String = "Invoice"
Private Const LINEITEM_START As Long = 15
Private Const STANDARD_ITEM_ROWS As Long = 4
Private Const REPAIR_DESCRIPTION As String = "Repair Labor @ $80.00/hr"
Private Const TIRE_DESCRIPTION As String = "Install Tire Labor @ $50.00/hr"
Private Const REPAIR_RATE As Double = 80
Private Const TIRE_RATE As Double = 50
' Road Service is a plain qty x unit price line: no fixed rate, no rate-bucket
' roll-up, and no separate summary total. It flows into the subtotal directly.
Private Const ROADSERVICE_TAG As String = "Road Service"
Private Const TEMPLATE_TOKEN As String = "Invoice_Template_Formatted_V2"
Private Const MAC_VOLUME_ROOT As String = "/Volumes/invoices"

' ============================================================
'  PUBLIC WORKFLOW
' ============================================================

Public Sub InstallV2IntoWorkbook()
    Dim projectObject As Object
    Dim workbookComponent As Object
    Dim worksheetComponent As Object
    Dim workbookCode As String
    Dim worksheetCode As String

    If Not IsTemplateWorkbookV2(ThisWorkbook) Then
        MsgBox "V2 installation is only allowed in a workbook named " & _
               TEMPLATE_TOKEN & ".xlsm.", vbCritical, "V2 Installation Stopped"
        Exit Sub
    End If

    On Error GoTo InstallError
    Set projectObject = ThisWorkbook.VBProject
    Set workbookComponent = projectObject.VBComponents("ThisWorkbook")
    Set worksheetComponent = projectObject.VBComponents( _
        ThisWorkbook.Worksheets(INVOICE_SHEET).CodeName)

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

    ReplaceCodeModuleV2 workbookComponent.CodeModule, workbookCode
    ReplaceCodeModuleV2 worksheetComponent.CodeModule, worksheetCode
    InstallButtonsOnWorkbookV2 ThisWorkbook

    MsgBox "V2 code and buttons were installed in memory." & vbCrLf & _
           "Run the original SaveTemplate admin macro once to persist the build.", _
           vbInformation, "Invoice Maker V2 Installed"
    Exit Sub

InstallError:
    bV2InternalSave = False
    MsgBox "Excel could not install the V2 event code." & vbCrLf & vbCrLf & _
           "In Excel Preferences > Security, enable trust access to the VBA project, " & _
           "then run InstallV2IntoWorkbook again." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "V2 Installation Failed"
End Sub

Private Sub ReplaceCodeModuleV2(ByVal codeModule As Object, ByVal newCode As String)
    If codeModule.CountOfLines > 0 Then codeModule.DeleteLines 1, codeModule.CountOfLines
    codeModule.AddFromString newCode
End Sub

Public Sub InitializeInvoiceMakerV2()
    Dim startupError As String
    Dim startupSheet As Worksheet

    On Error GoTo InitError

    If IsTemplateWorkbookV2(ThisWorkbook) Then
        ResetTemplateV2 True
    Else
        ' A saved invoice copy is written with the sheet protected. FormatInvoiceV2
        ' writes NumberFormat and Interior directly and relies on its caller to
        ' unprotect first, which the change handler does but this path did not --
        ' reopening a saved invoice failed with 1004 on the first NumberFormat.
        ' ProtectInvoiceV2 below restores protection.
        Set startupSheet = InvoiceSheetV2(ThisWorkbook)
        startupSheet.Unprotect
        FormatInvoiceV2 startupSheet, True
    End If

    EnsureButtonsPresentV2 ThisWorkbook
    ProtectInvoiceV2 InvoiceSheetV2(ThisWorkbook)
    Exit Sub

InitError:
    startupError = "Error " & CStr(Err.Number) & ": " & Err.Description
    RestoreApplicationV2
    MsgBox "Invoice Maker could not finish starting." & vbCrLf & vbCrLf & _
           startupError, vbExclamation, "Invoice Maker V2"
End Sub

Public Sub SaveInvoiceV2()
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim outputBook As Workbook
    Dim invoiceRoot As String
    Dim customerFolder As String
    Dim excelFolder As String
    Dim pdfFolder As String
    Dim baseName As String
    Dim xlsmPath As String
    Dim pdfPath As String
    Dim finalXlsmPath As String
    Dim finalPdfPath As String
    Dim stageFolder As String
    Dim handoffError As String
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim saveSucceeded As Boolean

    Set sourceBook = ThisWorkbook
    Set sourceSheet = InvoiceSheetV2(sourceBook)

    If Not InvoiceFieldsFilledV2(sourceSheet) Then Exit Sub

    invoiceRoot = ResolveInvoiceRootV2(sourceBook)
    If invoiceRoot = "" Then
        MsgBox "The invoices share is not available." & vbCrLf & vbCrLf & _
               "Uruchom ponownie skrót Invoice Maker po połączeniu z siecią.", _
               vbExclamation, "Invoices Share Not Available"
        Exit Sub
    End If

    customerFolder = ChooseCustomerFolderV2(invoiceRoot)
    If customerFolder = "" Then Exit Sub

    #If Mac Then
    ' Do not inspect SMB subfolders from sandboxed Excel. The AppleScriptTask
    ' helper validates these folders while performing the final handoff.
    excelFolder = JoinPathV2(customerFolder, "Excel")
    pdfFolder = JoinPathV2(customerFolder, "PDF")
    #Else
    excelFolder = FindChildFolderV2(customerFolder, "excel")
    pdfFolder = FindChildFolderV2(customerFolder, "pdf")
    #End If
    If excelFolder = "" Or pdfFolder = "" Then
        MsgBox "This customer needs both an Excel folder and a PDF folder:" & _
               vbCrLf & vbCrLf & customerFolder, vbExclamation, _
               "Customer Folder Needs Attention"
        Exit Sub
    End If

    baseName = BuildInvoiceBaseNameV2(sourceSheet)

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo SaveError

#If Mac Then
    ' Excel for Mac is sandboxed. Build both files in Excel's private temporary
    ' directory, then let the installed AppleScriptTask helper copy them to SMB.
    ' This avoids Grant File Access prompts for every new invoice filename.
    stageFolder = GetMacStagingFolderV2()
    If stageFolder = "" Then
        Err.Raise vbObjectError + 301, , "Excel could not create its temporary invoice folder."
    End If
    GetAvailableOutputPathsV2 stageFolder, stageFolder, baseName, xlsmPath, pdfPath
#Else
    GetAvailableOutputPathsV2 excelFolder, pdfFolder, baseName, xlsmPath, pdfPath
    finalXlsmPath = xlsmPath
    finalPdfPath = pdfPath
#End If

    ' SaveCopyAs leaves the source/template unchanged. All destructive cleanup
    ' happens only in the newly created invoice copy.
    sourceBook.SaveCopyAs xlsmPath
    Set outputBook = Workbooks.Open(xlsmPath)

    PrepareOutputCopyV2 outputBook
    InstallButtonsOnWorkbookV2 outputBook

    bV2InternalSave = True
    outputBook.Save
    bV2InternalSave = False

    InvoiceSheetV2(outputBook).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

#If Mac Then
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing

    If Not MoveMacStagedFilesV2(xlsmPath, pdfPath, excelFolder, pdfFolder, _
                                baseName, finalXlsmPath, finalPdfPath, handoffError) Then
        Err.Raise vbObjectError + 302, , handoffError
    End If
#Else
    outputBook.Activate
#End If

    saveSucceeded = True
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Invoice saved successfully." & vbCrLf & vbCrLf & _
           "Excel: " & finalXlsmPath & vbCrLf & _
           "PDF: " & finalPdfPath, vbInformation, "Faktura zapisana / Invoice Saved"

    ' Last operation: close the unchanged template without saving it.
    If IsTemplateWorkbookV2(sourceBook) Then sourceBook.Close SaveChanges:=False
    Exit Sub

SaveError:
    Dim saveMessage As String
    Dim recoveredPath As String
    saveMessage = Err.Description
    bV2InternalSave = False
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    If saveSucceeded Then
        MsgBox "The invoice was saved, but the final handoff did not complete." & _
               vbCrLf & vbCrLf & saveMessage, vbExclamation, "Invoice Saved With Warning"
    ElseIf xlsmPath <> "" And FileExistsV2(xlsmPath) Then
        ' The invoice was fully built but never reached the customer folders.
        ' The staging area is $TMPDIR: invisible in Finder and cleared on
        ' restart, so a path there is not a usable recovery for this user.
        ' Ask the helper (outside the sandbox) to copy it to the Desktop.
        recoveredPath = RecoverStagedInvoiceV2(xlsmPath, pdfPath, baseName)
        If recoveredPath <> "" Then
            MsgBox "Saving to the customer folders did not finish, so the completed " & _
                   "invoice was placed on the Desktop instead:" & vbCrLf & vbCrLf & _
                   recoveredPath & vbCrLf & vbCrLf & _
                   "The customer folders were not updated. When the connection is " & _
                   "back, save it again from Invoice Maker." & vbCrLf & vbCrLf & _
                   saveMessage, vbExclamation, "Invoice Saved to Desktop"
        Else
            MsgBox "Saving did not finish and the invoice could not be copied to " & _
                   "the Desktop." & vbCrLf & _
                   "Do NOT restart the Mac - the temporary copy below disappears " & _
                   "on restart. Save the invoice again now, or copy this file first:" & _
                   vbCrLf & vbCrLf & xlsmPath & vbCrLf & vbCrLf & saveMessage, _
                   vbExclamation, "Partial Save"
        End If
    Else
        MsgBox "The invoice could not be saved. Your open invoice was not changed." & _
               vbCrLf & vbCrLf & saveMessage, vbCritical, "Save Failed"
    End If
End Sub

' Ask the AppleScriptTask helper to copy a staged-but-unsaved invoice to
' ~/Desktop/Invoice Recovery. Returns the recovered path, or "" when the
' helper is unavailable (for example an older installed version without the
' recoverInvoiceFiles handler). Never raises: this only runs inside the save
' error path, where a second failure must not mask the original message.
Private Function RecoverStagedInvoiceV2(ByVal stagedXlsm As String, _
                                        ByVal stagedPdf As String, _
                                        ByVal baseName As String) As String
    #If Mac Then
    Dim helperResult As String
    Dim resultParts() As String

    On Error GoTo RecoveryUnavailable
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "recoverInvoiceFiles", _
                                   stagedXlsm & Chr$(30) & stagedPdf & Chr$(30) & baseName)
    resultParts = Split(helperResult, Chr$(30))
    If UBound(resultParts) >= 1 Then
        If resultParts(0) = "OK" Then RecoverStagedInvoiceV2 = resultParts(1)
    End If
    Exit Function

RecoveryUnavailable:
    Err.Clear
    #End If
End Function

Public Sub NewInvoiceV2()
    If IsTemplateWorkbookV2(ThisWorkbook) Then
        ResetTemplateV2 False
        Exit Sub
    End If

    Dim invoiceRoot As String
    Dim templatePath As String
    invoiceRoot = ResolveInvoiceRootV2(ThisWorkbook)
    templatePath = JoinPathV2(invoiceRoot, TEMPLATE_TOKEN & ".xlsm")

    If invoiceRoot = "" Or Not FileExistsV2(templatePath) Then
        MsgBox "Please open " & TEMPLATE_TOKEN & ".xlsm from the invoices shortcut.", _
               vbInformation, "Start a New Invoice"
        Exit Sub
    End If

    Workbooks.Open templatePath
End Sub

Public Sub PrintInvoiceV2()
    Dim ws As Worksheet
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Set ws = InvoiceSheetV2(ThisWorkbook)

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo PrintError

    PrepareForEditingSaveV2 ws
    HideEmptyBucketsV2 ws
    ConfigurePageV2 ws
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    ws.PrintPreview
    ShowAllBucketsV2 ws
    ProtectInvoiceV2 ws
    Exit Sub

PrintError:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error Resume Next
    ShowAllBucketsV2 ws
    ProtectInvoiceV2 ws
    On Error GoTo 0
    MsgBox "Print Preview could not open." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Print Preview"
End Sub

Public Sub ResetTemplateV2(Optional ByVal skipConfirm As Boolean = False)
    Dim ws As Worksheet
    Dim subtotalRow As Long
    Dim itemRowCount As Long
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim resetStep As String
    Dim resetErrorNumber As Long
    Dim resetErrorDescription As String

    Set ws = InvoiceSheetV2(ThisWorkbook)

    If Not skipConfirm Then
        If MsgBox("Start a new invoice?" & vbCrLf & _
                  "All unsaved information on this invoice will be cleared." & vbCrLf & vbCrLf & _
                  "Rozpocząć nową fakturę?", vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Nowa faktura / New Invoice") <> vbYes Then Exit Sub
    End If

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ResetError

    resetStep = "unprotecting the invoice sheet"
    ws.Unprotect
    resetStep = "restoring hidden invoice rows"
    ShowAllBucketsV2 ws
    resetStep = "finding the Subtotal row"
    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then Err.Raise vbObjectError + 200, , "The Subtotal row was not found."

    itemRowCount = subtotalRow - LINEITEM_START

    ' Keep four formatted rows and delete every dynamically inserted extra row.
    If itemRowCount > STANDARD_ITEM_ROWS Then
        ws.Rows((LINEITEM_START + STANDARD_ITEM_ROWS) & ":" & _
                (subtotalRow - 1)).Delete Shift:=xlUp
    ElseIf itemRowCount < STANDARD_ITEM_ROWS Then
        ws.Rows(subtotalRow & ":" & _
                (subtotalRow + STANDARD_ITEM_ROWS - itemRowCount - 1)).Insert Shift:=xlDown
    End If

    resetStep = "clearing the line-item area"
    ClearValuesV2 ws.Range("A15:E18")

    ws.Cells(17, 1).Value = 0
    ws.Cells(17, 2).Value = "Labor"
    ws.Cells(17, 3).Value = REPAIR_DESCRIPTION
    ws.Cells(17, 4).Value = REPAIR_RATE
    ws.Cells(17, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"

    ws.Cells(18, 1).Value = 0
    ws.Cells(18, 2).Value = "Tires"
    ws.Cells(18, 3).Value = TIRE_DESCRIPTION
    ws.Cells(18, 4).Value = TIRE_RATE
    ws.Cells(18, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"

    resetStep = "clearing the invoice fields"
    ClearValuesV2 ws.Range("A7:B8")
    ClearValuesV2 ws.Range("C7:E8")
    ClearValuesV2 ws.Range("A11:B11")
    ClearValuesV2 ws.Range("C13:E13")
    ClearSummaryInputV2 ws, "Sales Tax"
    ClearSummaryInputV2 ws, "Payment/Credit"

    resetStep = "creating the invoice number and date"
    ws.Cells(2, 5).Value = NextInvoiceNumberV2()
    ws.Cells(3, 5).Value = Date
    ws.Cells(4, 5).Value = 1

    resetStep = "updating invoice calculations"
    UpdateLineAmountsV2 ws
    UpdateFormulasV2 ws
    resetStep = "formatting the invoice"
    ' Reset rebuilds the item rows, so the cached row count cannot be trusted.
    FormatInvoiceV2 ws, True
    resetStep = "restoring invoice rows"
    ShowAllBucketsV2 ws
    resetStep = "protecting the invoice sheet"
    ProtectInvoiceV2 ws

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    ws.Range("A7").Select
    Exit Sub

ResetError:
    resetErrorNumber = Err.Number
    resetErrorDescription = Err.Description
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error Resume Next
    ProtectInvoiceV2 ws
    On Error GoTo 0
    MsgBox "The invoice could not be reset." & vbCrLf & vbCrLf & _
           "While " & resetStep & ":" & vbCrLf & _
           "Error " & CStr(resetErrorNumber) & ": " & resetErrorDescription, _
           vbCritical, "Reset Failed"
End Sub

' ============================================================
'  EVENT-MODULE ENTRY POINTS
' ============================================================

Public Sub PrepareForEditingSaveV2(ByVal ws As Worksheet)
    ws.Unprotect
    SortLineItemsV2 ws
    AlignLineItemsV2 ws
    UpdateLineAmountsV2 ws
    UpdateFormulasV2 ws
    ' Sorting reorders rows, so banding must be reapplied unconditionally.
    FormatInvoiceV2 ws, True
    ProtectInvoiceV2 ws
End Sub

Public Sub HandleInvoiceChangeV2(ByVal ws As Worksheet, ByVal Target As Range)
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As XlCalculation
    Dim subtotalRow As Long
    Dim itemBlock As Range
    Dim r As Long
    Dim insertedRow As Boolean

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo ChangeError

    If Not Intersect(Target, ws.Range("A7")) Is Nothing Then
        Application.EnableEvents = False
        ws.Range("A8").Value = ws.Range("A7").Value
        ws.Range("C7").Value = ws.Range("A7").Value
        ws.Range("C8").Value = ws.Range("A7").Value
        GoTo ChangeDone
    End If

    If Not Intersect(Target, ws.Range("A11")) Is Nothing Then
        Application.EnableEvents = False
        ws.Range("B11").Value = ws.Range("A11").Value
        GoTo ChangeDone
    End If

    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then Exit Sub
    Set itemBlock = ws.Range("A15:D" & subtotalRow - 1)
    If Intersect(Target, itemBlock) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    ' Suppress repainting for the full recalculation pass; without this every
    ' keystroke visibly re-renders the banding and summary formatting.
    Application.ScreenUpdating = False
    ' Each formula written below would otherwise trigger a full dependency
    ' recalculation on its own. Nothing in this pass reads back a computed
    ' value, so one Calculate at the end is enough (see ChangeDone).
    Application.Calculation = xlCalculationManual
    ws.Unprotect

    For r = LINEITEM_START To subtotalRow - 1
        Select Case Trim$(CStr(ws.Cells(r, 3).Value))
            Case REPAIR_DESCRIPTION: ws.Cells(r, 4).Value = REPAIR_RATE
            Case TIRE_DESCRIPTION: ws.Cells(r, 4).Value = TIRE_RATE
        End Select
    Next r

    ' subtotalRow stays valid for these three: nothing above the subtotal row
    ' moves (UpdateFormulasV2 only ever inserts summary rows BELOW it).
    ' EnsureBlankEntryRowV2 reports whether it inserted an item row, which is
    ' the one thing that shifts the subtotal down, so the row stays accurate
    ' for the two calls after it.
    UpdateLineAmountsV2 ws, subtotalRow
    UpdateFormulasV2 ws, subtotalRow
    insertedRow = EnsureBlankEntryRowV2(ws, subtotalRow)
    If insertedRow Then subtotalRow = subtotalRow + 1
    FormatInvoiceV2 ws, insertedRow, subtotalRow
    ProtectInvoiceV2 ws, subtotalRow

ChangeDone:
    ' Release manual calculation before events, so the recalculation happens
    ' while this handler still owns the sheet.
    If Application.Calculation <> oldCalc Then
        Application.Calculation = oldCalc
        ws.Calculate
    End If
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub

ChangeError:
    If Application.Calculation <> oldCalc Then Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error Resume Next
    ProtectInvoiceV2 ws
    On Error GoTo 0
    MsgBox "Excel could not update the invoice totals." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Invoice Update"
End Sub

' ============================================================
'  BUTTONS AND SHEET PROTECTION
' ============================================================

Public Sub InstallButtonsV2()
    InstallButtonsOnWorkbookV2 ThisWorkbook
End Sub

Private Sub EnsureButtonsPresentV2(ByVal targetBook As Workbook)
    Dim ws As Worksheet
    Dim actionPrefix As String

    On Error GoTo ButtonsMissing
    Set ws = InvoiceSheetV2(targetBook)
    actionPrefix = "'" & Replace(targetBook.Name, "'", "''") & "'!"

    ws.Shapes("btnV2Save").OnAction = actionPrefix & "SaveInvoiceV2"
    ws.Shapes("btnV2Print").OnAction = actionPrefix & "PrintInvoiceV2"
    ws.Shapes("btnV2New").OnAction = actionPrefix & "NewInvoiceV2"

    On Error Resume Next
    ws.DrawingObjects("btnV2Save").PrintObject = False
    ws.DrawingObjects("btnV2Print").PrintObject = False
    ws.DrawingObjects("btnV2New").PrintObject = False
    On Error GoTo 0
    Exit Sub

ButtonsMissing:
    Err.Clear
    InstallButtonsOnWorkbookV2 targetBook
End Sub

Private Sub InstallButtonsOnWorkbookV2(ByVal targetBook As Workbook)

    Dim ws As Worksheet
    Dim actionPrefix As String
    Set ws = InvoiceSheetV2(targetBook)
    actionPrefix = "'" & Replace(targetBook.Name, "'", "''") & "'!"

    ws.Unprotect
    DeleteShapeIfPresentV2 ws, "btnV2Save"
    DeleteShapeIfPresentV2 ws, "btnV2Print"
    DeleteShapeIfPresentV2 ws, "btnV2New"

    AddButtonV2 ws, "btnV2Save", "ZAPISZ FAKTURĘ" & vbLf & "SAVE INVOICE", _
                actionPrefix & "SaveInvoiceV2", ws.Range("G2").Left, ws.Range("G2").Top, _
                RGB(45, 125, 70)
    AddButtonV2 ws, "btnV2Print", "DRUKUJ" & vbLf & "PRINT / PREVIEW", _
                actionPrefix & "PrintInvoiceV2", ws.Range("G6").Left, ws.Range("G6").Top, _
                RGB(43, 102, 176)
    AddButtonV2 ws, "btnV2New", "NOWA FAKTURA" & vbLf & "NEW INVOICE", _
                actionPrefix & "NewInvoiceV2", ws.Range("G10").Left, ws.Range("G10").Top, _
                RGB(181, 64, 64)

    ProtectInvoiceV2 ws
End Sub

Private Sub AddButtonV2(ByVal ws As Worksheet, ByVal buttonName As String, _
                        ByVal caption As String, ByVal macroName As String, _
                        ByVal buttonLeft As Double, ByVal buttonTop As Double, _
                        ByVal fillColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, 190, 58)

    With shp
        .Name = buttonName
        .OnAction = macroName
        .AlternativeText = caption
        .Placement = xlFreeFloating
        .Locked = True
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 1.5
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Name = "Arial"
        .TextFrame2.TextRange.Font.Size = 15
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' DrawingObjects.PrintObject is supported for worksheet drawing controls.
    ' The print area is also restricted to A:E as a second safety layer.
    On Error Resume Next
    ws.DrawingObjects(buttonName).PrintObject = False
    On Error GoTo 0
End Sub

Private Sub DeleteShapeIfPresentV2(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Public Sub ProtectInvoiceV2(ByVal ws As Worksheet, _
                            Optional ByVal knownSubtotalRow As Long = 0)
    Dim subtotalRow As Long
    Dim dueCell As Range
    Dim salesTaxCell As Range
    Dim creditCell As Range

    On Error GoTo ProtectionDone
    ws.Unprotect
    ws.Cells.Locked = True

    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow > LINEITEM_START Then
        ws.Range("A15:D" & subtotalRow - 1).Locked = False
        ws.Range("A15:D" & subtotalRow - 1).Interior.Color = RGB(255, 252, 220)
    End If

    ws.Range("A7:B8,C7:E8,A11:B11,C13:E13").Locked = False
    ws.Range("A7:B8,C7:E8,A11:B11,C13:E13").Interior.Color = RGB(255, 252, 220)

    Set salesTaxCell = FindLabelV2(ws, "Sales Tax", xlPart)
    If Not salesTaxCell Is Nothing Then salesTaxCell.Offset(0, 1).Locked = False
    Set creditCell = FindLabelV2(ws, "Payment/Credit", xlPart)
    If Not creditCell Is Nothing Then creditCell.Offset(0, 1).Locked = False

    ws.Protect UserInterfaceOnly:=True, AllowFormattingRows:=True, _
               AllowInsertingRows:=True, AllowDeletingRows:=True, AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells

ProtectionDone:
End Sub

' ============================================================
'  CALCULATION, SORTING, AND FORMATTING
' ============================================================

Public Function FindSubtotalRowV2(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = FindLabelV2(ws, "Subtotal", xlPart)
    If foundCell Is Nothing Then
        FindSubtotalRowV2 = 0
    Else
        FindSubtotalRowV2 = foundCell.Row
    End If
End Function

Private Function FindLabelV2(ByVal ws As Worksheet, ByVal labelText As String, _
                             ByVal matchType As XlLookAt) As Range
    Set FindLabelV2 = ws.Columns("D").Find( _
        What:=labelText, After:=ws.Cells(ws.Rows.Count, 4), LookIn:=xlValues, _
        LookAt:=matchType, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
End Function

Public Sub SortLineItemsV2(ByVal ws As Worksheet)
    Dim subtotalRow As Long
    Dim lastItemRow As Long
    Dim data() As Variant
    Dim n As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim outRow As Long
    Dim keyValue As Double
    Dim tempRow(1 To 6) As Variant
    Dim tag As String
    Dim description As String

    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow <= LINEITEM_START + 1 Then Exit Sub
    lastItemRow = subtotalRow - 1
    ReDim data(1 To lastItemRow - LINEITEM_START + 1, 1 To 6)

    For r = LINEITEM_START To lastItemRow
        If Not IsMergedRowV2(ws, r) Then
            tag = Trim$(CStr(ws.Cells(r, 2).Value))
            description = Trim$(CStr(ws.Cells(r, 3).Value))
            If tag <> "" Or description <> "" Or Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
                n = n + 1
                For c = 1 To 5
                    data(n, c) = ws.Cells(r, c).Value
                Next c
                data(n, 6) = SortKeyV2(tag, description)
            End If
        End If
    Next r

    If n <= 1 Then Exit Sub

    For i = 2 To n
        For c = 1 To 6
            tempRow(c) = data(i, c)
        Next c
        keyValue = CDbl(data(i, 6))
        j = i - 1
        Do While j >= 1
            If CDbl(data(j, 6)) <= keyValue Then Exit Do
            For c = 1 To 6
                data(j + 1, c) = data(j, c)
            Next c
            j = j - 1
        Loop
        For c = 1 To 6
            data(j + 1, c) = tempRow(c)
        Next c
    Next i

    outRow = LINEITEM_START
    For i = 1 To n
        For c = 1 To 5
            ws.Cells(outRow, c).Value = data(i, c)
        Next c
        outRow = outRow + 1
    Next i

    Do While outRow <= lastItemRow
        If Not IsMergedRowV2(ws, outRow) Then ClearValuesV2 ws.Range("A" & outRow & ":E" & outRow)
        outRow = outRow + 1
    Loop
End Sub

Private Function SortKeyV2(ByVal tag As String, ByVal description As String) As Double
    Dim result As Double
    Select Case LCase$(Trim$(tag))
        Case "parts": result = 100
        Case "labor", "install": result = 200
        Case "tires": result = 300
        Case "road service": result = 350
        Case Else: result = 400
    End Select
    If StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) = 0 Then result = result + 90
    If StrComp(description, TIRE_DESCRIPTION, vbTextCompare) = 0 Then result = result + 90
    SortKeyV2 = result
End Function

Public Sub UpdateLineAmountsV2(ByVal ws As Worksheet, _
                               Optional ByVal knownSubtotalRow As Long = 0)
    Dim subtotalRow As Long
    Dim lastItemRow As Long
    Dim repairRow As Long
    Dim tireRow As Long
    Dim laborHours As Double
    Dim tireCount As Double
    Dim r As Long
    Dim tag As String
    Dim description As String
    Dim block As Variant

    ' The change handler passes the row it already located; every worksheet
    ' Find here repeats on each keystroke otherwise.
    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow <= LINEITEM_START Then Exit Sub
    lastItemRow = subtotalRow - 1

    ' One bulk read replaces roughly eight per-row property crossings. On Mac
    ' Excel each Range access is an expensive bridge call, and this runs on
    ' every keystroke. Column indexes into block() are 1-based off column A.
    block = ws.Range("A" & LINEITEM_START & ":E" & lastItemRow).Value2

    For r = 1 To UBound(block, 1)
        If StrComp(Trim$(CStr(block(r, 2) & "")), "Install", vbTextCompare) = 0 Then
            ws.Cells(LINEITEM_START + r - 1, 2).Value = "Labor"
            block(r, 2) = "Labor"
        End If

        description = Trim$(CStr(block(r, 3) & ""))
        If StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) = 0 Then
            repairRow = LINEITEM_START + r - 1
        End If
        If StrComp(description, TIRE_DESCRIPTION, vbTextCompare) = 0 Then
            tireRow = LINEITEM_START + r - 1
        End If
    Next r

    For r = 1 To UBound(block, 1)
        tag = Trim$(CStr(block(r, 2) & ""))
        description = Trim$(CStr(block(r, 3) & ""))

        If description <> "" And IsNumeric(block(r, 1)) Then
            If LINEITEM_START + r - 1 <> repairRow And _
               StrComp(tag, "Labor", vbTextCompare) = 0 Then
                laborHours = laborHours + CDbl(block(r, 1))
            ElseIf LINEITEM_START + r - 1 <> tireRow And _
                   StrComp(tag, "Tires", vbTextCompare) = 0 Then
                tireCount = tireCount + CDbl(block(r, 1))
            End If
        End If
    Next r

    If repairRow > 0 Then
        ws.Cells(repairRow, 1).Value = laborHours
        ws.Cells(repairRow, 2).Value = "Labor"
        ws.Cells(repairRow, 4).Value = REPAIR_RATE
        block(repairRow - LINEITEM_START + 1, 1) = laborHours
    End If

    If tireRow > 0 Then
        ws.Cells(tireRow, 1).Value = tireCount
        ws.Cells(tireRow, 2).Value = "Tires"
        ws.Cells(tireRow, 4).Value = TIRE_RATE
        block(tireRow - LINEITEM_START + 1, 1) = tireCount
    End If

    ' IsMergedRowV2 must stay a live sheet check: Value2 returns Empty for a
    ' merged cell's non-anchor members and cannot report merge state.
    For r = 1 To UBound(block, 1)
        If Not IsMergedRowV2(ws, LINEITEM_START + r - 1) Then
            If IsNumeric(block(r, 1)) And Trim$(CStr(block(r, 1) & "")) <> "" Then
                ws.Cells(LINEITEM_START + r - 1, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
            Else
                ClearValuesV2 ws.Cells(LINEITEM_START + r - 1, 5)
            End If
        End If
    Next r
End Sub

Public Sub UpdateFormulasV2(ByVal ws As Worksheet, _
                            Optional ByVal knownSubtotalRow As Long = 0)
    Dim subtotalRow As Long
    Dim itemAmountRange As String
    Dim tagRange As String
    Dim partsCell As Range
    Dim laborCell As Range
    Dim tireCell As Range
    Dim taxCell As Range
    Dim totalCell As Range
    Dim dueCell As Range
    Dim creditCell As Range

    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then Exit Sub
    itemAmountRange = "E15:E" & subtotalRow - 1
    tagRange = "B15:B" & subtotalRow - 1

    ws.Cells(subtotalRow, 5).Formula = "=SUM(" & itemAmountRange & ")"
    ws.Cells(subtotalRow, 5).NumberFormat = "$#,##0.00"

    Set partsCell = FindLabelV2(ws, "Parts Total", xlWhole)
    If Not partsCell Is Nothing Then
        partsCell.Offset(0, 1).Formula = "=SUMIF(" & tagRange & ",""Parts""," & itemAmountRange & ")"
    End If

    Set laborCell = FindLabelV2(ws, "Labor Total", xlWhole)
    If laborCell Is Nothing And Not partsCell Is Nothing Then
        ws.Range("A" & partsCell.Row + 1 & ":E" & partsCell.Row + 1).Insert Shift:=xlDown
        ws.Cells(partsCell.Row + 1, 4).Value = "Labor Total"
        Set laborCell = ws.Cells(partsCell.Row + 1, 4)
    End If
    If Not laborCell Is Nothing Then
        laborCell.Offset(0, 1).Formula = "=SUMIF(" & tagRange & ",""Labor""," & itemAmountRange & ")"
    End If

    Set tireCell = FindLabelV2(ws, "Tire Labor Total", xlWhole)
    If tireCell Is Nothing And Not laborCell Is Nothing Then
        ws.Range("A" & laborCell.Row + 1 & ":E" & laborCell.Row + 1).Insert Shift:=xlDown
        ws.Cells(laborCell.Row + 1, 4).Value = "Tire Labor Total"
        Set tireCell = ws.Cells(laborCell.Row + 1, 4)
    End If
    If Not tireCell Is Nothing Then
        tireCell.Offset(0, 1).Formula = "=SUMIF(" & tagRange & ",""Tires""," & itemAmountRange & ")"
    End If

    Set taxCell = FindLabelV2(ws, "Sales Tax", xlPart)
    Set totalCell = FindLabelV2(ws, "Total Invoice Amount", xlPart)
    If Not taxCell Is Nothing And Not totalCell Is Nothing Then
        ' subtotalRow is unchanged: the summary-row inserts above happen below
        ' it, so a second worksheet Find here would return the same row.
        totalCell.Offset(0, 1).Formula = "=" & ws.Cells(subtotalRow, 5).Address & _
                                               "+" & taxCell.Offset(0, 1).Address
    End If

    Set dueCell = FindLabelV2(ws, "Total Due", xlPart)
    Set creditCell = FindLabelV2(ws, "Payment/Credit", xlPart)
    If Not dueCell Is Nothing And Not totalCell Is Nothing Then
        If creditCell Is Nothing Then
            dueCell.Offset(0, 1).Formula = "=" & totalCell.Offset(0, 1).Address
        Else
            dueCell.Offset(0, 1).Formula = "=" & totalCell.Offset(0, 1).Address & _
                "-IF(ISNUMBER(" & creditCell.Offset(0, 1).Address & ")," & _
                creditCell.Offset(0, 1).Address & ",0)"
        End If
    End If

    If Not partsCell Is Nothing Then partsCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    If Not laborCell Is Nothing Then laborCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    If Not tireCell Is Nothing Then tireCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    If Not totalCell Is Nothing Then totalCell.Offset(0, 1).NumberFormat = "$#,##0.00"
    If Not dueCell Is Nothing Then dueCell.Offset(0, 1).NumberFormat = "$#,##0.00"
End Sub

' The banding, dropdown, number formats and row heights depend only on HOW MANY
' item rows there are, not on their contents. Reapplying them on every keystroke
' cost a per-row Interior write and a full Validation teardown/rebuild for no
' visible change. Pass forceFullFormat:=True after any structural row change,
' and from the install/reset paths where the cached count cannot be trusted.
Public Sub FormatInvoiceV2(ByVal ws As Worksheet, _
                           Optional ByVal forceFullFormat As Boolean = False, _
                           Optional ByVal knownSubtotalRow As Long = 0)
    Dim subtotalRow As Long
    Dim lastItemRow As Long
    Dim dueCell As Range
    Dim rowNumber As Long
    Dim separator As String
    Dim itemRowCount As Long
    Dim whiteBand As Range
    Dim tintBand As Range

    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow <= LINEITEM_START Then Exit Sub
    lastItemRow = subtotalRow - 1
    itemRowCount = lastItemRow - LINEITEM_START + 1

    If Not forceFullFormat And itemRowCount = lV2FormattedItemRows Then Exit Sub

    ws.Range("D15:E" & lastItemRow).NumberFormat = "$#,##0.00"
    ws.Range("A15:A" & lastItemRow).NumberFormat = "0.##"
    AlignLineItemsV2 ws, subtotalRow

    ' Build the two banding groups and colour each in a single write instead of
    ' one Interior.Color crossing per row.
    For rowNumber = LINEITEM_START To lastItemRow
        If (rowNumber - LINEITEM_START) Mod 2 = 0 Then
            If whiteBand Is Nothing Then
                Set whiteBand = ws.Range("A" & rowNumber & ":E" & rowNumber)
            Else
                Set whiteBand = Union(whiteBand, ws.Range("A" & rowNumber & ":E" & rowNumber))
            End If
        Else
            If tintBand Is Nothing Then
                Set tintBand = ws.Range("A" & rowNumber & ":E" & rowNumber)
            Else
                Set tintBand = Union(tintBand, ws.Range("A" & rowNumber & ":E" & rowNumber))
            End If
        End If
    Next rowNumber
    If Not whiteBand Is Nothing Then whiteBand.Interior.Color = RGB(255, 255, 255)
    If Not tintBand Is Nothing Then tintBand.Interior.Color = RGB(244, 247, 252)

    separator = Application.International(xlListSeparator)
    ApplyLineTypeValidationV2 ws.Range("B15:B" & lastItemRow), separator

    Set dueCell = FindLabelV2(ws, "Total Due", xlPart)
    If Not dueCell Is Nothing Then ws.Rows(subtotalRow & ":" & dueCell.Row).RowHeight = 22

    lV2FormattedItemRows = itemRowCount
End Sub

Private Sub ApplyLineTypeValidationV2(ByVal targetRange As Range, ByVal separator As String)
    ' Validation.Add differs between Mac Excel releases and regional settings.
    ' A missing dropdown must never prevent an invoice from opening or resetting.
    On Error GoTo ValidationUnavailable
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Parts" & separator & "Labor" & separator & _
                                            "Tires" & separator & ROADSERVICE_TAG
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "Choose a type"
        .InputMessage = "Parts, Labor, Tires, or Road Service"
        .ShowError = True
        .ErrorTitle = "Choose from the list"
        .ErrorMessage = "Use Parts, Labor, Tires, or Road Service."
    End With
    Exit Sub

ValidationUnavailable:
    Err.Clear
End Sub

Public Sub AlignLineItemsV2(ByVal ws As Worksheet, _
                            Optional ByVal knownSubtotalRow As Long = 0)
    Dim subtotalRow As Long
    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow <= LINEITEM_START Then Exit Sub

    With ws.Range("A15:A" & subtotalRow - 1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("B15:C" & subtotalRow - 1)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
    With ws.Range("D15:E" & subtotalRow - 1)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
End Sub

' Returns True when a fresh blank entry row was inserted, which shifts the
' subtotal row down by one. Callers holding a subtotal row must add 1.
Private Function EnsureBlankEntryRowV2(ByVal ws As Worksheet, _
                                       Optional ByVal knownSubtotalRow As Long = 0) As Boolean
    Dim subtotalRow As Long
    Dim laborRow As Long
    Dim previousRow As Long
    Dim r As Long

    subtotalRow = knownSubtotalRow
    If subtotalRow = 0 Then subtotalRow = FindSubtotalRowV2(ws)
    ' Anchor on the fixed repair-rate row by its description, not on the first
    ' "Labor" tag: a user line tagged Labor above the anchor would otherwise
    ' become the insertion point.
    For r = LINEITEM_START To subtotalRow - 1
        If StrComp(Trim$(CStr(ws.Cells(r, 3).Value)), REPAIR_DESCRIPTION, vbTextCompare) = 0 Then
            laborRow = r
            Exit For
        End If
    Next r

    If laborRow < LINEITEM_START + 1 Then Exit Function
    previousRow = laborRow - 1
    ' Insert a fresh entry row only when the row above is a COMPLETE line.
    ' Checking tag+description alone fired mid-entry (users pick the type and
    ' description first), leaving a half-filled row that survived both the
    ' sort filter and the save-time cleanup as a phantom invoice line.
    If Trim$(CStr(ws.Cells(previousRow, 2).Value)) <> "" And _
       Trim$(CStr(ws.Cells(previousRow, 3).Value)) <> "" And _
       Trim$(CStr(ws.Cells(previousRow, 1).Value)) <> "" And _
       Trim$(CStr(ws.Cells(previousRow, 4).Value)) <> "" Then
        ws.Rows(laborRow).Insert Shift:=xlDown
        ClearValuesV2 ws.Range("A" & laborRow & ":E" & laborRow)
        EnsureBlankEntryRowV2 = True
    End If
End Function

Private Function IsMergedRowV2(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim mergeState As Variant
    mergeState = ws.Range("A" & rowNumber & ":E" & rowNumber).MergeCells
    If IsNull(mergeState) Then
        IsMergedRowV2 = True
    Else
        IsMergedRowV2 = CBool(mergeState)
    End If
End Function

' ============================================================
'  OUTPUT COPY AND PRINT LAYOUT
' ============================================================

' True when a row carries a type or description but is missing the quantity or
' the price, so no amount can be computed for it. The two fixed rate rows are
' never incomplete: they always hold a hardcoded rate and a quantity of 0, and
' are removed by their own zero-amount rule instead.
Private Function IsIncompleteLineV2(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim tag As String
    Dim description As String

    description = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
    If StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) = 0 Then Exit Function
    If StrComp(description, TIRE_DESCRIPTION, vbTextCompare) = 0 Then Exit Function

    tag = Trim$(CStr(ws.Cells(rowNumber, 2).Value))
    If tag = "" And description = "" Then Exit Function

    IsIncompleteLineV2 = Not (IsNumeric(ws.Cells(rowNumber, 1).Value) And _
                              IsNumeric(ws.Cells(rowNumber, 4).Value))
End Function

Private Sub PrepareOutputCopyV2(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim subtotalRow As Long
    Dim description As String

    Set ws = InvoiceSheetV2(wb)
    ws.Unprotect
    PrepareForEditingSaveV2 ws
    ws.Unprotect

    subtotalRow = FindSubtotalRowV2(ws)
    For rowNumber = subtotalRow - 1 To LINEITEM_START Step -1
        If Not IsMergedRowV2(ws, rowNumber) Then
            description = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
            If (description = REPAIR_DESCRIPTION Or description = TIRE_DESCRIPTION) And _
               Val(ws.Cells(rowNumber, 5).Value) = 0 Then
                ws.Rows(rowNumber).Delete Shift:=xlUp
            ElseIf Trim$(CStr(ws.Cells(rowNumber, 1).Value)) = "" And _
                   Trim$(CStr(ws.Cells(rowNumber, 2).Value)) = "" And _
                   description = "" Then
                ws.Rows(rowNumber).Delete Shift:=xlUp
            ElseIf IsIncompleteLineV2(ws, rowNumber) Then
                ' A half-filled row (a type and description but no quantity or
                ' price) is abandoned data, not a real charge. The entry-row
                ' gate only ever inspects the row above the repair anchor, so a
                ' row left incomplete elsewhere - fill one row, jump to another,
                ' come back - reaches here still half-filled and would print as
                ' a blank-value line on the customer's invoice.
                ws.Rows(rowNumber).Delete Shift:=xlUp
            End If
        End If
    Next rowNumber

    UpdateFormulasV2 ws
    ' Rows were deleted above, so the cached row count is stale by definition.
    FormatInvoiceV2 ws, True
    HideEmptyBucketsV2 ws
    ConfigurePageV2 ws
    ProtectInvoiceV2 ws
End Sub

Private Sub ConfigurePageV2(ByVal ws As Worksheet)
    Dim dueCell As Range
    Dim lastRow As Long
    Set dueCell = FindLabelV2(ws, "Total Due", xlPart)
    If dueCell Is Nothing Then
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Else
        lastRow = dueCell.Row + 1
    End If

    With ws.PageSetup
        .PrintArea = "A1:E" & lastRow
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
    End With
End Sub

Public Sub HideEmptyBucketsV2(ByVal ws As Worksheet)
    Dim subtotalRow As Long
    Dim repairRow As Long
    Dim tireRow As Long
    Dim repairAmount As Double
    Dim tireAmount As Double
    Dim laborTotal As Range
    Dim tireTotal As Range
    Dim r As Long

    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then Exit Sub

    For r = LINEITEM_START To subtotalRow - 1
        Select Case Trim$(CStr(ws.Cells(r, 3).Value))
            Case REPAIR_DESCRIPTION
                repairRow = r
                If IsNumeric(ws.Cells(r, 5).Value) Then repairAmount = CDbl(ws.Cells(r, 5).Value)
            Case TIRE_DESCRIPTION
                tireRow = r
                If IsNumeric(ws.Cells(r, 5).Value) Then tireAmount = CDbl(ws.Cells(r, 5).Value)
        End Select
    Next r

    Set laborTotal = FindLabelV2(ws, "Labor Total", xlWhole)
    If repairRow > 0 Then
        ws.Rows(repairRow).Hidden = (repairAmount = 0)
        If Not laborTotal Is Nothing Then ws.Rows(laborTotal.Row).Hidden = (repairAmount = 0)
    End If

    Set tireTotal = FindLabelV2(ws, "Tire Labor Total", xlWhole)
    If tireRow > 0 Then
        ws.Rows(tireRow).Hidden = (tireAmount = 0)
        If Not tireTotal Is Nothing Then ws.Rows(tireTotal.Row).Hidden = (tireAmount = 0)
    End If
End Sub

Public Sub ShowAllBucketsV2(ByVal ws As Worksheet)
    Dim dueCell As Range
    Dim lastRow As Long
    Set dueCell = FindLabelV2(ws, "Total Due", xlPart)
    If dueCell Is Nothing Then
        lastRow = 40
    Else
        lastRow = dueCell.Row
    End If
    ws.Rows("15:" & lastRow).Hidden = False
End Sub

' ============================================================
'  SAVE LOCATIONS AND NUMBERING
' ============================================================

Private Function ChooseCustomerFolderV2(ByVal invoiceRoot As String) As String
    ' A simple numbered list is deliberate: browsing an SMB folder from
    ' sandboxed Mac Excel triggers the repeated Grant File Access dialog.
    ChooseCustomerFolderV2 = ChooseCustomerFolderByNumberV2(invoiceRoot)
End Function

Private Function ChooseCustomerFolderByNumberV2(ByVal invoiceRoot As String) As String
    Dim folders() As String
    Dim folderCount As Long
    Dim entry As String
    Dim listText As String
    Dim choice As Variant
    Dim choiceNumber As Double
    Dim i As Long
    Dim candidate As String

    #If Mac Then
    Dim helperResult As String
    Dim resultParts As Variant

    On Error GoTo CustomerListError
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "listCustomerFolders", invoiceRoot)
    resultParts = Split(helperResult, Chr$(30))
    If UBound(resultParts) < 1 Or resultParts(0) <> "OK" Then
        If UBound(resultParts) >= 1 Then
            Err.Raise vbObjectError + 303, , resultParts(1)
        Else
            Err.Raise vbObjectError + 303, , "The Mac invoice helper returned no customer folders."
        End If
    End If

    folderCount = UBound(resultParts)
    ReDim folders(1 To folderCount)
    For i = 1 To folderCount
        folders(i) = CStr(resultParts(i))
    Next i
    GoTo CustomerListReady

CustomerListError:
    MsgBox "The customer list could not be read." & vbCrLf & vbCrLf & _
           "Run the Mac installer again, then retry." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Customer List Failed"
    Exit Function
    #Else
    entry = Dir(JoinPathV2(invoiceRoot, "*"), vbDirectory)
    Do While entry <> ""
        If entry <> "." And entry <> ".." And Left$(entry, 1) <> "." And _
           LCase$(entry) <> "archive" Then
            candidate = JoinPathV2(invoiceRoot, entry)
            If FolderExistsV2(candidate) Then
                folderCount = folderCount + 1
                ReDim Preserve folders(1 To folderCount)
                folders(folderCount) = entry
            End If
        End If
        entry = Dir()
    Loop
    #End If

CustomerListReady:
    If folderCount = 0 Then
        MsgBox "No customer folders were found in " & invoiceRoot, vbExclamation, _
               "No Customer Folders"
        Exit Function
    End If

    SortStringsV2 folders, folderCount
    For i = 1 To folderCount
        listText = listText & i & ")  " & folders(i) & vbCrLf
    Next i

    choice = InputBox( _
        Prompt:="Choose the customer number:" & vbCrLf & _
               "Wybierz numer klienta:" & vbCrLf & vbCrLf & listText, _
        Title:="Customer / Klient")

    If Len(Trim$(CStr(choice))) = 0 Then Exit Function
    If IsError(choice) Or IsEmpty(choice) Or Not IsNumeric(choice) Then
        MsgBox "Please choose a number from 1 to " & folderCount & ".", _
               vbExclamation, "Invalid Customer Number"
        Exit Function
    End If

    choiceNumber = CDbl(choice)
    If choiceNumber < 1 Or choiceNumber > folderCount Or _
       choiceNumber <> Fix(choiceNumber) Then
        MsgBox "Please choose a number from 1 to " & folderCount & ".", _
               vbExclamation, "Invalid Customer Number"
        Exit Function
    End If

    ChooseCustomerFolderByNumberV2 = JoinPathV2(invoiceRoot, folders(CLng(choiceNumber)))
End Function

Private Function FindChildFolderV2(ByVal parentPath As String, _
                                   ByVal requestedName As String) As String
    Dim entry As String
    Dim candidate As String
    entry = Dir(JoinPathV2(parentPath, "*"), vbDirectory)
    Do While entry <> ""
        If entry <> "." And entry <> ".." Then
            candidate = JoinPathV2(parentPath, entry)
            If FolderExistsV2(candidate) And _
               StrComp(entry, requestedName, vbTextCompare) = 0 Then
                FindChildFolderV2 = candidate
                Exit Function
            End If
        End If
        entry = Dir()
    Loop
End Function

Private Function ResolveInvoiceRootV2(ByVal wb As Workbook) As String
    Dim candidate As String
    Dim parentPath As String

    #If Mac Then
    ' Reading wb.Path does not traverse the share. Directory existence checks
    ' here would cause macOS Grant File Access prompts for the SMB hierarchy.
    If wb.Path <> "" And _
       InStr(1, wb.Name, TEMPLATE_TOKEN, vbTextCompare) > 0 Then
        ResolveInvoiceRootV2 = wb.Path
    End If
    Exit Function
    #Else
    candidate = wb.Path
    If IsInvoiceRootV2(candidate) Then
        ResolveInvoiceRootV2 = candidate
        Exit Function
    End If

    ' A saved invoice normally lives in <root>/<customer>/Excel.
    parentPath = ParentFolderV2(ParentFolderV2(candidate))
    If IsInvoiceRootV2(parentPath) Then
        ResolveInvoiceRootV2 = parentPath
        Exit Function
    End If

    candidate = JoinPathV2(Environ$("HOME"), "mnt/invoices")
    If IsInvoiceRootV2(candidate) Then
        ResolveInvoiceRootV2 = candidate
        Exit Function
    End If

    candidate = MAC_VOLUME_ROOT
    If IsInvoiceRootV2(candidate) Then ResolveInvoiceRootV2 = candidate
    #End If
End Function

Private Function IsInvoiceRootV2(ByVal candidate As String) As Boolean
    If candidate = "" Then Exit Function
    If Not FolderExistsV2(candidate) Then Exit Function
    IsInvoiceRootV2 = (Dir(JoinPathV2(candidate, "Invoice_Template_Formatted*.xlsm")) <> "")
End Function

Private Function NextInvoiceNumberV2() As String
    Randomize Timer
    NextInvoiceNumberV2 = CStr(10000 + Int(Rnd() * 90000))
End Function

Private Function BuildInvoiceBaseNameV2(ByVal ws As Worksheet) As String
    Dim invoiceDate As Variant
    Dim dateText As String
    Dim invoiceNumber As String
    Dim model As String
    Dim fleetNumber As String

    invoiceDate = ws.Cells(3, 5).Value
    If IsDate(invoiceDate) Then
        dateText = Format$(CDate(invoiceDate), "yyyymmdd")
    Else
        dateText = Format$(Date, "yyyymmdd")
    End If

    invoiceNumber = CleanFileNameV2(CStr(ws.Cells(2, 5).Value))
    model = CleanFileNameV2(CStr(ws.Cells(13, 5).Value))
    fleetNumber = CleanFileNameV2(CStr(ws.Cells(13, 4).Value))
    BuildInvoiceBaseNameV2 = dateText & "_" & invoiceNumber & "_" & model & "_" & fleetNumber
End Function

Private Sub GetAvailableOutputPathsV2(ByVal excelFolder As String, ByVal pdfFolder As String, _
                                      ByVal baseName As String, ByRef xlsmPath As String, _
                                      ByRef pdfPath As String)
    Dim suffix As Long
    Dim candidateName As String
    candidateName = baseName

    Do
        xlsmPath = JoinPathV2(excelFolder, candidateName & ".xlsm")
        pdfPath = JoinPathV2(pdfFolder, candidateName & ".pdf")
        If Not FileExistsV2(xlsmPath) And Not FileExistsV2(pdfPath) Then Exit Do
        suffix = suffix + 1
        candidateName = baseName & "-" & Format$(suffix + 1, "00")
    Loop
End Sub

Private Function GetMacStagingFolderV2() As String
#If Mac Then
    Dim temporaryRoot As String
    Dim stagingFolder As String

    On Error GoTo StageError
    temporaryRoot = Environ$("TMPDIR")
    If temporaryRoot = "" Then temporaryRoot = Environ$("HOME")
    If temporaryRoot = "" Then Exit Function

    stagingFolder = JoinPathV2(temporaryRoot, "InvoiceMakerV2")
    If Not FolderExistsV2(stagingFolder) Then MkDir stagingFolder
    If FolderExistsV2(stagingFolder) Then GetMacStagingFolderV2 = stagingFolder
    Exit Function

StageError:
    GetMacStagingFolderV2 = ""
#End If
End Function

Private Function MoveMacStagedFilesV2(ByVal stagedXlsmPath As String, _
                                      ByVal stagedPdfPath As String, _
                                      ByVal excelFolder As String, _
                                      ByVal pdfFolder As String, _
                                      ByVal baseName As String, _
                                      ByRef finalXlsmPath As String, _
                                      ByRef finalPdfPath As String, _
                                      ByRef errorMessage As String) As Boolean
#If Mac Then
    Dim delimiter As String
    Dim payload As String
    Dim helperResult As String
    Dim resultParts As Variant

    On Error GoTo HelperError
    delimiter = Chr$(30)
    payload = stagedXlsmPath & delimiter & stagedPdfPath & delimiter & _
              excelFolder & delimiter & pdfFolder & delimiter & baseName

    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "moveInvoiceFiles", payload)
    resultParts = Split(helperResult, delimiter)

    If UBound(resultParts) >= 2 And resultParts(0) = "OK" Then
        finalXlsmPath = resultParts(1)
        finalPdfPath = resultParts(2)
        MoveMacStagedFilesV2 = True
        Exit Function
    End If

    If UBound(resultParts) >= 1 Then
        errorMessage = resultParts(1)
    Else
        errorMessage = helperResult
    End If
    If errorMessage = "" Then errorMessage = "The Mac invoice helper returned no result."
    Exit Function

HelperError:
    errorMessage = "The Mac invoice helper is missing or could not run." & vbCrLf & _
                   "Run the Mac installer again, then retry." & vbCrLf & vbCrLf & _
                   Err.Description
#Else
    errorMessage = "The Mac invoice helper is only available on macOS."
#End If
End Function

Private Function InvoiceFieldsFilledV2(ByVal ws As Worksheet) As Boolean
    Dim missing As String
    If Trim$(CStr(ws.Cells(7, 1).Value)) = "" Then missing = missing & vbCrLf & "- Bill To"
    If Trim$(CStr(ws.Cells(7, 3).Value)) = "" Then missing = missing & vbCrLf & "- Ship To"
    If Trim$(CStr(ws.Cells(11, 1).Value)) = "" Then missing = missing & vbCrLf & "- Customer ID"
    If Trim$(CStr(ws.Cells(11, 2).Value)) = "" Then missing = missing & vbCrLf & "- Customer PO"
    If Trim$(CStr(ws.Cells(13, 3).Value)) = "" Then missing = missing & vbCrLf & "- Truck/Trailer"
    If Trim$(CStr(ws.Cells(13, 4).Value)) = "" Then missing = missing & vbCrLf & "- Fleet No"
    If Trim$(CStr(ws.Cells(13, 5).Value)) = "" Then missing = missing & vbCrLf & "- Model"

    If missing <> "" Then
        MsgBox "Please fill in these fields before saving:" & missing, _
               vbExclamation, "Complete the Invoice"
        Exit Function
    End If

    InvoiceFieldsFilledV2 = True
End Function

' ============================================================
'  SMALL HELPERS
' ============================================================

Private Function InvoiceSheetV2(ByVal wb As Workbook) As Worksheet
    Set InvoiceSheetV2 = wb.Worksheets(INVOICE_SHEET)
End Function

Private Function IsTemplateWorkbookV2(ByVal wb As Workbook) As Boolean
    IsTemplateWorkbookV2 = (InStr(1, wb.Name, TEMPLATE_TOKEN, vbTextCompare) > 0)
End Function

Private Sub ClearSummaryInputV2(ByVal ws As Worksheet, ByVal labelText As String)
    Dim labelCell As Range
    Set labelCell = FindLabelV2(ws, labelText, xlPart)
    If Not labelCell Is Nothing Then ClearValuesV2 labelCell.Offset(0, 1)
End Sub

Private Sub ClearValuesV2(ByVal targetRange As Range)
    Dim currentCell As Range
    Dim mergeTopLeft As Range

    ' Some Mac Excel releases raise error 1004 when ClearContents is called on
    ' a protected or merge-adjacent multi-cell range. Assigning Empty to each
    ' logical cell preserves formatting and validation without that API call.
    For Each currentCell In targetRange.Cells
        If currentCell.MergeCells Then
            Set mergeTopLeft = currentCell.MergeArea.Cells(1, 1)
            If currentCell.Address = mergeTopLeft.Address Then
                mergeTopLeft.Value = Empty
            End If
        Else
            currentCell.Value = Empty
        End If
    Next currentCell
End Sub

Private Function JoinPathV2(ByVal parentPath As String, ByVal childName As String) As String
    Dim separator As String
    If parentPath = "" Then Exit Function
    If InStr(parentPath, "/") > 0 Then
        separator = "/"
    Else
        separator = "\"
    End If

    If Right$(parentPath, 1) = "/" Or Right$(parentPath, 1) = "\" Then
        JoinPathV2 = parentPath & childName
    Else
        JoinPathV2 = parentPath & separator & childName
    End If
End Function

Private Function ParentFolderV2(ByVal childPath As String) As String
    Dim slashPosition As Long
    Dim backslashPosition As Long
    slashPosition = InStrRev(childPath, "/")
    backslashPosition = InStrRev(childPath, "\")
    If backslashPosition > slashPosition Then slashPosition = backslashPosition
    If slashPosition > 1 Then ParentFolderV2 = Left$(childPath, slashPosition - 1)
End Function

Private Function FolderExistsV2(ByVal folderPath As String) As Boolean
    Dim attributes As Long
    On Error GoTo DoesNotExist
    attributes = GetAttr(folderPath)
    FolderExistsV2 = ((attributes And vbDirectory) = vbDirectory)
DoesNotExist:
End Function

Private Function FileExistsV2(ByVal filePath As String) As Boolean
    Dim attributes As Long
    On Error GoTo DoesNotExist
    attributes = GetAttr(filePath)
    FileExistsV2 = ((attributes And vbDirectory) = 0)
DoesNotExist:
End Function

Private Function CleanFileNameV2(ByVal value As String) As String
    Dim illegalCharacters As String
    Dim index As Long
    Dim result As String
    illegalCharacters = "/\:*?""<>|" & Chr$(0)
    result = Application.WorksheetFunction.Clean(Trim$(value))
    For index = 1 To Len(illegalCharacters)
        result = Replace(result, Mid$(illegalCharacters, index, 1), "_")
    Next index
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    CleanFileNameV2 = result
End Function

Private Sub SortStringsV2(ByRef values() As String, ByVal count As Long)
    Dim i As Long
    Dim j As Long
    Dim temporary As String
    For i = 2 To count
        temporary = values(i)
        j = i - 1
        Do While j >= 1
            If StrComp(values(j), temporary, vbTextCompare) <= 0 Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = temporary
    Next i
End Sub

Private Sub RestoreApplicationV2()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
