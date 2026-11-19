Attribute VB_Name = "Module1"
Option Explicit

' Invoice Maker V2
' Production runtime module. This is the workbook's only standard VBA module.

Public bV2InternalSave As Boolean
Public bV2WorkflowBusy As Boolean

' Item-row count that FormatInvoiceV2 last applied banding/validation for.
' Purely a repaint cache: it is re-derived from the sheet on every call, and
' the worst case if it ever went stale is cosmetic banding, never wrong data.
Private lV2FormattedItemRows As Long

Private Const INVOICE_SHEET As String = "Invoice"
Private Const LINEITEM_START As Long = 15
Private Const STANDARD_ITEM_ROWS As Long = 4
' Print layout. Item rows hold one line at a compact height so a typical
' invoice prints on a single page; a description long enough to wrap grows
' only its own row instead of every row reserving room for several lines.
Private Const ITEM_ROW_HEIGHT As Double = 22
Private Const SUMMARY_ROW_HEIGHT As Double = 22
Private Const PAGE_WIDTH_INCHES As Double = 8.5
Private Const PAGE_HEIGHT_INCHES As Double = 11
Private Const PAGE_MARGIN_INCHES As Double = 0.5
' Shrink an invoice onto one page only while the text stays readable. Below
' this scale a long invoice continues onto page two (still one page wide).
Private Const MIN_ONE_PAGE_SCALE As Double = 0.75
Private Const REPAIR_DESCRIPTION As String = "Repair Labor @ $80.00/hr"
Private Const TIRE_DESCRIPTION As String = "Install Tire Labor @ $50.00/hr"
Private Const REPAIR_RATE As Double = 80
Private Const TIRE_RATE As Double = 50
' Road Service is a plain qty x unit price line: no fixed rate, no rate-bucket
' roll-up, and no separate summary total. It flows into the subtotal directly.
Private Const ROADSERVICE_TAG As String = "Road Service"
Private Const TEMPLATE_TOKEN As String = "Invoice_Template_Formatted_V2"
Private Const MAC_VOLUME_ROOT As String = "/Volumes/invoices"
Private Const MAX_LEGACY_ITEMS As Long = 250
Private Const IMPORT_TOLERANCE As Double = 0.005
Private Const QUANTITY_TOLERANCE As Double = 0.0000001
Private Const MAX_INVOICE_NUMBER_FILENAME_BYTES As Long = 40
Private Const MAX_MODEL_FILENAME_BYTES As Long = 72
Private Const MAX_FLEET_FILENAME_BYTES As Long = 72
Private Const MAX_FILENAME_BASE_BYTES As Long = 220

Private Type LegacyInvoiceDataV2
    InvoiceNumber As String
    InvoiceDate As Variant
    PageNumber As Variant
    BillToLine1 As String
    BillToLine2 As String
    ShipToLine1 As String
    ShipToLine2 As String
    CustomerID As String
    CustomerPO As String
    PaymentTerms As String
    ShipDate As Variant
    DueDate As Variant
    TruckTrailer As String
    FleetNumber As String
    Model As String
    HasSalesTax As Boolean
    SalesTax As Variant
    HasPaymentCredit As Boolean
    PaymentCredit As Variant
    HasOldSubtotal As Boolean
    OldSubtotal As Double
    HasOldTotal As Boolean
    OldTotal As Double
    HasOldDue As Boolean
    OldDue As Double
    ExpectedSubtotal As Double
    ExpectedTotal As Double
    ExpectedDue As Double
    Items As Variant
    ItemCount As Long
End Type

' ============================================================
'  PUBLIC WORKFLOW
' ============================================================

Public Sub InitializeInvoiceMakerV2()
    Dim startupSheet As Worksheet
    Dim startupError As String

    On Error GoTo InitError

    ' Old invoices can store manual calculation mode. Excel then carries that
    ' application-wide setting into the next workbook and draws a line through
    ' formula results to mark them stale. Invoice totals must always calculate.
    Application.Calculation = xlCalculationAutomatic
    Set startupSheet = InvoiceSheetV2(ThisWorkbook)

    If IsTemplateWorkbookV2(ThisWorkbook) Then
        ResetTemplateV2 True
    Else
        ' A saved invoice copy is written with the sheet protected.
        ' FormatInvoiceV2 writes NumberFormat and Interior directly and relies
        ' on its caller to unprotect first, which the change handler does but
        ' this path did not -- reopening a saved invoice failed with 1004 on
        ' the first NumberFormat. ProtectInvoiceV2 below restores protection.
        startupSheet.Unprotect
        FormatInvoiceV2 startupSheet, True
    End If

    EnsureButtonsPresentV2 ThisWorkbook
    ProtectInvoiceV2 InvoiceSheetV2(ThisWorkbook)
    If IsTemplateWorkbookV2(ThisWorkbook) Then
        NormalizeInvoiceViewV2 ThisWorkbook, startupSheet, "A7"
    Else
        NormalizeInvoiceViewV2 ThisWorkbook, startupSheet, "A1"
    End If
    Exit Sub

InitError:
    startupError = "Error " & CStr(Err.Number) & ": " & Err.Description
    RestoreApplicationV2
    MsgBox "Invoice Maker could not finish starting." & vbCrLf & vbCrLf & _
           startupError, vbExclamation, "Invoice Maker V2"
End Sub

' Administrative save used only when updating the production template itself.
' Normal users save invoices through SaveInvoiceV2; disabling events here keeps
' Workbook_BeforeSave from redirecting this maintenance save into that workflow.
Public Sub SaveTemplateV2()
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim saveError As String

    If Not IsTemplateWorkbookV2(ThisWorkbook) Then
        MsgBox "This maintenance command only saves the V2 template.", _
               vbExclamation, "Template Save Stopped"
        Exit Sub
    End If

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo SaveTemplateError

    bV2WorkflowBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    bV2WorkflowBusy = False
    MsgBox "Template saved.", vbInformation, "Template Saved"
    Exit Sub

SaveTemplateError:
    saveError = Err.Description
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    bV2WorkflowBusy = False
    MsgBox "The template could not be saved." & vbCrLf & vbCrLf & saveError, _
           vbCritical, "Template Save Failed"
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
    Dim outputSaved As Boolean
    Dim saveSucceeded As Boolean

    Set sourceBook = ThisWorkbook
    Set sourceSheet = InvoiceSheetV2(sourceBook)

    If bV2WorkflowBusy Then Exit Sub
    If Not InvoiceFieldsFilledV2(sourceSheet) Then Exit Sub

    invoiceRoot = ResolveInvoiceRootV2(sourceBook)
    If invoiceRoot = "" Then
        MsgBox "The invoices share is not available." & vbCrLf & vbCrLf & _
               UnicodeTextV2("Uruchom ponownie skr\u00F3t Invoice Maker po " & _
                             "po\u0142\u0105czeniu z sieci\u0105."), _
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

    bV2WorkflowBusy = True
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
    outputSaved = True
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
    bV2WorkflowBusy = False

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
    bV2WorkflowBusy = False

    ' A failure while preparing, saving, or exporting the staged copy must not
    ' leave a temporary invoice open beside the source workbook. Keep cleanup
    ' guarded so it cannot replace the original save error with a second one.
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    Set outputBook = Nothing
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    If saveSucceeded Then
        MsgBox "The invoice was saved, but the final handoff did not complete." & _
               vbCrLf & vbCrLf & saveMessage, vbExclamation, "Invoice Saved With Warning"
    ElseIf outputSaved And xlsmPath <> "" And FileExistsV2(xlsmPath) Then
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

' Ask the AppleScriptTask helper to copy a fully saved staged invoice to
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

' Administrator-only migration command. It intentionally has no worksheet
' button: open the Macros dialog (Control+Option+F8), select
' ImportLegacyInvoiceV2, and choose one old invoice. The legacy workbook is
' opened read-only with its macros disabled; only approved input values are
' copied into a clean SaveCopyAs of this trusted V2 template.
Public Sub ImportLegacyInvoiceV2()
    Dim selectedFile As Variant
    Dim legacyPath As String
    Dim legacyName As String
    Dim legacyBase As String
    Dim expectedFingerprint As String
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim stageBook As Workbook
    Dim stageSheet As Worksheet
    Dim stageFolder As String
    Dim stagePath As String
    Dim stagePdfPath As String
    Dim finalPath As String
    Dim finalPdfPath As String
    Dim backupPath As String
    Dim pdfBackupPath As String
    Dim pdfError As String
    Dim pdfPlaced As Boolean
    Dim helperError As String
    Dim validationError As String
    Dim comparisonText As String
    Dim recoveredPath As String
    Dim replacementOutcome As String
    Dim importErrorNumber As Long
    Dim importErrorDescription As String
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldSecurity As Long
    Dim oldStatusBar As Variant
    Dim settingsCaptured As Boolean
    Dim stageReady As Boolean
    Dim importCompleted As Boolean
    Dim invoiceData As LegacyInvoiceDataV2

    If bV2WorkflowBusy Then
        MsgBox "Another Invoice Maker operation is still running.", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If

    If Not IsTemplateWorkbookV2(ThisWorkbook) Then
        MsgBox "Run this private maintenance command from the production V2 template.", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If

    selectedFile = Application.GetOpenFilename(Title:="Choose old XLSM invoice to upgrade")
    If VarType(selectedFile) = vbBoolean Then
        If selectedFile = False Then Exit Sub
    End If

    legacyPath = Trim$(CStr(selectedFile))
    legacyName = FileNameFromPathV2(legacyPath)
    If LCase$(Right$(legacyName, 5)) <> ".xlsm" Then
        MsgBox "Choose an old macro-enabled Excel invoice (.xlsm).", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If
    If Left$(legacyName, 2) = "~$" Then
        MsgBox "That is Excel's temporary lock file, not an invoice.", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If
    If InStr(1, legacyName, TEMPLATE_TOKEN, vbTextCompare) > 0 Then
        MsgBox "A template cannot be imported as an old invoice.", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If
    If SamePathV2(legacyPath, ThisWorkbook.FullName) Then
        MsgBox "The production template cannot replace itself.", _
               vbExclamation, "Import Not Started"
        Exit Sub
    End If
    If WorkbookPathIsOpenV2(legacyPath) Then
        MsgBox "Close the old invoice in Excel, then run the import again.", _
               vbExclamation, "Old Invoice Is Open"
        Exit Sub
    End If

    If Not FingerprintFileV2(legacyPath, expectedFingerprint, helperError) Then
        MsgBox "The selected invoice could not be verified before import." & _
               vbCrLf & vbCrLf & helperError, vbCritical, "Import Not Started"
        Exit Sub
    End If

    On Error GoTo ImportError
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldSecurity = Application.AutomationSecurity
    oldStatusBar = Application.StatusBar
    settingsCaptured = True

    bV2WorkflowBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    ' 3 = msoAutomationSecurityForceDisable. A numeric value avoids a broken
    ' Office enum reference on older Mac Excel builds.
    Application.AutomationSecurity = 3
    Application.StatusBar = "Reading the old invoice safely..."

    Set sourceBook = Workbooks.Open( _
        Filename:=legacyPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True, Notify:=False)
    If sourceBook Is ThisWorkbook Or _
       SamePathV2(sourceBook.FullName, ThisWorkbook.FullName) Then
        Set sourceBook = Nothing
        Err.Raise vbObjectError + 351, , _
                  "The selected path resolves to the production template itself."
    End If
    If Not sourceBook.ReadOnly Then
        sourceBook.Close SaveChanges:=False
        Set sourceBook = Nothing
        Err.Raise vbObjectError + 352, , _
                  "Excel did not open the old invoice read-only. Nothing was imported."
    End If
    legacyPath = sourceBook.FullName
    legacyName = sourceBook.Name

    If Not TryGetInvoiceSheetV2(sourceBook, sourceSheet) Then
        Err.Raise vbObjectError + 340, , _
                  "The selected workbook does not contain an Invoice worksheet."
    End If
    If Not ReadLegacyInvoiceV2(sourceSheet, invoiceData, validationError) Then
        Err.Raise vbObjectError + 341, , validationError
    End If

    sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    Set sourceSheet = Nothing

    ' Every subsequent open is a SaveCopyAs of this trusted, already-running
    ' V2 template. Allowing that known code suppresses extra Mac security
    ' prompts; events remain disabled for the entire staging operation.
    Application.AutomationSecurity = 1
    ' The legacy workbook was deliberately read while calculation was manual.
    ' Restore automatic mode before creating the V2 package so Excel does not
    ' save calcMode="manual" and render valid totals with stale-value lines.
    Application.Calculation = xlCalculationAutomatic

    legacyBase = Left$(legacyName, Len(legacyName) - 5)
    stageFolder = GetLegacyStageFolderV2()
    If stageFolder = "" Then
        Err.Raise vbObjectError + 342, , _
                  "Excel could not create a private staging folder for the upgraded invoice."
    End If
    stagePath = JoinPathV2(stageFolder, legacyName)

    Application.StatusBar = "Building the clean V2 replacement..."
    ThisWorkbook.SaveCopyAs stagePath
    Set stageBook = Workbooks.Open( _
        Filename:=stagePath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True, Notify:=False)
    Set stageSheet = InvoiceSheetV2(stageBook)

    BuildImportedInvoiceV2 stageBook, stageSheet, invoiceData
    ValidateStableImportV2 stageSheet, invoiceData
    ValidateLegacySubtotalV2 invoiceData, stageSheet
    comparisonText = LegacyTotalComparisonV2(invoiceData, stageSheet)

    Application.ScreenUpdating = True
    If MsgBox( _
        "Ready to upgrade this invoice:" & vbCrLf & vbCrLf & _
        legacyPath & vbCrLf & vbCrLf & _
        "Invoice #: " & invoiceData.InvoiceNumber & vbCrLf & _
        "Invoice date: " & ImportedDateTextV2(invoiceData.InvoiceDate) & _
        vbCrLf & vbCrLf & comparisonText & vbCrLf & _
        "The original XLSM will be replaced at the same path and filename." & _
        vbCrLf & "A fresh one-page PDF will replace the customer's matching PDF." & _
        vbCrLf & "Verified recovery backups of both will be kept on your Desktop." & _
        vbCrLf & vbCrLf & "Continue?", _
        vbYesNo + vbExclamation + vbDefaultButton2, _
        "Replace Old Invoice With V2?") <> vbYes Then
        GoTo ImportCancelled
    End If
    Application.ScreenUpdating = False

    Application.StatusBar = "Saving and verifying the V2 replacement..."
    Application.CalculateFullRebuild
    bV2InternalSave = True
    stageBook.Save
    bV2InternalSave = False
    stageReady = True
    stageBook.Close SaveChanges:=False
    Set stageBook = Nothing
    Set stageSheet = Nothing

    ' Reopen the just-saved package with macros still force-disabled and verify
    ' its stored identity, imported fields, row count and calculated totals.
    Set stageBook = Workbooks.Open( _
        Filename:=stagePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True, Notify:=False)
    Set stageSheet = InvoiceSheetV2(stageBook)
    ValidateSavedStageV2 stageSheet, invoiceData

    ' The saved XLSM keeps its editable blank rows. Render the PDF from this
    ' read-only in-memory copy, prepared exactly as SAVE INVOICE prepares its
    ' output, and never save it. Doing this before the replacement means a PDF
    ' failure leaves the original invoice untouched.
    Application.StatusBar = "Creating the one-page PDF..."
    stagePdfPath = JoinPathV2(stageFolder, legacyBase & ".pdf")
    PrepareOutputCopyV2 stageBook
    InvoiceSheetV2(stageBook).ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=stagePdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    stageBook.Close SaveChanges:=False
    Set stageBook = Nothing
    Set stageSheet = Nothing

    If Not ReplaceLegacyInvoiceV2(stagePath, legacyPath, expectedFingerprint, _
                                  finalPath, backupPath, replacementOutcome, _
                                  helperError) Then
        Err.Raise vbObjectError + 343, , helperError
    End If
    importCompleted = True

    ' The upgraded XLSM is already in place. A PDF problem from here on is a
    ' warning, not a failed import: running the import again retries it.
    Application.StatusBar = "Placing the PDF in the customer folder..."
    pdfPlaced = PlaceImportedPdfV2(stagePdfPath, finalPath, finalPdfPath, _
                                   pdfBackupPath, pdfError)
    GoTo ImportFinish

ImportCancelled:
    On Error Resume Next
    If Not stageBook Is Nothing Then stageBook.Close SaveChanges:=False
    Set stageBook = Nothing
    If stagePath <> "" And FileExistsV2(stagePath) Then Kill stagePath
    If stagePdfPath <> "" And FileExistsV2(stagePdfPath) Then Kill stagePdfPath
    If stageFolder <> "" And FolderExistsV2(stageFolder) Then RmDir stageFolder
    On Error GoTo 0
    GoTo ImportFinish

ImportError:
    importErrorNumber = Err.Number
    importErrorDescription = Err.Description
    bV2InternalSave = False
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Not stageBook Is Nothing Then stageBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    Set stageBook = Nothing
    If Not stageReady Then
        If stagePath <> "" And FileExistsV2(stagePath) Then Kill stagePath
        If stageFolder <> "" And FolderExistsV2(stageFolder) Then RmDir stageFolder
    End If
    On Error GoTo 0

ImportFinish:
    bV2InternalSave = False
    bV2WorkflowBusy = False
    If settingsCaptured Then
        On Error Resume Next
        Application.AutomationSecurity = oldSecurity
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
        Application.StatusBar = oldStatusBar
        On Error GoTo 0
    End If

    If importCompleted Then
        On Error Resume Next
        If stagePdfPath <> "" And FileExistsV2(stagePdfPath) Then Kill stagePdfPath
        If stageFolder <> "" And FolderExistsV2(stageFolder) Then RmDir stageFolder
        On Error GoTo 0

        If pdfPlaced Then
            MsgBox "The old invoice was upgraded successfully." & vbCrLf & vbCrLf & _
                   "Updated file:" & vbCrLf & finalPath & vbCrLf & vbCrLf & _
                   "New PDF:" & vbCrLf & finalPdfPath & vbCrLf & vbCrLf & _
                   "Recovery backups:" & vbCrLf & backupPath & _
                   IIf(pdfBackupPath <> "", vbCrLf & pdfBackupPath, ""), _
                   vbInformation, "Legacy Invoice Upgraded"
        Else
            MsgBox "The invoice was upgraded, but its PDF was not updated." & _
                   vbCrLf & vbCrLf & "Updated file:" & vbCrLf & finalPath & _
                   vbCrLf & vbCrLf & "Recovery backup:" & vbCrLf & backupPath & _
                   vbCrLf & vbCrLf & "Run the import again on the same file to " & _
                   "retry the PDF." & vbCrLf & vbCrLf & pdfError, _
                   vbExclamation, "Invoice Upgraded, PDF Not Updated"
        End If
        Exit Sub
    End If

    If importErrorDescription <> "" Then
        If stageReady And stagePath <> "" And FileExistsV2(stagePath) Then
            recoveredPath = RecoverStagedInvoiceV2(stagePath, stagePdfPath, _
                                                   legacyBase & "_V2_recovery")
            If recoveredPath <> "" Then
                On Error Resume Next
                Kill stagePath
                If stagePdfPath <> "" And FileExistsV2(stagePdfPath) Then Kill stagePdfPath
                If stageFolder <> "" And FolderExistsV2(stageFolder) Then RmDir stageFolder
                On Error GoTo 0
            End If
        End If

        If replacementOutcome = "RECOVERY_REQUIRED" Then
            If recoveredPath <> "" Then
                MsgBox "The final replacement state could not be verified." & _
                       vbCrLf & "Do not edit the selected invoice until it is checked." & _
                       vbCrLf & vbCrLf & "Recovery V2 copy:" & vbCrLf & _
                       recoveredPath & vbCrLf & vbCrLf & "Verified original backup:" & _
                       vbCrLf & backupPath & vbCrLf & vbCrLf & _
                       "Error " & CStr(importErrorNumber) & ": " & _
                       importErrorDescription, vbCritical, "Import Needs Recovery"
            Else
                MsgBox "The final replacement state could not be verified." & _
                       vbCrLf & "Do not edit the selected invoice until it is checked." & _
                       vbCrLf & vbCrLf & "Verified original backup:" & vbCrLf & _
                       backupPath & vbCrLf & vbCrLf & _
                       "Error " & CStr(importErrorNumber) & ": " & _
                       importErrorDescription, vbCritical, "Import Needs Recovery"
            End If
        ElseIf recoveredPath <> "" Then
            MsgBox "The original invoice was not replaced (or was restored safely)." & _
                   vbCrLf & vbCrLf & _
                   "A completed V2 recovery copy was placed here:" & vbCrLf & _
                   recoveredPath & vbCrLf & vbCrLf & _
                   "Error " & CStr(importErrorNumber) & ": " & _
                   importErrorDescription, vbCritical, "Import Stopped Safely"
        Else
            MsgBox "The original invoice was not replaced (or was restored safely)." & _
                   vbCrLf & vbCrLf & _
                   "Error " & CStr(importErrorNumber) & ": " & _
                   importErrorDescription, vbCritical, "Import Stopped Safely"
        End If
    End If
End Sub

Private Function ReadLegacyInvoiceV2(ByVal ws As Worksheet, _
                                     ByRef invoiceData As LegacyInvoiceDataV2, _
                                     ByRef errorMessage As String) As Boolean
    Dim subtotalCell As Range
    Dim taxCell As Range
    Dim creditCell As Range
    Dim oldTotalCell As Range
    Dim oldDueCell As Range
    Dim itemValues As Variant
    Dim subtotalRow As Long
    Dim rowNumber As Long
    Dim itemCount As Long
    Dim repairAnchorCount As Long
    Dim tireAnchorCount As Long
    Dim repairAnchorQty As Double
    Dim tireAnchorQty As Double
    Dim laborDetailQty As Double
    Dim tireDetailQty As Double
    Dim rawTag As String
    Dim cleanTag As String
    Dim description As String
    Dim quantityValue As Variant
    Dim priceValue As Variant
    Dim hasPrice As Boolean
    Dim fieldError As String

    ReDim itemValues(1 To MAX_LEGACY_ITEMS, 1 To 4)

    If InStr(1, Trim$(CStr(ws.Cells(2, 4).Value2)), "Invoice", vbTextCompare) = 0 Or _
       InStr(1, Trim$(CStr(ws.Cells(3, 4).Value2)), "Date", vbTextCompare) = 0 Or _
       InStr(1, Trim$(CStr(ws.Cells(14, 1).Value2)), "QTY", vbTextCompare) = 0 Or _
       InStr(1, Trim$(CStr(ws.Cells(14, 3).Value2)), "DESCRIPTION", vbTextCompare) = 0 Then
        errorMessage = "The selected workbook does not match a supported invoice layout."
        Exit Function
    End If

    If Not FindUniqueLegacyLabelV2(ws, "Subtotal", True, True, subtotalCell, fieldError) Then
        errorMessage = fieldError
        Exit Function
    End If
    subtotalRow = subtotalCell.Row
    If subtotalRow <= LINEITEM_START Or subtotalRow > LINEITEM_START + MAX_LEGACY_ITEMS + 10 Then
        errorMessage = "The line-item area is outside the supported invoice layout."
        Exit Function
    End If

    If Not ReadLegacyTextV2(ws.Cells(2, 5), "Invoice #", 50, _
                            invoiceData.InvoiceNumber, fieldError, True) Then
        errorMessage = fieldError
        Exit Function
    End If
    If Not ReadLegacyValueV2(ws.Cells(3, 5), "Invoice date", _
                             invoiceData.InvoiceDate, fieldError, True, True) Then
        errorMessage = fieldError
        Exit Function
    End If
    If Not ReadLegacyValueV2(ws.Cells(4, 5), "Page number", _
                             invoiceData.PageNumber, fieldError, False, True) Then
        errorMessage = fieldError
        Exit Function
    End If

    If Not ReadLegacyTextV2(ws.Cells(7, 1), "Bill To line 1", 500, _
                            invoiceData.BillToLine1, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(8, 1), "Bill To line 2", 500, _
                            invoiceData.BillToLine2, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(7, 3), "Ship To line 1", 500, _
                            invoiceData.ShipToLine1, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(8, 3), "Ship To line 2", 500, _
                            invoiceData.ShipToLine2, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(11, 1), "Customer ID", 250, _
                            invoiceData.CustomerID, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(11, 2), "Customer PO", 250, _
                            invoiceData.CustomerPO, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(11, 3), "Payment Terms", 250, _
                            invoiceData.PaymentTerms, fieldError, False, True) Then GoTo FieldFailed
    If Not ReadLegacyValueV2(ws.Cells(11, 4), "Ship Date", _
                             invoiceData.ShipDate, fieldError, False, True) Then GoTo FieldFailed
    If Not ReadLegacyValueV2(ws.Cells(11, 5), "Due Date", _
                             invoiceData.DueDate, fieldError, False, True) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(13, 3), "Truck/Trailer", 250, _
                            invoiceData.TruckTrailer, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(13, 4), "Fleet No.", 250, _
                            invoiceData.FleetNumber, fieldError, False) Then GoTo FieldFailed
    If Not ReadLegacyTextV2(ws.Cells(13, 5), "Model", 250, _
                            invoiceData.Model, fieldError, False) Then GoTo FieldFailed

    If Not FindUniqueLegacyLabelV2(ws, "Sales Tax", False, False, taxCell, fieldError) Then
        errorMessage = fieldError
        Exit Function
    End If
    If Not taxCell Is Nothing Then
        If Not ReadLegacyValueV2(taxCell.Offset(0, 1), "Sales Tax", _
                                 invoiceData.SalesTax, fieldError, False, True) Then GoTo FieldFailed
        If Not IsEmpty(invoiceData.SalesTax) And Trim$(CStr(invoiceData.SalesTax)) <> "" Then
            If IsLegacyBlankMarkerV2(invoiceData.SalesTax) Then
                invoiceData.SalesTax = Empty
            ElseIf Not IsNumeric(invoiceData.SalesTax) Then
                errorMessage = "Sales Tax must be blank or numeric in the old invoice."
                Exit Function
            Else
                invoiceData.HasSalesTax = True
            End If
        End If
    End If

    If Not FindUniqueLegacyLabelV2(ws, "Payment/Credit", False, False, _
                                    creditCell, fieldError) Then
        errorMessage = fieldError
        Exit Function
    End If
    If Not creditCell Is Nothing Then
        If Not ReadLegacyValueV2(creditCell.Offset(0, 1), "Payment/Credit", _
                                 invoiceData.PaymentCredit, fieldError, False, True) Then GoTo FieldFailed
        If Not IsEmpty(invoiceData.PaymentCredit) And _
           Trim$(CStr(invoiceData.PaymentCredit)) <> "" Then
            If IsLegacyBlankMarkerV2(invoiceData.PaymentCredit) Then
                invoiceData.PaymentCredit = Empty
            ElseIf Not IsNumeric(invoiceData.PaymentCredit) Then
                errorMessage = "Payment/Credit must be blank or numeric in the old invoice."
                Exit Function
            Else
                invoiceData.HasPaymentCredit = True
            End If
        End If
    End If

    For rowNumber = LINEITEM_START To subtotalRow - 1
        If LegacyRowHasDataV2(ws, rowNumber) Then
            If IsMergedRowV2(ws, rowNumber) Then
                If IsIgnorableLegacyNoteRowV2(ws, rowNumber) Then
                    GoTo NextLegacyRow
                Else
                    errorMessage = "Merged cells were found in line-item row " & _
                                   CStr(rowNumber) & "."
                    Exit Function
                End If
            End If
            If LegacyInputRowHasFormulaV2(ws, rowNumber) Then
                errorMessage = "A formula was found in imported input cells on row " & _
                               CStr(rowNumber) & ". Only data values can be migrated."
                Exit Function
            End If

            rawTag = Trim$(CStr(ws.Cells(rowNumber, 2).Value2))
            description = Trim$(CStr(ws.Cells(rowNumber, 3).Value2))
            quantityValue = ws.Cells(rowNumber, 1).Value2
            priceValue = ws.Cells(rowNumber, 4).Value2

            If Len(description) > 1000 Then
                errorMessage = "The description on row " & CStr(rowNumber) & _
                               " is longer than 1,000 characters."
                Exit Function
            End If

            If StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) = 0 Then
                repairAnchorCount = repairAnchorCount + 1
                If repairAnchorCount > 1 Then
                    errorMessage = "More than one repair-labor rate row was found."
                    Exit Function
                End If
                If Trim$(CStr(quantityValue)) <> "" Then
                    If Not ValidLegacyNumberV2(quantityValue, True) Then
                        errorMessage = "The repair-labor quantity is invalid."
                        Exit Function
                    End If
                    repairAnchorQty = CDbl(quantityValue)
                End If
                If Trim$(CStr(priceValue)) <> "" Then
                    If Not IsNumeric(priceValue) Or _
                       Abs(CDbl(priceValue) - REPAIR_RATE) > IMPORT_TOLERANCE Then
                        errorMessage = "The repair-labor rate is not the expected $80.00."
                        Exit Function
                    End If
                End If
            ElseIf StrComp(description, TIRE_DESCRIPTION, vbTextCompare) = 0 Then
                tireAnchorCount = tireAnchorCount + 1
                If tireAnchorCount > 1 Then
                    errorMessage = "More than one tire-labor rate row was found."
                    Exit Function
                End If
                If Trim$(CStr(quantityValue)) <> "" Then
                    If Not ValidLegacyNumberV2(quantityValue, True) Then
                        errorMessage = "The tire-labor quantity is invalid."
                        Exit Function
                    End If
                    tireAnchorQty = CDbl(quantityValue)
                End If
                If Trim$(CStr(priceValue)) <> "" Then
                    If Not IsNumeric(priceValue) Or _
                       Abs(CDbl(priceValue) - TIRE_RATE) > IMPORT_TOLERANCE Then
                        errorMessage = "The tire-labor rate is not the expected $50.00."
                        Exit Function
                    End If
                End If
            Else
                If description = "" Or Not ValidLegacyNumberV2(quantityValue, True) Then
                    errorMessage = "Line-item row " & CStr(rowNumber) & _
                                   " needs a description and a valid numeric quantity."
                    Exit Function
                End If

                hasPrice = (Trim$(CStr(priceValue)) <> "")
                If hasPrice Then
                    If Not ValidLegacyNumberV2(priceValue, False) Then
                        errorMessage = "Line-item row " & CStr(rowNumber) & _
                                       " has an invalid unit price."
                        Exit Function
                    End If
                End If

                If Not CanonicalLegacyTagV2(rawTag, description, priceValue, _
                                             hasPrice, cleanTag, fieldError) Then
                    errorMessage = "Line-item row " & CStr(rowNumber) & ": " & fieldError
                    Exit Function
                End If
                If (cleanTag = "Parts" Or cleanTag = ROADSERVICE_TAG) And Not hasPrice Then
                    errorMessage = "Line-item row " & CStr(rowNumber) & _
                                   " needs a numeric unit price."
                    Exit Function
                End If

                If itemCount >= MAX_LEGACY_ITEMS Then
                    errorMessage = "The invoice contains more than " & _
                                   CStr(MAX_LEGACY_ITEMS) & " supported line items."
                    Exit Function
                End If
                itemCount = itemCount + 1
                itemValues(itemCount, 1) = CDbl(quantityValue)
                itemValues(itemCount, 2) = cleanTag
                itemValues(itemCount, 3) = description
                If cleanTag = "Parts" Or cleanTag = ROADSERVICE_TAG Then
                    itemValues(itemCount, 4) = CDbl(priceValue)
                Else
                    itemValues(itemCount, 4) = Empty
                End If

                If cleanTag = "Labor" Then laborDetailQty = laborDetailQty + CDbl(quantityValue)
                If cleanTag = "Tires" Then tireDetailQty = tireDetailQty + CDbl(quantityValue)
            End If
        End If
NextLegacyRow:
    Next rowNumber

    If repairAnchorQty > laborDetailQty + QUANTITY_TOLERANCE Then
        If Not AddLegacyAdjustmentV2(itemValues, itemCount, _
                                     repairAnchorQty - laborDetailQty, "Labor", _
                                     "Additional repair labor (migrated)", fieldError) Then
            errorMessage = fieldError
            Exit Function
        End If
    End If
    If tireAnchorQty > tireDetailQty + QUANTITY_TOLERANCE Then
        If Not AddLegacyAdjustmentV2(itemValues, itemCount, _
                                     tireAnchorQty - tireDetailQty, "Tires", _
                                     "Additional tire-install labor (migrated)", fieldError) Then
            errorMessage = fieldError
            Exit Function
        End If
    End If

    invoiceData.ItemCount = itemCount
    invoiceData.Items = itemValues

    If Not IsNumeric(subtotalCell.Offset(0, 1).Value2) Then
        errorMessage = "The old invoice does not contain a stored numeric Subtotal."
        Exit Function
    End If
    invoiceData.HasOldSubtotal = True
    invoiceData.OldSubtotal = CDbl(subtotalCell.Offset(0, 1).Value2)
    If FindUniqueLegacyLabelV2(ws, "Total Invoice Amount", False, False, _
                               oldTotalCell, fieldError) Then
        If Not oldTotalCell Is Nothing Then
            If IsNumeric(oldTotalCell.Offset(0, 1).Value2) Then
                invoiceData.HasOldTotal = True
                invoiceData.OldTotal = CDbl(oldTotalCell.Offset(0, 1).Value2)
            End If
        End If
    Else
        errorMessage = fieldError
        Exit Function
    End If
    If FindUniqueLegacyLabelV2(ws, "Total Due", False, False, _
                               oldDueCell, fieldError) Then
        If Not oldDueCell Is Nothing Then
            If IsNumeric(oldDueCell.Offset(0, 1).Value2) Then
                invoiceData.HasOldDue = True
                invoiceData.OldDue = CDbl(oldDueCell.Offset(0, 1).Value2)
            End If
        End If
    Else
        errorMessage = fieldError
        Exit Function
    End If

    ReadLegacyInvoiceV2 = True
    Exit Function

FieldFailed:
    errorMessage = fieldError
End Function

Private Sub BuildImportedInvoiceV2(ByVal targetBook As Workbook, _
                                   ByVal ws As Worksheet, _
                                   ByRef invoiceData As LegacyInvoiceDataV2)
    Dim subtotalRow As Long
    Dim currentItemRows As Long
    Dim requiredItemRows As Long
    Dim repairRow As Long
    Dim tireRow As Long
    Dim rowNumber As Long
    Dim itemIndex As Long
    Dim salesTaxCell As Range
    Dim creditCell As Range

    ws.Unprotect
    ShowAllBucketsV2 ws
    ' Repairs the known production defect before any summary inputs are copied:
    ' a missing Sales Tax row previously caused Tire Labor to be added twice.
    UpdateFormulasV2 ws

    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then
        Err.Raise vbObjectError + 344, , "The V2 template has no Subtotal row."
    End If

    requiredItemRows = invoiceData.ItemCount + 3
    If requiredItemRows < STANDARD_ITEM_ROWS Then requiredItemRows = STANDARD_ITEM_ROWS
    currentItemRows = subtotalRow - LINEITEM_START

    If currentItemRows > requiredItemRows Then
        ws.Rows((LINEITEM_START + requiredItemRows) & ":" & _
                (subtotalRow - 1)).Delete Shift:=xlUp
    ElseIf currentItemRows < requiredItemRows Then
        ws.Rows(subtotalRow & ":" & _
                (subtotalRow + requiredItemRows - currentItemRows - 1)).Insert Shift:=xlDown
    End If

    subtotalRow = FindSubtotalRowV2(ws)
    ClearValuesV2 ws.Range("A" & LINEITEM_START & ":E" & subtotalRow - 1)
    repairRow = subtotalRow - 2
    tireRow = subtotalRow - 1

    For itemIndex = 1 To invoiceData.ItemCount
        rowNumber = LINEITEM_START + itemIndex - 1
        ws.Cells(rowNumber, 1).Value2 = invoiceData.Items(itemIndex, 1)
        WriteLiteralTextV2 ws.Cells(rowNumber, 2), _
                           CStr(invoiceData.Items(itemIndex, 2))
        WriteLiteralTextV2 ws.Cells(rowNumber, 3), _
                           CStr(invoiceData.Items(itemIndex, 3))
        If Not IsEmpty(invoiceData.Items(itemIndex, 4)) Then
            ws.Cells(rowNumber, 4).Value2 = invoiceData.Items(itemIndex, 4)
        End If
    Next itemIndex

    ws.Cells(repairRow, 1).Value2 = 0
    WriteLiteralTextV2 ws.Cells(repairRow, 2), "Labor"
    WriteLiteralTextV2 ws.Cells(repairRow, 3), REPAIR_DESCRIPTION
    ws.Cells(repairRow, 4).Value2 = REPAIR_RATE

    ws.Cells(tireRow, 1).Value2 = 0
    WriteLiteralTextV2 ws.Cells(tireRow, 2), "Tires"
    WriteLiteralTextV2 ws.Cells(tireRow, 3), TIRE_DESCRIPTION
    ws.Cells(tireRow, 4).Value2 = TIRE_RATE

    WriteLiteralTextV2 ws.Cells(2, 5), invoiceData.InvoiceNumber
    WriteImportedValueV2 ws.Cells(3, 5), invoiceData.InvoiceDate
    WriteImportedValueV2 ws.Cells(4, 5), invoiceData.PageNumber
    WriteLiteralTextV2 ws.Cells(7, 1), invoiceData.BillToLine1
    WriteLiteralTextV2 ws.Cells(8, 1), invoiceData.BillToLine2
    WriteLiteralTextV2 ws.Cells(7, 3), invoiceData.ShipToLine1
    WriteLiteralTextV2 ws.Cells(8, 3), invoiceData.ShipToLine2
    WriteLiteralTextV2 ws.Cells(11, 1), invoiceData.CustomerID
    WriteLiteralTextV2 ws.Cells(11, 2), invoiceData.CustomerPO
    WriteLiteralTextV2 ws.Cells(11, 3), invoiceData.PaymentTerms
    WriteImportedValueV2 ws.Cells(11, 4), invoiceData.ShipDate
    WriteImportedValueV2 ws.Cells(11, 5), invoiceData.DueDate
    WriteLiteralTextV2 ws.Cells(13, 3), invoiceData.TruckTrailer
    WriteLiteralTextV2 ws.Cells(13, 4), invoiceData.FleetNumber
    WriteLiteralTextV2 ws.Cells(13, 5), invoiceData.Model

    Set salesTaxCell = FindLabelV2(ws, "Sales Tax", xlPart)
    If salesTaxCell Is Nothing Then
        Err.Raise vbObjectError + 345, , "The V2 Sales Tax row could not be repaired."
    End If
    ClearValuesV2 salesTaxCell.Offset(0, 1)
    If invoiceData.HasSalesTax Then
        salesTaxCell.Offset(0, 1).Value2 = CDbl(invoiceData.SalesTax)
    End If

    Set creditCell = FindLabelV2(ws, "Payment/Credit", xlPart)
    If creditCell Is Nothing Then
        Err.Raise vbObjectError + 346, , "The V2 Payment/Credit row was not found."
    End If
    ClearValuesV2 creditCell.Offset(0, 1)
    If invoiceData.HasPaymentCredit Then
        creditCell.Offset(0, 1).Value2 = CDbl(invoiceData.PaymentCredit)
    End If

    UpdateLineAmountsV2 ws
    UpdateFormulasV2 ws
    FormatInvoiceV2 ws, True
    ShowAllBucketsV2 ws
    ConfigurePageV2 ws
    InstallButtonsOnWorkbookV2 targetBook
    NormalizeInvoiceViewV2 targetBook, ws, "A1"
    ProtectInvoiceV2 ws
    ws.Calculate
End Sub

Private Sub ValidateStableImportV2(ByVal ws As Worksheet, _
                                   ByRef invoiceData As LegacyInvoiceDataV2)
    Dim firstSubtotal As Double
    Dim firstTotal As Double
    Dim firstDue As Double
    Dim secondSubtotal As Double
    Dim secondTotal As Double
    Dim secondDue As Double

    If Not TryReadSummaryNumberV2(ws, "Subtotal", firstSubtotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Invoice Amount", firstTotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Due", firstDue) Then
        Err.Raise vbObjectError + 347, , _
                  "The upgraded invoice totals could not be calculated."
    End If

    ' A second complete pass is deliberate. If any legacy value could feed a
    ' total twice, this catches the drift before the original is replaced.
    ws.Unprotect
    UpdateLineAmountsV2 ws
    UpdateFormulasV2 ws
    ProtectInvoiceV2 ws
    ws.Calculate

    If Not TryReadSummaryNumberV2(ws, "Subtotal", secondSubtotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Invoice Amount", secondTotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Due", secondDue) Then
        Err.Raise vbObjectError + 348, , _
                  "The upgraded invoice totals did not survive a second calculation."
    End If

    If Abs(firstSubtotal - secondSubtotal) > IMPORT_TOLERANCE Or _
       Abs(firstTotal - secondTotal) > IMPORT_TOLERANCE Or _
       Abs(firstDue - secondDue) > IMPORT_TOLERANCE Then
        Err.Raise vbObjectError + 349, , _
                  "The upgraded totals changed during the verification pass. " & _
                  "The original was not replaced."
    End If

    invoiceData.ExpectedSubtotal = secondSubtotal
    invoiceData.ExpectedTotal = secondTotal
    invoiceData.ExpectedDue = secondDue
End Sub

Private Sub ValidateSavedStageV2(ByVal ws As Worksheet, _
                                 ByRef invoiceData As LegacyInvoiceDataV2)
    Dim subtotalRow As Long
    Dim rowNumber As Long
    Dim importedItemCount As Long
    Dim description As String
    Dim storedSubtotal As Double
    Dim storedTotal As Double
    Dim storedDue As Double

    If StrComp(CStr(ws.Cells(2, 5).Value2), invoiceData.InvoiceNumber, _
               vbBinaryCompare) <> 0 Or _
       Not ImportedValuesMatchV2(ws.Cells(3, 5).Value2, invoiceData.InvoiceDate) Or _
       StrComp(CStr(ws.Cells(7, 1).Value2), invoiceData.BillToLine1, _
               vbBinaryCompare) <> 0 Or _
       StrComp(CStr(ws.Cells(11, 1).Value2), invoiceData.CustomerID, _
               vbBinaryCompare) <> 0 Or _
       StrComp(CStr(ws.Cells(13, 3).Value2), invoiceData.TruckTrailer, _
               vbBinaryCompare) <> 0 Or _
       StrComp(CStr(ws.Cells(13, 4).Value2), invoiceData.FleetNumber, _
               vbBinaryCompare) <> 0 Or _
       StrComp(CStr(ws.Cells(13, 5).Value2), invoiceData.Model, _
               vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 355, , _
                  "The saved V2 copy did not preserve its invoice identity fields."
    End If

    subtotalRow = FindSubtotalRowV2(ws)
    For rowNumber = LINEITEM_START To subtotalRow - 1
        description = Trim$(CStr(ws.Cells(rowNumber, 3).Value2))
        If description <> "" And _
           StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) <> 0 And _
           StrComp(description, TIRE_DESCRIPTION, vbTextCompare) <> 0 Then
            importedItemCount = importedItemCount + 1
        End If
    Next rowNumber
    If importedItemCount <> invoiceData.ItemCount Then
        Err.Raise vbObjectError + 356, , _
                  "The saved V2 copy does not contain the expected number of line items."
    End If

    If Not TryReadSummaryNumberV2(ws, "Subtotal", storedSubtotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Invoice Amount", storedTotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Due", storedDue) Then
        Err.Raise vbObjectError + 357, , _
                  "The saved V2 copy does not contain readable calculated totals."
    End If
    If Abs(storedSubtotal - invoiceData.ExpectedSubtotal) > IMPORT_TOLERANCE Or _
       Abs(storedTotal - invoiceData.ExpectedTotal) > IMPORT_TOLERANCE Or _
       Abs(storedDue - invoiceData.ExpectedDue) > IMPORT_TOLERANCE Then
        Err.Raise vbObjectError + 358, , _
                  "The saved V2 copy did not retain its verified totals."
    End If
End Sub

Private Function ImportedValuesMatchV2(ByVal storedValue As Variant, _
                                       ByVal importedValue As Variant) As Boolean
    If IsNumeric(storedValue) And IsNumeric(importedValue) Then
        ImportedValuesMatchV2 = _
            (Abs(CDbl(storedValue) - CDbl(importedValue)) < 0.0000001)
    Else
        ImportedValuesMatchV2 = _
            (StrComp(CStr(storedValue), CStr(importedValue), vbBinaryCompare) = 0)
    End If
End Function

Private Sub ValidateLegacySubtotalV2(ByRef invoiceData As LegacyInvoiceDataV2, _
                                    ByVal ws As Worksheet)
    Dim rebuiltSubtotal As Double
    If Not invoiceData.HasOldSubtotal Or _
       Not TryReadSummaryNumberV2(ws, "Subtotal", rebuiltSubtotal) Then
        Err.Raise vbObjectError + 353, , _
                  "The old and rebuilt subtotals could not be compared."
    End If

    If Abs(invoiceData.OldSubtotal - rebuiltSubtotal) > IMPORT_TOLERANCE Then
        Err.Raise vbObjectError + 354, , _
                  "The line-item data does not reproduce the old Subtotal." & _
                  vbCrLf & vbCrLf & "Old Subtotal: " & _
                  Format$(invoiceData.OldSubtotal, "$#,##0.00") & vbCrLf & _
                  "Rebuilt Subtotal: " & Format$(rebuiltSubtotal, "$#,##0.00") & _
                  vbCrLf & vbCrLf & _
                  "The original was not replaced. Review this invoice manually."
    End If
End Sub

Private Function LegacyTotalComparisonV2(ByRef invoiceData As LegacyInvoiceDataV2, _
                                         ByVal ws As Worksheet) As String
    Dim newSubtotal As Double
    Dim newTotal As Double
    Dim newDue As Double
    Dim comparisonText As String
    Dim changedText As String

    If Not TryReadSummaryNumberV2(ws, "Subtotal", newSubtotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Invoice Amount", newTotal) Or _
       Not TryReadSummaryNumberV2(ws, "Total Due", newDue) Then
        Err.Raise vbObjectError + 350, , "The upgraded totals are unavailable."
    End If

    comparisonText = "TOTAL CHECK" & vbCrLf & _
                     "Subtotal: " & LegacyAmountTextV2(invoiceData.HasOldSubtotal, _
                                                        invoiceData.OldSubtotal) & _
                     "  ->  " & Format$(newSubtotal, "$#,##0.00") & vbCrLf & _
                     "Invoice total: " & LegacyAmountTextV2(invoiceData.HasOldTotal, _
                                                            invoiceData.OldTotal) & _
                     "  ->  " & Format$(newTotal, "$#,##0.00") & vbCrLf & _
                     "Total due: " & LegacyAmountTextV2(invoiceData.HasOldDue, _
                                                        invoiceData.OldDue) & _
                     "  ->  " & Format$(newDue, "$#,##0.00")

    If (invoiceData.HasOldSubtotal And _
        Abs(invoiceData.OldSubtotal - newSubtotal) > IMPORT_TOLERANCE) Or _
       (invoiceData.HasOldTotal And _
        Abs(invoiceData.OldTotal - newTotal) > IMPORT_TOLERANCE) Or _
       (invoiceData.HasOldDue And _
        Abs(invoiceData.OldDue - newDue) > IMPORT_TOLERANCE) Then
        changedText = vbCrLf & vbCrLf & _
                      "The amounts differ. Review them carefully. This may be the " & _
                      "legacy double-calculation defect being corrected."
    End If

    LegacyTotalComparisonV2 = comparisonText & changedText
End Function

Private Function FindUniqueLegacyLabelV2(ByVal ws As Worksheet, _
                                         ByVal labelText As String, _
                                         ByVal exactMatch As Boolean, _
                                         ByVal required As Boolean, _
                                         ByRef foundCell As Range, _
                                         ByRef errorMessage As String) As Boolean
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim candidateText As String
    Dim isMatch As Boolean
    Dim matchCount As Long

    Set foundCell = Nothing
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow > 1000 Then lastRow = 1000

    For rowNumber = 1 To lastRow
        candidateText = Trim$(CStr(ws.Cells(rowNumber, 4).Value2))
        If exactMatch Then
            isMatch = (StrComp(candidateText, labelText, vbTextCompare) = 0)
        Else
            isMatch = (InStr(1, candidateText, labelText, vbTextCompare) > 0)
        End If
        If isMatch Then
            matchCount = matchCount + 1
            Set foundCell = ws.Cells(rowNumber, 4)
        End If
    Next rowNumber

    If matchCount > 1 Then
        errorMessage = "The old invoice contains duplicate '" & labelText & "' rows."
        Set foundCell = Nothing
        Exit Function
    End If
    If required And matchCount = 0 Then
        errorMessage = "The old invoice is missing its '" & labelText & "' row."
        Exit Function
    End If
    FindUniqueLegacyLabelV2 = True
End Function

Private Function ReadLegacyTextV2(ByVal sourceCell As Range, _
                                  ByVal fieldName As String, _
                                  ByVal maximumLength As Long, _
                                  ByRef resultText As String, _
                                  ByRef errorMessage As String, _
                                  ByVal required As Boolean, _
                                  Optional ByVal allowCachedFormula As Boolean = False) As Boolean
    If CellHasFormulaV2(sourceCell) And Not allowCachedFormula Then
        errorMessage = fieldName & " contains a formula; only stored data values can be imported."
        Exit Function
    End If

    If fieldName = "Invoice #" Then
        resultText = Trim$(CStr(sourceCell.Text))
        If InStr(resultText, "#") > 0 Or _
           InStr(1, resultText, "E+", vbTextCompare) > 0 Then
            resultText = Trim$(CStr(sourceCell.Value2))
        End If
    Else
        resultText = Trim$(CStr(sourceCell.Value2))
    End If
    If required And resultText = "" Then
        errorMessage = fieldName & " is blank in the old invoice."
        Exit Function
    End If
    If Len(resultText) > maximumLength Then
        errorMessage = fieldName & " is longer than " & CStr(maximumLength) & " characters."
        Exit Function
    End If
    ReadLegacyTextV2 = True
End Function

Private Function ReadLegacyValueV2(ByVal sourceCell As Range, _
                                   ByVal fieldName As String, _
                                   ByRef resultValue As Variant, _
                                   ByRef errorMessage As String, _
                                   ByVal required As Boolean, _
                                   Optional ByVal allowCachedFormula As Boolean = False) As Boolean
    If CellHasFormulaV2(sourceCell) And Not allowCachedFormula Then
        errorMessage = fieldName & " contains a formula; only stored data values can be imported."
        Exit Function
    End If
    resultValue = sourceCell.Value2
    If required And Trim$(CStr(resultValue)) = "" Then
        errorMessage = fieldName & " is blank in the old invoice."
        Exit Function
    End If
    ReadLegacyValueV2 = True
End Function

Private Function CellHasFormulaV2(ByVal sourceCell As Range) As Boolean
    Dim formulaState As Variant
    formulaState = sourceCell.HasFormula
    If Not IsNull(formulaState) Then CellHasFormulaV2 = CBool(formulaState)
End Function

Private Function LegacyRowHasDataV2(ByVal ws As Worksheet, _
                                    ByVal rowNumber As Long) As Boolean
    LegacyRowHasDataV2 = _
        (Trim$(CStr(ws.Cells(rowNumber, 1).Value2)) <> "" Or _
         Trim$(CStr(ws.Cells(rowNumber, 2).Value2)) <> "" Or _
         Trim$(CStr(ws.Cells(rowNumber, 3).Value2)) <> "" Or _
         Trim$(CStr(ws.Cells(rowNumber, 4).Value2)) <> "")
End Function

Private Function IsIgnorableLegacyNoteRowV2(ByVal ws As Worksheet, _
                                            ByVal rowNumber As Long) As Boolean
    Dim noteText As String
    noteText = Trim$(CStr(ws.Cells(rowNumber, 1).Value2))
    If InStr(1, noteText, "Note:", vbTextCompare) <> 1 Then Exit Function
    IsIgnorableLegacyNoteRowV2 = _
        (Trim$(CStr(ws.Cells(rowNumber, 2).Value2)) = "" And _
         Trim$(CStr(ws.Cells(rowNumber, 3).Value2)) = "" And _
         Trim$(CStr(ws.Cells(rowNumber, 4).Value2)) = "")
End Function

Private Function LegacyInputRowHasFormulaV2(ByVal ws As Worksheet, _
                                            ByVal rowNumber As Long) As Boolean
    Dim columnNumber As Long
    For columnNumber = 1 To 4
        If CellHasFormulaV2(ws.Cells(rowNumber, columnNumber)) Then
            LegacyInputRowHasFormulaV2 = True
            Exit Function
        End If
    Next columnNumber
End Function

Private Function ValidLegacyNumberV2(ByVal valueToCheck As Variant, _
                                     ByVal nonNegative As Boolean) As Boolean
    Dim numericValue As Double
    If Not IsNumeric(valueToCheck) Then Exit Function
    numericValue = CDbl(valueToCheck)
    If nonNegative And numericValue < 0 Then Exit Function
    If Abs(numericValue) > 10000000# Then Exit Function
    ValidLegacyNumberV2 = True
End Function

Private Function IsLegacyBlankMarkerV2(ByVal valueToCheck As Variant) As Boolean
    Dim markerText As String
    markerText = LCase$(Trim$(CStr(valueToCheck)))
    Select Case markerText
        Case "n/a", "na", "none", "-"
            IsLegacyBlankMarkerV2 = True
    End Select
End Function

Private Function CanonicalLegacyTagV2(ByVal rawTag As String, _
                                      ByVal description As String, _
                                      ByVal priceValue As Variant, _
                                      ByVal hasPrice As Boolean, _
                                      ByRef cleanTag As String, _
                                      ByRef errorMessage As String) As Boolean
    Dim lowerTag As String
    Dim lowerDescription As String
    lowerTag = LCase$(Trim$(rawTag))
    lowerDescription = LCase$(Trim$(description))

    Select Case lowerTag
        Case "parts"
            cleanTag = "Parts"
        Case "road service"
            cleanTag = ROADSERVICE_TAG
        Case "labor", "install", "repair"
            cleanTag = "Labor"
        Case "tires", "tire"
            ' VBA And is eager. Keep the conversion inside the validated-price
            ' branch so a blank Variant can never reach CDbl unexpectedly.
            If hasPrice Then
                If Abs(CDbl(priceValue)) > IMPORT_TOLERANCE Then
                    ' A tire with a unit price is merchandise, not installation
                    ' hours. Classifying it as Parts prevents the fixed $50 roll-up
                    ' from turning a $200 tire into 200 hours of tire labor.
                    cleanTag = "Parts"
                Else
                    cleanTag = "Tires"
                End If
            Else
                cleanTag = "Tires"
            End If
        Case Else
            If hasPrice Then
                ' Older invoices used item names such as Supplies or Air Bag in
                ' the type column. A priced item remains a normal taxable part.
                cleanTag = "Parts"
            ElseIf InStr(1, lowerTag & " " & lowerDescription, _
                          "tire", vbTextCompare) > 0 Then
                cleanTag = "Tires"
            ElseIf InStr(1, lowerTag & " " & lowerDescription, _
                          "labor", vbTextCompare) > 0 Or _
                   InStr(1, lowerTag & " " & lowerDescription, _
                          "repair", vbTextCompare) > 0 Or _
                   InStr(1, lowerTag & " " & lowerDescription, _
                          "install", vbTextCompare) > 0 Then
                cleanTag = "Labor"
            Else
                errorMessage = "unknown unpriced item type '" & rawTag & "'."
                Exit Function
            End If
    End Select

    CanonicalLegacyTagV2 = True
End Function

Private Function AddLegacyAdjustmentV2(ByRef itemValues As Variant, _
                                       ByRef itemCount As Long, _
                                       ByVal quantityValue As Double, _
                                       ByVal cleanTag As String, _
                                       ByVal description As String, _
                                       ByRef errorMessage As String) As Boolean
    If itemCount >= MAX_LEGACY_ITEMS Then
        errorMessage = "The migrated invoice needs more than " & _
                       CStr(MAX_LEGACY_ITEMS) & " line items."
        Exit Function
    End If
    itemCount = itemCount + 1
    itemValues(itemCount, 1) = quantityValue
    itemValues(itemCount, 2) = cleanTag
    itemValues(itemCount, 3) = description
    itemValues(itemCount, 4) = Empty
    AddLegacyAdjustmentV2 = True
End Function

Private Sub WriteLiteralTextV2(ByVal targetCell As Range, ByVal textValue As String)
    ' Text format is applied before Value2 so strings beginning with '=' remain
    ' data and can never become formulas in the upgraded workbook.
    targetCell.NumberFormat = "@"
    targetCell.Value2 = textValue
End Sub

Private Sub WriteImportedValueV2(ByVal targetCell As Range, _
                                 ByVal importedValue As Variant)
    If IsEmpty(importedValue) Or Trim$(CStr(importedValue)) = "" Then
        ClearValuesV2 targetCell
    ElseIf VarType(importedValue) = vbString Then
        WriteLiteralTextV2 targetCell, CStr(importedValue)
    Else
        targetCell.Value2 = importedValue
    End If
End Sub

Private Function TryReadSummaryNumberV2(ByVal ws As Worksheet, _
                                        ByVal labelText As String, _
                                        ByRef resultValue As Double) As Boolean
    Dim labelCell As Range
    Set labelCell = FindLabelV2(ws, labelText, xlPart)
    If labelCell Is Nothing Then Exit Function
    If Not IsNumeric(labelCell.Offset(0, 1).Value2) Then Exit Function
    resultValue = CDbl(labelCell.Offset(0, 1).Value2)
    TryReadSummaryNumberV2 = True
End Function

Private Function LegacyAmountTextV2(ByVal hasAmount As Boolean, _
                                    ByVal amountValue As Double) As String
    If hasAmount Then
        LegacyAmountTextV2 = Format$(amountValue, "$#,##0.00")
    Else
        LegacyAmountTextV2 = "not stored"
    End If
End Function

Private Function ImportedDateTextV2(ByVal importedValue As Variant) As String
    Dim serialNumber As Double

    If IsDate(importedValue) Then
        ImportedDateTextV2 = Format$(CDate(importedValue), "mm/dd/yyyy")
    ElseIf IsNumeric(importedValue) Then
        ' VBA And is eager, so CDbl must not share a condition with IsNumeric.
        serialNumber = CDbl(importedValue)
        If serialNumber > 0 And serialNumber < 2958466# Then
            ImportedDateTextV2 = Format$(DateSerial(1899, 12, 30) + _
                                         serialNumber, "mm/dd/yyyy")
        Else
            ImportedDateTextV2 = CStr(importedValue)
        End If
    Else
        ImportedDateTextV2 = CStr(importedValue)
    End If
End Function

Private Function TryGetInvoiceSheetV2(ByVal wb As Workbook, _
                                      ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(INVOICE_SHEET)
    On Error GoTo 0
    TryGetInvoiceSheetV2 = Not ws Is Nothing
End Function

Private Function WorkbookPathIsOpenV2(ByVal workbookPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If SamePathV2(wb.FullName, workbookPath) Then
            WorkbookPathIsOpenV2 = True
            Exit Function
        End If
    Next wb
End Function

Private Function SamePathV2(ByVal firstPath As String, _
                            ByVal secondPath As String) As Boolean
    Dim normalizedFirst As String
    Dim normalizedSecond As String
    normalizedFirst = Replace(Trim$(firstPath), "\", "/")
    normalizedSecond = Replace(Trim$(secondPath), "\", "/")
    Do While Len(normalizedFirst) > 1 And Right$(normalizedFirst, 1) = "/"
        normalizedFirst = Left$(normalizedFirst, Len(normalizedFirst) - 1)
    Loop
    Do While Len(normalizedSecond) > 1 And Right$(normalizedSecond, 1) = "/"
        normalizedSecond = Left$(normalizedSecond, Len(normalizedSecond) - 1)
    Loop
    SamePathV2 = (StrComp(normalizedFirst, normalizedSecond, vbTextCompare) = 0)
End Function

Private Function FileNameFromPathV2(ByVal filePath As String) As String
    Dim slashPosition As Long
    Dim backslashPosition As Long
    slashPosition = InStrRev(filePath, "/")
    backslashPosition = InStrRev(filePath, "\")
    If backslashPosition > slashPosition Then slashPosition = backslashPosition
    If slashPosition > 0 Then
        FileNameFromPathV2 = Mid$(filePath, slashPosition + 1)
    Else
        FileNameFromPathV2 = filePath
    End If
End Function

Private Function GetLegacyStageFolderV2() As String
    Dim stagingRoot As String
    Dim candidateFolder As String
    Dim attemptNumber As Long

    stagingRoot = GetMacStagingFolderV2()
    If stagingRoot = "" Then Exit Function
    Randomize Timer
    On Error GoTo StageFolderError
    For attemptNumber = 1 To 20
        candidateFolder = JoinPathV2(stagingRoot, _
            "LegacyImport_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
            Format$(Int(Rnd() * 1000000#), "000000"))
        If Not FolderExistsV2(candidateFolder) Then
            MkDir candidateFolder
            If FolderExistsV2(candidateFolder) Then
                GetLegacyStageFolderV2 = candidateFolder
                Exit Function
            End If
        End If
    Next attemptNumber
StageFolderError:
End Function

Private Function FingerprintFileV2(ByVal filePath As String, _
                                   ByRef fingerprint As String, _
                                   ByRef errorMessage As String) As Boolean
#If Mac Then
    Dim helperResult As String
    Dim resultParts As Variant
    On Error GoTo HelperError
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "fingerprintFile", filePath)
    resultParts = Split(helperResult, Chr$(30))
    If UBound(resultParts) >= 1 And resultParts(0) = "OK" Then
        fingerprint = resultParts(1)
        FingerprintFileV2 = (fingerprint <> "")
        Exit Function
    End If
    If UBound(resultParts) >= 1 Then errorMessage = resultParts(1)
    If errorMessage = "" Then errorMessage = "The Mac helper returned no fingerprint."
    Exit Function
HelperError:
    errorMessage = "The Mac invoice helper is missing or out of date." & _
                   vbCrLf & "Run the Mac installer again, then retry." & _
                   vbCrLf & vbCrLf & Err.Description
#Else
    errorMessage = "Legacy replacement is currently supported on macOS only."
#End If
End Function

Private Function ReplaceLegacyInvoiceV2(ByVal stagedPath As String, _
                                        ByVal originalPath As String, _
                                        ByVal expectedFingerprint As String, _
                                        ByRef finalPath As String, _
                                        ByRef backupPath As String, _
                                        ByRef replacementOutcome As String, _
                                        ByRef errorMessage As String) As Boolean
#If Mac Then
    Dim delimiter As String
    Dim payload As String
    Dim helperResult As String
    Dim resultParts As Variant

    On Error GoTo HelperError
    delimiter = Chr$(30)
    payload = stagedPath & delimiter & originalPath & delimiter & expectedFingerprint
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "replaceLegacyInvoice", payload)
    resultParts = Split(helperResult, delimiter)
    If UBound(resultParts) >= 2 And resultParts(0) = "OK" Then
        finalPath = resultParts(1)
        backupPath = resultParts(2)
        ReplaceLegacyInvoiceV2 = True
        Exit Function
    End If
    If UBound(resultParts) >= 2 Then backupPath = resultParts(2)
    If UBound(resultParts) >= 1 Then errorMessage = resultParts(1)
    Select Case resultParts(0)
        Case "ERROR_UNTOUCHED"
            replacementOutcome = "UNTOUCHED"
        Case "ERROR_RESTORED"
            replacementOutcome = "RESTORED"
        Case Else
            replacementOutcome = "RECOVERY_REQUIRED"
    End Select
    If errorMessage = "" Then errorMessage = "The Mac helper returned no replacement result."
    Exit Function
HelperError:
    replacementOutcome = "RECOVERY_REQUIRED"
    errorMessage = "The Mac invoice helper is missing or could not finish the " & _
                   "verified replacement." & vbCrLf & vbCrLf & Err.Description
#Else
    errorMessage = "Legacy replacement is currently supported on macOS only."
#End If
End Function

' Hands the rendered PDF of an upgraded invoice to the Mac helper, which puts
' it in the customer's PDF folder under the invoice's own base name, backing up
' any existing PDF first. Returns False with errorMessage set; never raises.
Private Function PlaceImportedPdfV2(ByVal stagedPdfPath As String, _
                                    ByVal invoicePath As String, _
                                    ByRef finalPdfPath As String, _
                                    ByRef pdfBackupPath As String, _
                                    ByRef errorMessage As String) As Boolean
#If Mac Then
    Dim delimiter As String
    Dim helperResult As String
    Dim resultParts As Variant

    On Error GoTo HelperError
    delimiter = Chr$(30)
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", "placeImportedPdf", _
                                   stagedPdfPath & delimiter & invoicePath)
    resultParts = Split(helperResult, delimiter)
    If UBound(resultParts) >= 2 And resultParts(0) = "OK" Then
        finalPdfPath = Trim$(CStr(resultParts(1)))
        pdfBackupPath = Trim$(CStr(resultParts(2)))
        PlaceImportedPdfV2 = (finalPdfPath <> "")
        If Not PlaceImportedPdfV2 Then errorMessage = "The Mac helper returned no PDF path."
        Exit Function
    End If
    If UBound(resultParts) >= 1 Then errorMessage = resultParts(1)
    If errorMessage = "" Then errorMessage = "The Mac helper returned no PDF result."
    Exit Function
HelperError:
    errorMessage = "The Mac invoice helper is missing or out of date (it needs " & _
                   "placeImportedPdf). Run the Mac installer again." & _
                   vbCrLf & vbCrLf & Err.Description
#Else
    errorMessage = "PDF replacement is currently supported on macOS only."
#End If
End Function

Public Sub PrintInvoiceV2()
    Dim ws As Worksheet
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim printerRestoreToken As String
    Dim printerError As String
    Dim printerDefaultsChanged As Boolean
    Dim printDialogErrorNumber As Long
    Dim printErrorNumber As Long
    Dim printErrorDescription As String

    If bV2WorkflowBusy Then Exit Sub
    Set ws = InvoiceSheetV2(ThisWorkbook)

    bV2WorkflowBusy = True
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo PrintError

    PrepareForEditingSaveV2 ws
    ' SAVE INVOICE deletes these rows from its output copy. Hide the same rows
    ' temporarily here so paper and the archived PDF paginate identically while
    ' leaving the editable workbook untouched after the print dialog closes.
    HideNonPrintableItemRowsV2 ws
    HideEmptyBucketsV2 ws
    ConfigurePageV2 ws

    If Not BeginOneSidedPrintingV2(printerRestoreToken, printerError) Then
        Err.Raise vbObjectError + 304, , printerError
    End If
    printerDefaultsChanged = True

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    ' Show the real printer dialog, not Print Preview. The Mac helper has
    ' temporarily changed every configured queue to one-sided, so whichever
    ' printer the user chooses starts with duplex disabled.
    On Error Resume Next
    Application.Dialogs(xlDialogPrint).Show
    printDialogErrorNumber = Err.Number
    Err.Clear
    On Error GoTo PrintError

    RestorePrintingDefaultsV2 printerRestoreToken
    printerDefaultsChanged = False

    ' Excel for Mac reports Cancel from its native Print dialog as error 1004
    ' instead of returning False. Treat that exact result as a normal cancel.
    If printDialogErrorNumber <> 0 And printDialogErrorNumber <> 1004 Then
        Err.Raise printDialogErrorNumber, , "The native Print dialog failed."
    End If

PrintDone:
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ShowAllBucketsV2 ws
    ProtectInvoiceV2 ws
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    bV2WorkflowBusy = False
    Exit Sub

PrintError:
    printErrorNumber = Err.Number
    printErrorDescription = Err.Description
    ' Errors raised while an error handler is active escape to the caller. Arm
    ' best-effort cleanup and release the busy gate before touching Excel state.
    On Error Resume Next
    bV2WorkflowBusy = False
    If printerDefaultsChanged Then RestorePrintingDefaultsV2 printerRestoreToken
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ShowAllBucketsV2 ws
    ProtectInvoiceV2 ws
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox "The invoice could not be printed." & vbCrLf & vbCrLf & _
           "Error " & CStr(printErrorNumber) & ": " & printErrorDescription, _
           vbExclamation, "Print Failed"
End Sub

Private Function BeginOneSidedPrintingV2(ByRef restoreToken As String, _
                                         ByRef errorMessage As String) As Boolean
    #If Mac Then
    Dim helperResult As String
    Dim resultParts As Variant

    On Error GoTo HelperError
    helperResult = AppleScriptTask("InvoiceMakerV2.applescript", _
                                   "beginOneSidedPrinting", "")
    resultParts = Split(helperResult, Chr$(30))
    If UBound(resultParts) >= 0 And resultParts(0) = "OK" Then
        restoreToken = helperResult
        BeginOneSidedPrintingV2 = True
        Exit Function
    End If

    If UBound(resultParts) >= 1 Then errorMessage = resultParts(1)
    If errorMessage = "" Then errorMessage = "The Mac printer helper returned no result."
    Exit Function

HelperError:
    errorMessage = "One-sided printing could not be prepared." & vbCrLf & _
                   "Run the Mac installer again, then retry." & vbCrLf & vbCrLf & _
                   Err.Description
    #Else
    BeginOneSidedPrintingV2 = True
    #End If
End Function

Private Sub RestorePrintingDefaultsV2(ByVal restoreToken As String)
    #If Mac Then
    If restoreToken = "" Then Exit Sub
    On Error Resume Next
    AppleScriptTask "InvoiceMakerV2.applescript", _
                    "restorePrintingDefaults", restoreToken
    On Error GoTo 0
    #End If
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
                  UnicodeTextV2("Rozpocz\u0105\u0107 now\u0105 faktur\u0119?"), _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
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
    ws.Cells(3, 5).NumberFormat = "m/d/yy"
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
    NormalizeInvoiceViewV2 ThisWorkbook, ws, "A7"
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

    If bV2WorkflowBusy Then Exit Sub

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
    Dim buttonName As Variant

    On Error GoTo ButtonsMissing
    Set ws = InvoiceSheetV2(targetBook)
    actionPrefix = "'" & Replace(targetBook.Name, "'", "''") & "'!"

    ws.Shapes("btnV2Save").OnAction = actionPrefix & "SaveInvoiceV2"
    ws.Shapes("btnV2Print").OnAction = actionPrefix & "PrintInvoiceV2"
    ws.Shapes("btnV2New").OnAction = actionPrefix & "NewInvoiceV2"
    If IsTemplateWorkbookV2(targetBook) Then
        ws.Shapes("btnV2Import").OnAction = actionPrefix & "ImportLegacyInvoiceV2"
    Else
        DeleteShapeIfPresentV2 ws, "btnV2Import"
    End If

    ' Rewrite every caption on open. Workbooks built by earlier imports of this
    ' module carry garbled Polish button text that nothing else repairs.
    For Each buttonName In ButtonNamesV2(targetBook)
        ws.Shapes(buttonName).AlternativeText = ButtonCaptionV2(CStr(buttonName))
        ws.Shapes(buttonName).TextFrame2.TextRange.Text = ButtonCaptionV2(CStr(buttonName))
    Next buttonName

    On Error Resume Next
    For Each buttonName In ButtonNamesV2(targetBook)
        ws.DrawingObjects(buttonName).PrintObject = False
    Next buttonName
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
    DeleteShapeIfPresentV2 ws, "btnV2Import"

    AddButtonV2 ws, "btnV2Save", ButtonCaptionV2("btnV2Save"), _
                actionPrefix & "SaveInvoiceV2", ws.Range("G2").Left, ws.Range("G2").Top, _
                RGB(45, 125, 70)
    AddButtonV2 ws, "btnV2Print", ButtonCaptionV2("btnV2Print"), _
                actionPrefix & "PrintInvoiceV2", ws.Range("G6").Left, ws.Range("G6").Top, _
                RGB(43, 102, 176)
    AddButtonV2 ws, "btnV2New", ButtonCaptionV2("btnV2New"), _
                actionPrefix & "NewInvoiceV2", ws.Range("G10").Left, ws.Range("G10").Top, _
                RGB(181, 64, 64)
    ' The importer runs only from the production template. On a saved invoice
    ' the button could do nothing but refuse, so it is not added there.
    If IsTemplateWorkbookV2(targetBook) Then
        AddButtonV2 ws, "btnV2Import", ButtonCaptionV2("btnV2Import"), _
                    actionPrefix & "ImportLegacyInvoiceV2", ws.Range("G14").Left, _
                    ws.Range("G14").Top, RGB(96, 80, 140)
    End If

    ProtectInvoiceV2 ws
End Sub

Private Function ButtonNamesV2(ByVal targetBook As Workbook) As Variant
    If IsTemplateWorkbookV2(targetBook) Then
        ButtonNamesV2 = Array("btnV2Save", "btnV2Print", "btnV2New", "btnV2Import")
    Else
        ButtonNamesV2 = Array("btnV2Save", "btnV2Print", "btnV2New")
    End If
End Function

' One source for every button label, so installing and the open-time refresh
' always agree. Polish letters go through UnicodeTextV2; see its comment.
Private Function ButtonCaptionV2(ByVal buttonName As String) As String
    Select Case buttonName
        Case "btnV2Save"
            ButtonCaptionV2 = UnicodeTextV2("ZAPISZ FAKTUR\u0118") & vbLf & "SAVE INVOICE"
        Case "btnV2Print"
            ButtonCaptionV2 = "DRUKUJ" & vbLf & "PRINT"
        Case "btnV2New"
            ButtonCaptionV2 = "NOWA FAKTURA" & vbLf & "NEW INVOICE"
        Case "btnV2Import"
            ButtonCaptionV2 = UnicodeTextV2("IMPORTUJ FAKTUR\u0118") & vbLf & "IMPORT INVOICE"
    End Select
End Function

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

    ' Protection is an operational guardrail: if this fails, let the caller's
    ' normal error path notify the user instead of silently leaving formulas
    ' and labels editable.
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
        ' Road Service sorts after the Tires group's +90 rate anchor (390).
        ' At 350 it landed between Tires detail lines and their rate row.
        Case "road service": result = 400
        Case Else: result = 500
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
    Dim currentRow As Long
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
        currentRow = LINEITEM_START + r - 1
        If Not IsMergedRowV2(ws, currentRow) Then
            tag = Trim$(CStr(block(r, 2) & ""))
            If currentRow <> repairRow And currentRow <> tireRow And _
               IsRollupTagV2(tag) Then
                ' Labor and Tires detail rows contribute hours only. Their
                ' fixed-rate anchor row is the sole billed amount; clearing D/E
                ' makes repeated recalculation idempotent and prevents doubling.
                ClearValuesV2 ws.Cells(currentRow, 4)
                ClearValuesV2 ws.Cells(currentRow, 5)
            ElseIf IsNumeric(block(r, 1)) And _
                   Trim$(CStr(block(r, 1) & "")) <> "" Then
                ws.Cells(currentRow, 5).FormulaR1C1 = "=RC[-4]*RC[-1]"
            Else
                ClearValuesV2 ws.Cells(currentRow, 5)
            End If
        End If
    Next r
End Sub

Private Function IsRollupTagV2(ByVal tag As String) As Boolean
    Select Case LCase$(Trim$(tag))
        Case "labor", "install", "tires"
            IsRollupTagV2 = True
    End Select
End Function

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
    Dim summaryAnchor As Range

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
    If taxCell Is Nothing Then
        If Not tireCell Is Nothing Then
            Set summaryAnchor = tireCell
        ElseIf Not laborCell Is Nothing Then
            Set summaryAnchor = laborCell
        Else
            Set summaryAnchor = partsCell
        End If

        If Not summaryAnchor Is Nothing Then
            ws.Range("A" & summaryAnchor.Row + 1 & ":E" & _
                     summaryAnchor.Row + 1).Insert Shift:=xlDown
            ClearValuesV2 ws.Range("A" & summaryAnchor.Row + 1 & ":E" & _
                                  summaryAnchor.Row + 1)
            ws.Cells(summaryAnchor.Row + 1, 4).Value = "Sales Tax"
            Set taxCell = ws.Cells(summaryAnchor.Row + 1, 4)
        End If
    End If

    Set totalCell = FindLabelV2(ws, "Total Invoice Amount", xlPart)
    If Not totalCell Is Nothing Then
        ' Summary rows are below the subtotal, so inserting a missing label
        ' never changes subtotalRow. Always replace the formula to eliminate
        ' legacy totals that accidentally added Tire Labor a second time.
        If taxCell Is Nothing Then
            totalCell.Offset(0, 1).Formula = "=" & ws.Cells(subtotalRow, 5).Address
        Else
            totalCell.Offset(0, 1).Formula = "=" & ws.Cells(subtotalRow, 5).Address & _
                "+IF(ISNUMBER(" & taxCell.Offset(0, 1).Address & ")," & _
                taxCell.Offset(0, 1).Address & ",0)"
        End If
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
    If Not taxCell Is Nothing Then taxCell.Offset(0, 1).NumberFormat = "$#,##0.00"
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

    ' Inserted worksheet rows inherit the Subtotal row's borders and font on
    ' Mac Excel. Reapply the known-good first item-row format across the whole
    ' item area before banding so migrated invoices retain the template layout.
    ws.Range("A15:E15").Copy
    ws.Range("A15:E" & lastItemRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("D15:E" & lastItemRow).NumberFormat = "$#,##0.00"
    ' "0.##" printed whole quantities with a trailing point ("9."). General
    ' shows 9 and 1.5 both as typed.
    ws.Range("A15:A" & lastItemRow).NumberFormat = "General"
    AlignLineItemsV2 ws, subtotalRow
    ' After AlignLineItemsV2: AutoFit needs WrapText already set on B:C.
    SizeItemRowsV2 ws, lastItemRow

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
    If Not dueCell Is Nothing Then ws.Rows(subtotalRow & ":" & dueCell.Row).RowHeight = SUMMARY_ROW_HEIGHT
    lV2FormattedItemRows = itemRowCount
End Sub

' One line of text per item row at a fixed compact height. AutoFit lets a
' wrapped description take the extra lines it needs; the floor keeps short
' rows from collapsing to Excel's bare text height.
Private Sub SizeItemRowsV2(ByVal ws As Worksheet, ByVal lastItemRow As Long)
    Dim rowNumber As Long

    With ws.Rows(LINEITEM_START & ":" & lastItemRow)
        .RowHeight = ITEM_ROW_HEIGHT
        .AutoFit
    End With
    For rowNumber = LINEITEM_START To lastItemRow
        If ws.Rows(rowNumber).RowHeight < ITEM_ROW_HEIGHT Then
            ws.Rows(rowNumber).RowHeight = ITEM_ROW_HEIGHT
        End If
    Next rowNumber
End Sub

' Workbook view state is stored inside each XLSM. A template saved while zoomed
' into the item area otherwise makes every imported invoice reopen at that same
' zoom and scroll position, which looks like damaged formatting even though the
' underlying styles are correct.
Private Sub NormalizeInvoiceViewV2(ByVal wb As Workbook, _
                                   ByVal ws As Worksheet, _
                                   ByVal activeAddress As String)
    On Error Resume Next
    wb.Activate
    ws.Activate
    ws.Range(activeAddress).Select
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.Zoom = 100
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    On Error GoTo 0
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
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("B15:C" & subtotalRow - 1)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("D15:E" & subtotalRow - 1)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
End Sub

' Returns True when a fresh blank entry row was inserted, which shifts the
' subtotal row down by one. Callers holding a subtotal row must add 1.
Private Function EnsureBlankEntryRowV2(ByVal ws As Worksheet, _
                                       Optional ByVal knownSubtotalRow As Long = 0) As Boolean
    Dim subtotalRow As Long
    Dim laborRow As Long
    Dim previousRow As Long
    Dim previousTag As String
    Dim priceComplete As Boolean
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
    previousTag = Trim$(CStr(ws.Cells(previousRow, 2).Value))
    priceComplete = IsRollupTagV2(previousTag) Or _
                    Trim$(CStr(ws.Cells(previousRow, 4).Value)) <> ""
    ' Insert a fresh entry row only when the row above is a COMPLETE line.
    ' Checking tag+description alone fired mid-entry (users pick the type and
    ' description first), leaving a half-filled row that survived both the
    ' sort filter and the save-time cleanup as a phantom invoice line.
    If previousTag <> "" And _
       Trim$(CStr(ws.Cells(previousRow, 3).Value)) <> "" And _
       Trim$(CStr(ws.Cells(previousRow, 1).Value)) <> "" And _
       priceComplete Then
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

' True when a user row cannot produce a valid charge. Labor and Tires detail
' rows require a type, description, and numeric quantity but no price because
' their hours are billed once through the fixed-rate anchor row.
Private Function IsIncompleteLineV2(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim tag As String
    Dim description As String

    description = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
    If StrComp(description, REPAIR_DESCRIPTION, vbTextCompare) = 0 Then Exit Function
    If StrComp(description, TIRE_DESCRIPTION, vbTextCompare) = 0 Then Exit Function

    tag = Trim$(CStr(ws.Cells(rowNumber, 2).Value))
    If tag = "" And description = "" Then
        ' A quantity or price by itself is abandoned input, not a printable
        ' charge. A completely empty row remains a normal blank entry row.
        IsIncompleteLineV2 = _
            (Trim$(CStr(ws.Cells(rowNumber, 1).Value)) <> "" Or _
             Trim$(CStr(ws.Cells(rowNumber, 4).Value)) <> "")
        Exit Function
    End If

    If tag = "" Or description = "" Or _
       Not IsNumeric(ws.Cells(rowNumber, 1).Value) Then
        IsIncompleteLineV2 = True
    ElseIf Not IsRollupTagV2(tag) Then
        IsIncompleteLineV2 = Not IsNumeric(ws.Cells(rowNumber, 4).Value)
    End If
End Function

Private Function IsNonPrintableItemRowV2(ByVal ws As Worksheet, _
                                         ByVal rowNumber As Long) As Boolean
    Dim description As String

    description = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
    If description = REPAIR_DESCRIPTION Or description = TIRE_DESCRIPTION Then
        IsNonPrintableItemRowV2 = (Val(ws.Cells(rowNumber, 5).Value) = 0)
    ElseIf Trim$(CStr(ws.Cells(rowNumber, 1).Value)) = "" And _
           Trim$(CStr(ws.Cells(rowNumber, 2).Value)) = "" And _
           description = "" And _
           Trim$(CStr(ws.Cells(rowNumber, 4).Value)) = "" Then
        IsNonPrintableItemRowV2 = True
    Else
        IsNonPrintableItemRowV2 = IsIncompleteLineV2(ws, rowNumber)
    End If
End Function

Private Sub HideNonPrintableItemRowsV2(ByVal ws As Worksheet)
    Dim subtotalRow As Long
    Dim rowNumber As Long

    subtotalRow = FindSubtotalRowV2(ws)
    If subtotalRow = 0 Then Exit Sub

    For rowNumber = LINEITEM_START To subtotalRow - 1
        If Not IsMergedRowV2(ws, rowNumber) Then
            If IsNonPrintableItemRowV2(ws, rowNumber) Then
                ws.Rows(rowNumber).Hidden = True
            End If
        End If
    Next rowNumber
End Sub

Private Sub PrepareOutputCopyV2(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim subtotalRow As Long

    Set ws = InvoiceSheetV2(wb)
    ws.Unprotect
    PrepareForEditingSaveV2 ws
    ws.Unprotect

    subtotalRow = FindSubtotalRowV2(ws)
    For rowNumber = subtotalRow - 1 To LINEITEM_START Step -1
        If Not IsMergedRowV2(ws, rowNumber) Then
            If IsNonPrintableItemRowV2(ws, rowNumber) Then
                ' Match the temporary paper-print cleanup: blank entry rows,
                ' unused rate anchors, and abandoned partial entries do not
                ' belong in the permanent customer copy.
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
        .LeftMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .RightMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .TopMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .BottomMargin = Application.InchesToPoints(PAGE_MARGIN_INCHES)
        .Zoom = False
        .FitToPagesWide = 1
        If FitsOnePageReadablyV2(ws.Range("A1:E" & lastRow)) Then
            .FitToPagesTall = 1
        Else
            .FitToPagesTall = False
        End If
        .CenterHorizontally = True
        .PrintGridlines = False
        ' Paper and PDF are always black and white: Excel drops cell fills and
        ' prints text and borders black, so white-on-navy headings stay legible
        ' and nothing uses colour ink. The on-screen invoice keeps its colours.
        .BlackAndWhite = True
    End With
End Sub

' True when squeezing the print area onto one sheet keeps it readable.
' Range.Height already excludes hidden rows, so the empty rate rows hidden
' for printing do not count against the page.
Private Function FitsOnePageReadablyV2(ByVal printRange As Range) As Boolean
    Dim widthScale As Double
    Dim heightScale As Double

    If printRange.Width <= 0 Or printRange.Height <= 0 Then Exit Function
    widthScale = Application.InchesToPoints(PAGE_WIDTH_INCHES - 2 * PAGE_MARGIN_INCHES) / _
                 printRange.Width
    If widthScale > 1 Then widthScale = 1
    heightScale = Application.InchesToPoints(PAGE_HEIGHT_INCHES - 2 * PAGE_MARGIN_INCHES) / _
                  printRange.Height

    ' Fitting one page wide already sets widthScale. If the height fits at
    ' that scale, one page costs nothing extra; otherwise allow the extra
    ' shrink only down to the readable floor.
    FitsOnePageReadablyV2 = (heightScale >= widthScale) Or _
                            (heightScale >= MIN_ONE_PAGE_SCALE)
End Function

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

    ReDim folders(1 To UBound(resultParts))
    For i = 1 To UBound(resultParts)
        candidate = Trim$(CStr(resultParts(i)))
        ' Only accept a nonempty leaf folder name. This prevents a malformed or
        ' outdated helper reply from resolving JoinPathV2(root, "") to the
        ' invoices share root.
        If candidate <> "" And candidate <> "." And candidate <> ".." And _
           InStr(candidate, "/") = 0 And InStr(candidate, "\") = 0 Then
            folderCount = folderCount + 1
            folders(folderCount) = candidate
        End If
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

    invoiceNumber = TruncateUtf8V2( _
        CleanFileNameV2(CStr(ws.Cells(2, 5).Value)), _
        MAX_INVOICE_NUMBER_FILENAME_BYTES)
    model = TruncateUtf8V2( _
        CleanFileNameV2(CStr(ws.Cells(13, 5).Value)), _
        MAX_MODEL_FILENAME_BYTES)
    fleetNumber = TruncateUtf8V2( _
        CleanFileNameV2(CStr(ws.Cells(13, 4).Value)), _
        MAX_FLEET_FILENAME_BYTES)
    BuildInvoiceBaseNameV2 = TruncateUtf8V2( _
        dateText & "_" & invoiceNumber & "_" & model & "_" & fleetNumber, _
        MAX_FILENAME_BASE_BYTES)
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
        finalXlsmPath = Trim$(CStr(resultParts(1)))
        finalPdfPath = Trim$(CStr(resultParts(2)))
        If finalXlsmPath <> "" And finalPdfPath <> "" And _
           LCase$(Right$(finalXlsmPath, 5)) = ".xlsm" And _
           LCase$(Right$(finalPdfPath, 4)) = ".pdf" Then
            MoveMacStagedFilesV2 = True
            Exit Function
        End If
        finalXlsmPath = ""
        finalPdfPath = ""
        errorMessage = "The Mac invoice helper reported success without valid output paths."
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

' The VBA editor reads an imported .bas file in the machine's legacy code page
' (Mac Roman or Windows-1252), not UTF-8, so a Polish letter typed directly in
' this source arrives garbled: the Save button's "FAKTUR\u0118" printed with two
' junk characters in place of the final letter. Keep this module pure ASCII:
' write each non-ASCII character as a backslash, "u" and four hex digits, as
' in the button captions, and pass the string through this function.
Private Function UnicodeTextV2(ByVal escapedText As String) As String
    Dim result As String
    Dim startPos As Long
    Dim hitPos As Long

    startPos = 1
    hitPos = InStr(startPos, escapedText, "\u", vbBinaryCompare)
    Do While hitPos > 0
        result = result & Mid$(escapedText, startPos, hitPos - startPos) & _
                 ChrW$(CLng("&H" & Mid$(escapedText, hitPos + 2, 4)))
        startPos = hitPos + 6
        hitPos = InStr(startPos, escapedText, "\u", vbBinaryCompare)
    Loop
    UnicodeTextV2 = result & Mid$(escapedText, startPos)
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
    ' Windows/SMB clients reject trailing dots and spaces even when macOS lets
    ' the user type them. Keep every generated name portable across the share.
    Do While Len(result) > 0 And (Right$(result, 1) = "." Or Right$(result, 1) = " ")
        result = Left$(result, Len(result) - 1)
    Loop
    If result = "" Then result = "NA"
    CleanFileNameV2 = result
End Function

Private Function TruncateUtf8V2(ByVal value As String, _
                                ByVal maximumBytes As Long) As String
    Dim index As Long
    Dim codeUnit As Long
    Dim nextUnit As Long
    Dim characterBytes As Long
    Dim unitsUsed As Long
    Dim bytesUsed As Long
    Dim result As String

    index = 1
    Do While index <= Len(value)
        codeUnit = AscW(Mid$(value, index, 1))
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        unitsUsed = 1

        If codeUnit <= &H7F Then
            characterBytes = 1
        ElseIf codeUnit <= &H7FF Then
            characterBytes = 2
        ElseIf codeUnit >= 55296 And codeUnit <= 56319 And index < Len(value) Then
            nextUnit = AscW(Mid$(value, index + 1, 1))
            If nextUnit < 0 Then nextUnit = nextUnit + 65536
            If nextUnit >= 56320 And nextUnit <= 57343 Then
                characterBytes = 4
                unitsUsed = 2
            Else
                characterBytes = 3
            End If
        Else
            characterBytes = 3
        End If

        If bytesUsed + characterBytes > maximumBytes Then Exit Do
        result = result & Mid$(value, index, unitsUsed)
        bytesUsed = bytesUsed + characterBytes
        index = index + unitsUsed
    Loop

    TruncateUtf8V2 = result
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
