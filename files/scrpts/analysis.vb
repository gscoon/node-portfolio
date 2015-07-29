Option Explicit
Public institution As String
Public adowkbk As Workbook, sqlwkbk As Workbook
Public adowksht As Worksheet, sqlwksht As Worksheet, currentsht As Worksheet, mfiwksht As Worksheet, datawksht As Worksheet
Public usdfx As Double, eurfx As Double

Sub showAnalysisTemplateDrop(control As IRibbonControl)
    currentMFISelectionType = "analysis_template"
    Call showDropMenu
End Sub

Sub refreshMFIFinancials(control As IRibbonControl)
    currentMFISelectionType = "refresh_analysis_template"
    Call showDropMenu
End Sub

'Callback for Load_Analysis_Tempalte onAction
Function Load_Analysis_Template(mfi_id As Integer, mfi_name As String, Optional fund_filter As String, Optional msboxInactive As Boolean) As Boolean

    If (checkDBConnection = False) Then Exit Function

    Dim rs As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dbConnectionObj
    sql = "SELECT mfi.mfi_id AS MFI_ID, mfi.mfi_name AS MFI_NAME, c.country_name, r.region_id, r.region_name, L.legal_status, mfi.network, mfi.reporting_currency AS CURRENCY_CODE, mfi.date_established, fx_historical.fx_date, fx_historical.fx_value AS USD_FX_RATE, fx2.fx_value AS USD_EUR_FX_RATE " & _
    "FROM mfi LEFT JOIN fx_historical ON ( mfi.reporting_currency = fx_historical.conversion_currency_code ) LEFT JOIN (SELECT * FROM fx_historical ORDER BY fx_date DESC) fx2 ON (fx2.conversion_currency_code = 'EUR') LEFT JOIN country c ON c.country_id = mfi.country_id LEFT JOIN region r ON r.region_id = c.region_id LEFT JOIN mfi_legal_status L ON L.legal_status_id = mfi.legal_status_id " & _
    "WHERE mfi.mfi_id = '" & mfi_id & "' AND fx_historical.base_currency_code = 'USD' AND fx2.base_currency_code = 'USD' ORDER BY fx_historical.fx_date DESC LIMIT 0 , 1"
    rs.Open sql

    If (rs.EOF = True) Then
        If (msboxInactive = False) Then MsgBox "MFI data could not be found"
        Exit Function
    End If

    Dim qtData As QueryTable

    Dim current_mfid As String

    Dim newWkbk As Workbook

    Dim riskAssessmentSheet As Worksheet

    Dim commentString As String

    Dim dummyBook As Workbook
    Dim configSheet As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set dummyBook = Workbooks.Add

    Application.Calculation = xlCalculationManual

    Call checkForMISFiles("Analysis_Template")
    If (isWorkbookOpen(templateFiles("Analysis_Template")) = True) Then
        Workbooks(templateFiles("Analysis_Template")).Close False
    End If

    Set adowkbk = Workbooks.Open(fileName:=template_location & templateFiles("Analysis_Template"), ReadOnly:=True)

    Set currentsht = adowkbk.Worksheets("Balance Sheet")
    Set adowksht = adowkbk.Worksheets("DB_LOAD")
    Set datawksht = adowkbk.Worksheets("DATA_FORMATTING")
    Set configSheet = adowkbk.Worksheets("Config")

    adowkbk.Activate
    adowksht.Activate

    adowksht.visible = xlSheetVisible
    datawksht.visible = xlSheetVisible

    current_mfid = rs.Fields("MFI_ID").Value
    currentsht.Range("C2") = rs.Fields("MFI_NAME")

    'Populate the template properties in the config sheet
    configSheet.Range("B4") = CStr(getDateModified(adowkbk.FullName)) & " by " & adowkbk.BuiltinDocumentProperties("Last author").Value
    configSheet.Range("B10") = rs.Fields("MFI_NAME").Value
    configSheet.Range("B11") = current_mfid
    configSheet.Range("B12") = rs.Fields("CURRENCY_CODE").Value
    configSheet.Range("B19") = rs.Fields("USD_FX_RATE").Value
    configSheet.Range("B13") = rs.Fields("country_name").Value
    configSheet.Range("B15") = rs.Fields("region_id").Value
    configSheet.Range("B14") = rs.Fields("region_name").Value
    configSheet.Range("B16") = rs.Fields("network").Value
    configSheet.Range("B17") = rs.Fields("date_established").Value
    configSheet.Range("B18") = rs.Fields("legal_status").Value
    configSheet.Range("B20") = rs.Fields("USD_EUR_FX_RATE").Value

    'country, network, legal status, year founded

    Dim sql_array() As Variant
    Dim sql_statement As Variant
    Dim statementIndex As Integer

    sql_array = returnSQLFunctionList(current_mfid, rs.Fields("CURRENCY_CODE").Value, fund_filter)

    Dim currentQTRange As String
    currentQTRange = "C1"
    Call clearQueryTables(adowksht)
    adowkbk.Activate

    statementIndex = 0
    For Each sql_statement In sql_array
        ' Open the recordset.
        Set rs = New ADODB.Recordset
        Set rs.ActiveConnection = dbConnectionObj
        rs.Open sql_statement
        If (statementIndex = 5) Then
            Call printQueryResults(rs, Range(currentQTRange).Column, 1)
            commentString = Range(currentQTRange).Offset(1, 1).Value
        Else
            Set qtData = adowksht.QueryTables.Add(rs, Range(currentQTRange))
            qtData.Refresh
        End If
        'Place next table next to current table with a column spacer...
        currentQTRange = adowksht.Range(currentQTRange).Offset(0, 1 + rs.Fields.Count).Address
        statementIndex = statementIndex + 1
        Set rs = Nothing
    Next sql_statement

    adowksht.Calculate
    datawksht.Calculate
    Sheets("Balance Sheet").Calculate

    adowksht.Cells.Copy
    adowksht.Cells.PasteSpecial xlPasteValues

    datawksht.Cells.Copy
    datawksht.Cells.PasteSpecial xlPasteValues

    adowksht.visible = xlSheetHidden
    datawksht.visible = xlSheetHidden

    Dim templateVersion As String
    templateVersion = configSheet.Cells.Find("#template_version").Offset(0, -1).Value

    configSheet.visible = xlSheetHidden

    'handle comment buttons
    Set riskAssessmentSheet = ActiveWorkbook.Sheets("Risk Assessment")

    adowkbk.Worksheets("Balance Sheet").Range("e5") = "USD"
    adowkbk.Worksheets("Balance Sheet").Range("e9") = adowkbk.Worksheets("Data_Formatting").Range("b2").Value
    adowkbk.Worksheets("Balance Sheet").Range("d9") = adowkbk.Worksheets("Data_Formatting").Range("c2").Value
    adowkbk.Worksheets("Income Statement").Range("e6") = adowkbk.Worksheets("Data_Formatting").Range("b2").Value
    adowkbk.Worksheets("Income Statement").Range("d6") = adowkbk.Worksheets("Data_Formatting").Range("c2").Value

    ActiveSheet.Range("d1").Select

    If (SheetExists("Refi-Annual Review")) Then
        Sheets("Refi-Annual Review").Calculate
        Sheets("MFI Home").Calculate

        If (templateVersion <> "8") Then
            Call formatRefiSheet
            Call formatHomeSheet
        End If
    End If

    Application.Calculation = xlCalculationAutomatic

    On Error Resume Next

    If (SheetExists("Risk Assessment")) Then
        Dim commentary_input As OLEObject
        Set commentary_input = riskAssessmentSheet.OLEObjects("commentary_input")
        commentary_input.Object.Text = commentString
        riskAssessmentSheet.Shapes("CommentarySaveButton").OnAction = "saveRiskAssessmentComment"
        riskAssessmentSheet.Shapes("CommentaryHideButton").OnAction = "hideAllCommentButtons"
        riskAssessmentSheet.Shapes("CommentaryHideButton2").OnAction = "hideAllCommentButtons"
        riskAssessmentSheet.Shapes("riskRatingSaveButton").OnAction = "saveRiskRating"
        riskAssessmentSheet.Range("BH3").Value = current_mfid
        'If (Not IsError(riskAssessmentSheet.Range("C3"))) Then riskAssessmentSheet.OLEObjects("riskRatingDrop").Object.Text = "Stable"
    End If

    On Error GoTo 0
    dummyBook.Close False

    'upload aggregated info
    Call uploadAggregatedData(CInt(current_mfid))
    Call UseBreakLink
    adowkbk.SaveAs fileName:=SHARED_FINANCIALS_PATH & current_mfid & ".xlsx"
    adowkbk.SaveAs fileName:=SHARED_FINANCIALS_PATH_2 & mfi_name & "_MFI Analysis Report.xlsx"

    'Call updateMFITimeStamps(current_mfid, unixTimestamp, SHARED_FINANCIALS_PATH & "mfi.txt")
    adowkbk.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Load_Analysis_Template = True
End Function

Sub removeUneededAnalysisSheets(keepShtName As String, wkbk As Workbook)
    Dim w As Worksheet
    For Each w In wkbk.Worksheets
        If ((w.Name <> "DB_LOAD") And (w.Name <> "DATA_FORMATTING") And (w.Name <> "Balance Sheet") And (w.Name <> "Income Statement") And (w.Name <> "Portfolio Quality") And (w.Name <> "Organizational Data") And (w.Name <> keepShtName) And (w.Name <> "Ratios & Trend Analysis") And (w.Name <> "MFI Home")) Then
            'msgBox w.Name
            w.Delete
        End If
    Next w
End Sub

Sub formatHomeSheet()
    Dim countryleft As Range, countryright As Range
    Dim refiSheet As Worksheet
    Set refiSheet = ActiveWorkbook.Sheets("Refi-Annual Review")
    On Error GoTo endFormatting
    refiSheet.Activate
    Set countryleft = refiSheet.Cells.Find("#countryleft")
    Set countryright = refiSheet.Cells.Find("#countryright")
    refiSheet.Range(countryleft.Offset(0, 1), countryright.Offset(0, -1)).Copy
    ActiveWorkbook.Sheets("MFI Home").Range("b51").PasteSpecial xlPasteValues
    ActiveWorkbook.Sheets("MFI Home").Range("b51").PasteSpecial xlPasteFormats
    Sheets("MFI Home").Activate
    ActiveSheet.Range("b45:J100").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$a45=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With

    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399945066682943
    End With
endFormatting:
    Selection.FormatConditions(1).StopIfTrue = False
    Sheets("Ratios & Trend Analysis").Move Before:=Sheets("Balance Sheet")
    Sheets("MFI Home").Move Before:=Sheets("Ratios & Trend Analysis")
    Sheets("MFI Home").Calculate

    Dim x As Long
    Dim legendcheck As Range

    Set legendcheck = Sheets("MFI Home").Range("O53")

    For x = 7 To 2 Step -1
        Sheets("MFI Home").ChartObjects("Home Chart").Activate
        ActiveChart.Legend.Select
        ActiveChart.Legend.LegendEntries(x).Select

        If legendcheck.Offset(x - 7, 0).Value = 1 Then Selection.Delete
    Next x

    ActiveWorkbook.Worksheets("MFI Home").Range("A1").Activate
End Sub

Sub formatRefiSheet()

    Dim deleteLabel As Range
    Dim deleteLabelEnd As Range

    Dim refiSheet As Worksheet
    Dim deleteRange As Range
    Dim r As Range
    Dim rFoundCell As Range
    Dim lCount As Integer

    Dim wasHidden As Boolean

    Set refiSheet = Sheets("Refi-Annual Review")
    refiSheet.Activate
    If (refiSheet.visible = xlSheetHidden) Then
        wasHidden = True
        refiSheet.visible = xlSheetVisible
    End If
    Set deleteLabel = refiSheet.Cells.Find("#deleteLabel")
    Set deleteLabelEnd = refiSheet.Cells.Find("#deleteLabelEnd")

    If (Not deleteLabel Is Nothing) Then

        refiSheet.Range(Cells(1, deleteLabel.Column), Cells(300, deleteLabelEnd.Column)).Copy
        refiSheet.Cells(1, deleteLabel.Column).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Set deleteRange = refiSheet.Range(deleteLabel, Cells(GetLastRowWithData(refiSheet.Columns(deleteLabel.Column)), deleteLabel.Column))
        For lCount = 1 To WorksheetFunction.CountIf(deleteRange, "1")
            Set rFoundCell = deleteRange.Find(What:="1", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            If (Not rFoundCell Is Nothing) Then
                Range(rFoundCell, rFoundCell.Offset(0, 10)).Select
                Selection.Delete xlShiftUp
            End If
        Next lCount
        deleteRange.Select
    End If
    'refiSheet.visible = xlSheetVisible
    If (wasHidden = True) Then
        refiSheet.visible = xlSheetHidden
    End If

End Sub


Function returnSQLFunctionList(mfi_id As String, currency_id As String, Optional fund_filter As String) As Variant
    Dim sql(0 To 9) As Variant
    Dim fundPart As String

    'financial information
    sql(0) = "SELECT DS.sheet_name, FL.field_label_english, FS.is_audited, FS.field_set_date, F.field_value FROM reported_financials F LEFT JOIN reporting_field_set FS ON FS.field_set_id = F.field_set_id LEFT JOIN mfi M ON (M.mfi_id = FS.mfi_id) LEFT JOIN reporting_field_label FL ON F.field_label_id = FL.field_label_id LEFT JOIN reporting_data_sheet DS ON DS.sheet_id = FS.sheet_id WHERE M.mfi_id = " & mfi_id

    'historical fx rates
    sql(1) = "SELECT concat(MONTH(FX.fx_date),'-',YEAR(FX.fx_date)), FX.conversion_currency_code, FX.fx_value, FX.fx_date FROM (SELECT FX.conversion_currency_code, MAX(FX.fx_date) as mfd FROM fx_historical FX WHERE FX.base_currency_code = 'USD' GROUP BY fx.conversion_currency_code, YEAR(FX.fx_date), MONTH(FX.fx_date) ) MF LEFT JOIN fx_historical FX ON (FX.conversion_currency_code = MF.conversion_currency_code AND FX.fx_date = MF.mfd AND FX.base_currency_code = 'USD') LEFT JOIN mfi M ON (M.reporting_currency = FX.conversion_currency_code) WHERE M.mfi_id = " & mfi_id & " ORDER BY FX.fx_date DESC"

    'filter by fund
    fundPart = ""
    If fund_filter <> "" Then fundPart = " AND f.fund_id IN (" & fund_filter & " )"

    'sum of transactions for each fund

    sql(2) = "SELECT f.fund_name, st.sumt as 'Sum of Deals (OC)', st.sumt_fx as 'Sum of Deals ($ at Transaction Date)',  (st.sumt/fx.fx_value) as 'Sum of Deals ($ at Latest FX Date)', st.currency_code as 'Deal Currency', st.deal_date, st.d_mat as 'Deal Maturity', if(st.index_name IS NULL, st.crate, concat(round(st.crate, 2), '% ', st.plus_or_minus_rate, ' ',st.index_name)) as 'Coupon Rate', st.deal_type_id as 'Deal Type' FROM " & _
    "(SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate, IRI.index_name, dd.plus_or_minus_rate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') " & _
    "LEFT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id LEFT JOIN interest_rate_index IRI ON IRI.index_id = dd.interest_rate_index_id GROUP BY d.mfi_id, d.deal_date, t.fund_id, t.currency_code HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi m ON m.mfi_id = st.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id " & _
    "LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code)  max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date " & _
    "WHERE m.mfi_id = " & mfi_id & " AND st.sumt > 0 ORDER BY st.d_mat ASC, f.fund_name"

    'amortization history
    sql(3) = "SELECT CONCAT(st.d_mat,'-',f.fund_name) as 'lookup_key', f.fund_name as 'Fund', st.sumt as 'Sum of Deals (OC)', (st.sumt/fx.fx_value) as 'Sum of Deals ($ at Latest FX Date)', st.currency_code as 'Deal Currency', st.d_mat as 'Tranche Maturity' FROM " & _
    "(SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate, IRI.index_name, dd.plus_or_minus_rate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') " & _
    "RIGHT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id LEFT JOIN interest_rate_index IRI ON IRI.index_id = dd.interest_rate_index_id GROUP BY d.mfi_id, t.fund_id, t.currency_code, dd.maturity_date HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi m ON m.mfi_id = st.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id " & _
    "LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code)  max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date " & _
    "WHERE m.mfi_id = " & mfi_id & " AND st.sumt > 0 ORDER BY st.d_mat ASC, f.fund_name, st.currency_code"

    'covenant information
    sql(4) = "SELECT ct.covenant_type, c.covenant_value, c.start_date, c.end_date FROM mfi_covenant c LEFT JOIN mfi_covenant_type ct ON ct.covenant_type_id = c.covenant_type_id WHERE c.mfi_id=" & mfi_id

    'mfi commentary
    sql(5) = "SELECT c.comment_date,  c.comments FROM mfi_commentary c WHERE c.mfi_id = " & mfi_id & " ORDER BY c.comment_date DESC"

    'mfi rating
    sql(6) = "SELECT r.rating_date, rl.rating_label_short, rl.rating_label_long FROM mfi_risk_rating r LEFT JOIN mfi_risk_rating_label rl ON r.rating_label_id = rl.rating_label_id WHERE r.mfi_id = " & mfi_id & " ORDER BY r.rating_date DESC"

    'country mfis
    sql(7) = "SELECT F.fund_name, M.mfi_name, M.mfi_id, C.country_name, st.sumt, st.sumt_fx, st.currency_code, st.deal_type_id, (st.sumt/fx.fx_value) as 'Amount as of Last FX'  FROM (SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') LEFT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id GROUP BY d.mfi_id, d.deal_date, t.fund_id, t.currency_code HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi M ON st.mfi_id = M.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id LEFT JOIN country C ON C.country_id = M.country_id " & _
    "LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code) max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date WHERE M.country_id IN (SELECT country_id FROM mfi WHERE mfi_id = " & mfi_id & ") AND st.sumt > 0"

    'social data
    sql(8) = "SELECT m.mfi_name, m.mfi_id, sf.field_name, sv.field_id, sv.field_value, ss.questionnaire_year FROM reporting_social_value sv JOIN reporting_social_field sf ON sf.field_id = sv.field_id JOIN reporting_social_submission ss ON ss.submission_id = sv.submission_id JOIN mfi m ON m.mfi_id = ss.mfi_id LEFT JOIN (SELECT mfi_id, MAX(questionnaire_year) as max_yr FROM reporting_social_submission GROUP BY mfi_id) max_ss ON max_ss.mfi_id = m.mfi_id WHERE sf.field_id = 353 AND ss.questionnaire_year = max_ss.max_yr AND m.mfi_id = " & mfi_id & ""

    'other ratings
    sql(9) = "SELECT rating.rating_type, rating.rating_date, rating.rating_value FROM mfi_other_rating rating JOIN mfi m ON m.mfi_id = rating.mfi_id LEFT JOIN (SELECT mfi_id, rating_type, MAX(rating_date) as max_date FROM mfi_other_rating GROUP BY mfi_id, rating_type) max_rating ON max_rating.mfi_id = rating.mfi_id AND max_rating.rating_type = rating.rating_type WHERE rating.rating_date = max_rating.max_date AND rating.mfi_id = " & mfi_id & ""

    returnSQLFunctionList = sql
End Function

Sub Switch_Institutions(control As IRibbonControl)
    currentMFISelectionType = "analysis_template"
    Call showDropMenu
End Sub

Sub hideAllCommentButtons()
    Dim sht As Worksheet
    Dim saveButton As Shape
    Dim hideButton As Shape

    For Each sht In ActiveWorkbook.Worksheets
        If (shapeExists("CommentarySaveButton", sht) = True) Then sht.Shapes("CommentarySaveButton").visible = msoFalse
        If (shapeExists("CommentaryHideButton", sht) = True) Then sht.Shapes("CommentaryHideButton").visible = msoFalse
        If (shapeExists("CommentaryHideButton2", sht) = True) Then sht.Shapes("CommentaryHideButton2").visible = msoFalse
        If (shapeExists("riskRatingSaveButton", sht) = True) Then sht.Shapes("riskRatingSaveButton").visible = msoFalse
        If (shapeExists("riskRatingDrop", sht) = True) Then sht.Shapes("riskRatingDrop").visible = msoFalse
    Next sht

    On Error GoTo 0
End Sub

Sub showAllCommentButtons()
    Dim sht As Worksheet
    Dim saveButton As Shape
    Dim hideButton As Shape

    For Each sht In ActiveWorkbook.Worksheets
        If (shapeExists("CommentarySaveButton", sht) = True) Then sht.Shapes("CommentarySaveButton").visible = msoCTrue
        If (shapeExists("CommentaryHideButton", sht) = True) Then sht.Shapes("CommentaryHideButton").visible = msoCTrue
        If (shapeExists("CommentaryHideButton2", sht) = True) Then sht.Shapes("CommentaryHideButton2").visible = msoCTrue
        If (shapeExists("riskRatingSaveButton", sht) = True) Then sht.Shapes("riskRatingSaveButton").visible = msoCTrue
        If (shapeExists("riskRatingDrop", sht) = True) Then sht.Shapes("riskRatingDrop").visible = msoCTrue
    Next sht

    On Error GoTo 0
End Sub


Sub saveRiskRating()
    Dim insertStatement As String
    Dim deleteStatement As String
    Dim mfi_id As Range
    Dim commentary_input As OLEObject
    Dim rating As String
    Dim dateString As String
    Dim ratingLabelID As Integer

    Dim riskRatingLabel As Range

    'rating_label_long
    If (checkDBConnection = False) Then Exit Sub


    Call handleAllGroupedRanges(ActiveSheet, False)

    rating = ActiveSheet.OLEObjects("riskRatingDrop").Object.Text
    Set riskRatingLabel = ActiveSheet.Cells.Find(What:="#rating_label_long", LookIn:=xlValues, LookAt:=xlPart)

    If (Not riskRatingLabel Is Nothing) Then
        ratingLabelID = CInt(Range(riskRatingLabel, riskRatingLabel.End(xlDown)).Find(rating).Offset(0, 2))

        Set mfi_id = Range("BH3")

        dateString = WorksheetFunction.Text(Date, "YYYY-MM-DD")

        On Error GoTo dberror:
        deleteStatement = "Delete FROM mfi_risk_rating WHERE mfi_id = " & mfi_id & " AND rating_date = '" & dateString & "'"
        dbConnectionObj.Execute deleteStatement

        insertStatement = "INSERT INTO mfi_risk_rating (mfi_id, rating_label_id, rating_date) VALUES (" & mfi_id & ", " & ratingLabelID & " ,'" & dateString & "')"
        dbConnectionObj.Execute insertStatement

        ActiveSheet.Range("C3") = rating

        MsgBox "Your rating has been saved."
    End If
    Call handleAllGroupedRanges(ActiveSheet, True)
    Exit Sub
dberror:
End Sub

Sub AnalysisTemplateScraper(sqlStatement As String, shtName As String, Optional activeTemplate As Workbook)

    Dim i As Integer, J As Integer
    Dim rs As ADODB.Recordset
    Dim f As Field
    Dim master As Workbook, analysis_template As Workbook
    Dim dwmImgObj As Object

    If (checkDBConnection = False) Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dbConnectionObj
    rs.Open sqlStatement
    rs.MoveFirst

    If (IsMissing(activeTemplate) = True Or activeTemplate Is Nothing) Then
        Set master = Workbooks.Add(gpath & "\templates\template.xlsx")
    Else
        Set master = activeTemplate
    End If

    While Not rs.EOF
        Call pullMFITemplateFromFile(mfi_id:=rs.Fields("mfi_id"), mfi_name:=rs.Fields("mfi_name"), silenced:=True)

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        Set analysis_template = ActiveWorkbook

        If (SheetExists(shtName) = True) Then
            analysis_template.Sheets(shtName).Copy Before:=master.Sheets(1)

            master.Sheets(1).Cells.Copy
            master.Sheets(1).Cells.PasteSpecial (xlPasteValues)

            Dim N As Name
            For Each N In master.Sheets(1).Names
                N.Delete
            Next N

            master.Sheets(1).visible = xlSheetVisible

            master.Sheets(1).Name = rs.Fields("mfi_name")

            analysis_template.Close False
        End If
        rs.MoveNext
    Wend

    rs.Close

    Call deleteDefaultSheets(master)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Unload mfi_filter_form
End Sub

Sub pullMFITemplateFromFile(mfi_id As Integer, mfi_name As String, Optional silenced As Boolean)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Compare timestamps in mfi.txt files (local and shared)
    Dim localMFIDate As Date, sharedMFIDate As Date
    Dim existingWorkbook As Workbook

    localMFIDate = getMFIDateModified(financials_location & mfi_id & ".xlsx")
    sharedMFIDate = getMFIDateModified(SHARED_FINANCIALS_PATH & mfi_id & ".xlsx")
    Application.StatusBar = "Pulling financials for " & mfi_name

    'Do financials exist for this mfi
    If (sharedMFIDate > 0) Then
        'Check if clean financials exist locally
        If (FileFolderExists(financials_location & mfi_id & ".xlsx") = True) Then

            'Is there a newer version in the shared drive
            If (sharedMFIDate > localMFIDate) Then
                Set existingWorkbook = pullFinancialsFromSharedDrive(mfi_id)
                Call saveTemplateToTemp(mfi_name, existingWorkbook)
            'Just use the one in the clean financials folder
            Else
                Set existingWorkbook = Workbooks.Open(fileName:=financials_location & mfi_id & ".xlsx", ReadOnly:=True)
                Call saveTemplateToTemp(mfi_name, existingWorkbook)
            End If

            'set commentary
            Call setCommentaryBox
        Else
            'File Does not exists locally
            Set existingWorkbook = pullFinancialsFromSharedDrive(mfi_id)
            Call saveTemplateToTemp(mfi_name, existingWorkbook)
        End If
    Else
        If (silenced = False) Then MsgBox "There are no financials for this MFI."
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Function pullFinancialsFromSharedDrive(mfi_id As Integer) As Workbook
    Dim existingWorkbook As Workbook
    Dim mfiSharedPath As String

    mfiSharedPath = SHARED_FINANCIALS_PATH & mfi_id & ".xlsx"
    FileCopy mfiSharedPath, financials_location & mfi_id & ".xlsx"
    Set existingWorkbook = Workbooks.Open(fileName:=financials_location & mfi_id & ".xlsx", ReadOnly:=True)
    'existingWorkbook.SaveAs financials_location & mfi_name & ".xlsx"
    'Call updateMFITimeStamps(mfi_id, ts, financials_location & "mfi.txt")
    Set pullFinancialsFromSharedDrive = existingWorkbook
End Function

Sub saveTemplateToTemp(book_name As String, wkbk As Workbook)
    If (Not wkbk Is Nothing) Then
        Dim oFS As New FileSystemObject
        Dim tempFolder As String
        Dim d As Date
        d = Now
        tempFolder = oFS.GetSpecialFolder(TemporaryFolder)
        'save with mfi name and the number of seconds since the beginning of this month
        wkbk.SaveAs tempFolder & "\" & book_name & "_" & Round(unixTimestamp - convertToTimestamp(CDate(Month(d) & "/1/" & Year(d)))) & ".xlsx"
    End If
End Sub

Sub updateMFITimeStamps(mfi_id, ts, filePath)
    Dim txtObj As New Scripting.FileSystemObject
    Dim txtFile As TextStream
    Dim colFiles As New Collection
    Dim x As Variant
    Dim i As Integer

    Set txtFile = txtObj.OpenTextFile(filePath, ForReading)
    Do Until txtFile.AtEndOfStream
        x = Split(txtFile.ReadLine, "|")
        'Ignore the current mfi. Itll be added to the end
        If (x(0) <> mfi_id) Then colFiles.Add x(0) & "|" & x(1)
    Loop

    'add mfi|timestamp list to mfi.txt file
    Set txtFile = txtObj.OpenTextFile(filePath, ForWriting)
    If colFiles.Count > 0 Then
        For i = 1 To colFiles.Count
            txtFile.WriteLine colFiles(i)
        Next i
    End If

    'Its a new mfi, so add it, man
    txtFile.WriteLine mfi_id & "|" & ts
End Sub

Function getMFIDateModified(filePath As String) As Date
    Dim oFS As Object
    If (FileFolderExists(filePath) = True) Then
        Set oFS = CreateObject("Scripting.FileSystemObject")
        getMFIDateModified = oFS.GetFile(filePath).DateLastModified
        Exit Function
    End If
    getMFIDateModified = 0
End Function


Sub uploadAggregatedData(mfi_id As Integer)
    Dim ag_sheet As Worksheet
    Dim start As Range
    Dim r As Range, c As Range
    Dim theCol As Range
    Dim valueString As String
    Dim dataset_id As Long
    Dim data_sheet_id As Integer
    Dim field_set_date As String
    Dim current_value As String
    Dim rEnd As Range

    Set ag_sheet = Sheets("Aggregated_Data")
    ag_sheet.visible = xlSheetVisible
    ag_sheet.Activate
    data_sheet_id = ag_sheet.Range("a1").Value

    Set start = ag_sheet.Range("c6")
    For Each r In Range(start, start.End(xlToRight))
        valueString = ""
        If (IsDate(r.Offset(-1, 0)) = True) Then
            field_set_date = WorksheetFunction.Text(r.Offset(-1, 0), "YYYY-MM-DD")

            Set rEnd = Cells(GetLastRowWithData(Columns(r.Column)), r.Column) 'r.End(xlDown)
            For Each c In Range(r, rEnd)
                If (TypeName(c.Value) = "Integer" Or TypeName(c.Value) = "Double") Then
                    current_value = c.Value
                Else
                    current_value = 0
                End If
                If (Trim(current_value) = "") Then current_value = "0"
                If (valueString <> "") Then valueString = valueString & ","
                valueString = valueString & "(" & Cells(c.row, 1).Value & "," & current_value & ", [datasetID])"
            Next c
            dataset_id = createDataset(mfi_id, data_sheet_id, field_set_date, 0)
            valueString = Replace(valueString, "[datasetID]", CStr(dataset_id))
            dbConnectionObj.Execute "INSERT INTO reported_financials (field_label_id, field_value, field_set_id) VALUES " & valueString
        End If
    Next r
    ag_sheet.visible = xlSheetHidden
End Sub

Sub setCommentaryBox()
    If (SheetExists("Risk Assessment")) Then
        Dim commentary_input As OLEObject
        Dim riskAssessmentSheet As Worksheet
        Set riskAssessmentSheet = Sheets("Risk Assessment")
        riskAssessmentSheet.Shapes("CommentarySaveButton").OnAction = "saveRiskAssessmentComment"
        riskAssessmentSheet.Shapes("CommentaryHideButton").OnAction = "hideAllCommentButtons"
        riskAssessmentSheet.Shapes("CommentaryHideButton2").OnAction = "hideAllCommentButtons"
        riskAssessmentSheet.Shapes("riskRatingSaveButton").OnAction = "saveRiskRating"
        'If (Not IsError(riskAssessmentSheet.Range("C3"))) Then riskAssessmentSheet.OLEObjects("riskRatingDrop").Object.Text = "Stable"
    End If
End Sub


Sub refreshAllMFIs(control As IRibbonControl)
    If (checkDBConnection = False) Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim sqlStatement As String
    Dim rs As ADODB.Recordset

    sqlStatement = "SELECT m.mfi_id, m.mfi_name FROM reporting_field_set FS JOIN mfi m ON m.mfi_id = FS.mfi_id GROUP BY m.mfi_id ORDER BY m.mfi_name DESC"

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dbConnectionObj
    rs.Open sqlStatement
    rs.MoveFirst

    While Not rs.EOF
        Call Load_Analysis_Template(mfi_id:=rs.Fields("mfi_id"), mfi_name:=rs.Fields("mfi_name"))
        rs.MoveNext
    Wend

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
