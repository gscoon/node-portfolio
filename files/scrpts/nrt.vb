Sub NRT_Upload(control As IRibbonControl)
    Call uploadThisNRT
End Sub

Sub uploadThisNRT(Optional mfiID As Integer, Optional mfiName As String)

    If (checkDBConnection = False) Then Exit Sub

    Dim ImportWkbk As Workbook, mfiWkbk As Workbook
    Dim appAccess As Object
    Dim ImportDB As String
    Dim UA As Worksheet, Format As Worksheet, Pre As Worksheet, Export As Worksheet
    Dim ExportCopy As Range, SumRow As Range

    Dim SumRangeCount As Long
    Dim RowCounter As Long
    Dim DeleteRow As Range
    Dim zero As Long
    Dim NextRow As Long
    Dim currentReportingMonth As String
    Dim uploadyear As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If (Application.Workbooks.Count = 0) Then
        MsgBox "There are no workbooks open."
        Exit Sub
    End If

    Set mfiWkbk = ActiveWorkbook
    Call trimSheetNames(mfiWkbk) 'some worksheets are returned with extra spaces in their names. get rid of those

    If (mfiName = "") Then 'this is for new reporting template
        If (SheetExists("MFI Info & Instructions") = False) Then
            MsgBox "Please use a valid reporting template"
            Exit Sub
        End If

        Application.StatusBar = "Starting new reporting template upload"

        Dim CopyRange_BS As Range, CopyRange_IS As Range, CopyRange_PQ As Range, CopyRange_LS As Range, PasteRange_BS As Range, PasteRange_IS As Range, PasteRange_PQ As Range, PasteRange_LS As Range

        Call checkForMISFiles("NRT_Upload")
        If IsFileOpen(template_location & templateFiles("NRT_Upload")) = True Then Workbooks(templateFiles("NRT_Upload")).Close False
        Set ImportWkbk = Workbooks.Open(template_location & templateFiles("NRT_Upload"), ReadOnly:=True)

        mfiName = mfiWkbk.Sheets(1).Range("d8").Value


        Set CopyRange_BS = mfiWkbk.Worksheets("Balance Sheet").Range("J6:AT107")
        Set CopyRange_IS = mfiWkbk.Worksheets("Income Statement").Range("J6:AT66")
        Set CopyRange_PQ = mfiWkbk.Worksheets("Portfolio & Organizational Data").Range("J6:AT49")
        Set CopyRange_LS = mfiWkbk.Worksheets("Funding & Shareholders").Range("D9:BA36")
        Set CopyRange_EO = mfiWkbk.Worksheets("Funding & Shareholders").Range("D50:H65")

        Set PasteRange_BS = ImportWkbk.Worksheets("Balance Sheet").Range("J6:AT107")
        Set PasteRange_IS = ImportWkbk.Worksheets("Income Statement").Range("J6:AT66")
        Set PasteRange_PQ = ImportWkbk.Worksheets("Portfolio & Organizational Data").Range("J6:AT49")
        Set PasteRange_LS = ImportWkbk.Worksheets("Funding & Shareholders").Range("E9")
        Set PasteRange_EO = ImportWkbk.Worksheets("Funding & Shareholders").Range("E54")

        mfiWkbk.Activate
        CopyRange_BS.Copy
        ImportWkbk.Activate
        PasteRange_BS.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        mfiWkbk.Activate
        CopyRange_IS.Copy
        ImportWkbk.Activate
        PasteRange_IS.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        mfiWkbk.Activate
        CopyRange_PQ.Copy
        ImportWkbk.Activate
        PasteRange_PQ.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        mfiWkbk.Activate
        CopyRange_LS.Copy
        ImportWkbk.Activate
        PasteRange_LS.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        mfiWkbk.Activate
        CopyRange_EO.Copy
        ImportWkbk.Activate
        PasteRange_EO.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        currentReportingMonth = CStr(WorksheetFunction.Text(ImportWkbk.Sheets(1).Range("a2").Value, "YYYY-MM-DD"))
        uploadyear = CStr(Year(mfiWkbk.Worksheets("Balance Sheet").Range("z6")))

    Else ' this is for old reporting template

        Application.StatusBar = "Starting old reporting template upload"

        Dim CopyRange As Range, CopyRange2 As Range, CopyRange3 As Range, CopyRange4 As Range, PasteRange As Range, PasteRange2 As Range, PasteRange3 As Range, PasteRange4 As Range

        Call checkForMISFiles("Upload_Template_Old")
        If IsFileOpen(template_location & templateFiles("Upload_Template_Old")) = True Then Workbooks(templateFiles("Upload_Template_Old")).Close False
        Set ImportWkbk = Workbooks.Open(template_location & templateFiles("Upload_Template_Old"), ReadOnly:=True)

        Set UA = ImportWkbk.Worksheets("UA")
        Set PasteRange = UA.Range("E6")
        Set PasteRange2 = UA.Range("E43")
        Set PasteRange3 = UA.Range("E103")
        Set PasteRange4 = UA.Range("E5")


        Set CopyRange = mfiWkbk.Worksheets(3).Range("B9:M36")
        Set CopyRange2 = mfiWkbk.Worksheets(3).Range("B44:M100")
        Set CopyRange3 = mfiWkbk.Worksheets(4).Range("B7:M34")
        Set CopyRange4 = mfiWkbk.Worksheets(3).Range("B8:M8")


        mfiWkbk.Activate
        CopyRange.Copy
        ImportWkbk.Activate
        PasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mfiWkbk.Activate
        CopyRange2.Copy
        ImportWkbk.Activate
        PasteRange2.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mfiWkbk.Activate
        CopyRange3.Copy
        ImportWkbk.Activate
        PasteRange3.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        CopyRange4.Copy
        ImportWkbk.Activate
        PasteRange4.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        uploadyear = CStr(Year(mfiWkbk.Worksheets(3).Range("b8")))
    End If

    Dim DateRow As Range, dataSheetIDColumnCell As Range, dataColumn As Integer, dataLabelRow As Integer
    Dim currentDataColumn As Range
    Dim auditLabelRow As Range
    Dim auditLabelString As Integer
    Dim currentSheet As Worksheet
    Dim wksht As Worksheet
    Dim worksheetCount As Integer
    Dim rs As ADODB.Recordset

    Dim firstSheetID As Range
    Dim currentFieldLabelID As Range
    Dim dataSheetIDColumn As Range
    Dim current_field_set_id As String
    Dim currentSheetID As String

    Dim fieldFindCount As Integer
    Dim dataFindCount As Integer
    Dim rFoundCell As Range
    Dim cFoundCell As Range
    Dim sqlString As String
    Dim valueString As String
    Dim fieldString As String
    Dim currentValue As String
    Dim lastCellRow As Long
    Dim dataSheetArray
    Dim currentTableName
    Dim J As Integer
    Dim t As Integer
    Dim dataRowsAdded As Integer

    Dim currencyID As String

    Dim currentDateString As String
    Dim deleteDuplicateString As String
    Dim newFieldSetString As String

    Dim Com As New ADODB.Command

    worksheetCount = 0
    dataRowsAdded = 0

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dbConnectionObj
    rs.Open "SELECT mfi_id from mfi WHERE mfi_name = '" & mfiName & "'"
    If (rs.EOF = True) Then
        MsgBox "The MFI (" & mfiName & ") in the 'MFI Info & Instructions' tab could not be found."
        ImportWkbk.Close False
        Exit Sub
    End If
    rs.MoveFirst
    mfiID = rs.Fields("mfi_id").Value

    For Each wksht In ActiveWorkbook.Worksheets

        Set currentSheet = wksht
        currentSheet.Activate

        Call handleAllGroupedRanges(currentSheet, False)

        'Pulls information from funding tab
        dataRowsAdded = handleOtherFinancialInformation(currentSheet, CInt(mfiID), currentReportingMonth) + dataRowsAdded

        Set DateRow = currentSheet.Cells.Find(What:="#DateRow", LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)

        Set dataSheetIDColumnCell = currentSheet.Cells.Find(What:="#DataSheetIDColumn", LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)

        Set currentDataColumn = currentSheet.Cells.Find(What:="#FieldSet", LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)

        Set auditLabelRow = currentSheet.Cells.Find(What:="#AuditLabelRow", LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False)

        If (Not DateRow Is Nothing) And (Not dataSheetIDColumnCell Is Nothing) And (Not currentDataColumn Is Nothing) Then
            worksheetCount = worksheetCount + 1
            dataLabelRow = currentDataColumn.row

            Set firstSheetID = dataSheetIDColumnCell.End(xlDown)
            lastCellRow = GetLastRowWithData(currentSheet.Columns(dataSheetIDColumnCell.Column))

            Set dataSheetIDColumn = currentSheet.Range(firstSheetID, Cells(lastCellRow, dataSheetIDColumnCell.Column))

            dataSheetIDColumn.Select

            dataSheetArray = UniqueItems(dataSheetIDColumn, False)
            For t = LBound(dataSheetArray) To UBound(dataSheetArray)
                currentSheetID = dataSheetArray(t)
                If (currentSheetID <> "") Then
                'for each field set
                    Set cFoundCell = Cells(currentDataColumn.row, 1)
                    For dataFindCount = 1 To WorksheetFunction.CountIf(Rows(dataLabelRow), "#FieldSet")

                        Set cFoundCell = Rows(dataLabelRow).Find(What:="#FieldSet", After:=cFoundCell, _
                            LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, MatchCase:=False)

                        '*** CREATE NEW FIELD SET AT THIS POINT. MAKE SURE YOU REMOVE ANY EXISTING FIELD SETS WITH THE SAME MFI-SHEET-DATE-AUDIT COMBO
                        auditLabelString = 0
                        currentDateString = Cells(DateRow.row, cFoundCell.Column)

                        If (Cells(auditLabelRow.row, cFoundCell.Column).Value = "Audited" Or Cells(auditLabelRow.row, cFoundCell.Column).Value = "Auditados") Then auditLabelString = 1

                        currentDateString = EOMonth(CDate(currentDateString)) 'make sure you get the last day of the month
                        currentDateString = WorksheetFunction.Text(currentDateString, "YYYY-MM-DD")

                        current_field_set_id = createDataset(CInt(mfiID), CInt(currentSheetID), currentDateString, auditLabelString)

                        valueString = ""

                        Set rFoundCell = currentSheet.Cells(1, dataSheetIDColumnCell.Column)

                        ' Loop through each cell below the #fieldLabelIDColoumn

                        For fieldFindCount = 1 To WorksheetFunction.CountIf(dataSheetIDColumn, currentSheetID)
                            currentSheet.Columns(dataSheetIDColumnCell.Column).Select
                            Set rFoundCell = currentSheet.Columns(dataSheetIDColumnCell.Column).Find(What:=currentSheetID, After:=rFoundCell, _
                                    LookIn:=xlValues, LookAt:=xlWhole, _
                                    SearchDirection:=xlNext, MatchCase:=False)
                            Set currentFieldLabelID = rFoundCell.Offset(0, 1)
                            If (currentFieldLabelID.Value <> "") Then
                                currentValue = Cells(currentFieldLabelID.row, cFoundCell.Column)
                                If (Trim(currentValue) = "" Or Trim(currentValue) = "-") Then currentValue = 0
                                If (valueString <> "") Then valueString = valueString & ", "
                                valueString = valueString & "(" & currentFieldLabelID & ", " & Replace(Replace(Trim(currentValue), ",", ""), "  ", "") & ", " & current_field_set_id & ")"
                            End If
                        Next fieldFindCount

                        If (valueString <> "") Then
                            sqlString = "INSERT INTO reported_financials (field_label_id, field_value, field_set_id) VALUES " & valueString
                            'Range("a1") = sqlString
                            dbConnectionObj.Execute sqlString
                            dataRowsAdded = dataRowsAdded + 1
                        End If
                    Next dataFindCount
                End If
            Next t

        End If 'end if if that makes sure the needed update fields are found on the page
        Call handleAllGroupedRanges(currentSheet, True)
    Next wksht

    ActiveWorkbook.Sheets(1).Activate
    ImportWkbk.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ActiveWorkbook.Sheets(1).Activate

    If worksheetCount = 0 Then
        MsgBox "This document did not contain any upload sheets"
    ElseIf dataRowsAdded > 0 Then
        Dim yearArray
        yearArray = returnRawFinancialsFolder(uploadyear)

        'if you're uploading the file thats already in the shared raw financials directory then just save it as opposed to a save as
        If (Dir(ActiveWorkbook.Path, vbDirectory) = Dir(Left(yearArray(0), Len(yearArray(0)) - 1), vbDirectory)) Then
            ActiveWorkbook.Save
        Else
            'if this raw financials file isnt already saved in the raw financials folder, save a copy
            ActiveWorkbook.SaveCopyAs yearArray(0) & mfiName & yearArray(1) & ".xls"
        End If
        Call Load_Analysis_Template(mfi_id:=mfiID, mfi_name:=mfiName)
        MsgBox "Update Complete"
    End If

End Sub

Function handleOtherFinancialInformation(currentSheet As Worksheet, mfi_id As Integer, currentReportingMonth As String) As Integer
    Dim dataSheetRowLabel As Range
    Dim fieldColumnLabel As Range
    Dim currentSheetIDRange As Range
    Dim sheetIDLabel As Range
    Dim sectionStart As Range
    Dim sectionEnd As Range
    Dim currentSection As Range
    Dim numberColumn As Range

    Dim dataSectionCount As Integer

    Dim currentDataRow As Range

    Dim r As Range
    Dim s As Range
    Dim t As Integer
    Dim fieldFindCount As Integer
    Dim dataSheetArray
    Dim rFoundCell As Range

    Dim fieldSetID As Long

    Dim insertString As String

    Dim dataSetCount As Integer
    dataSetCount = 0

    'find the column with the field labels
    Set fieldColumnLabel = currentSheet.Cells.Find(What:="#OtherDataFieldColumn", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If (Not fieldColumnLabel Is Nothing) Then
        'loop through sections
        Set sectionStart = currentSheet.Cells(1, fieldColumnLabel.Column)
        Set sectionEnd = currentSheet.Cells(1, fieldColumnLabel.Column)

        For dataSectionCount = 1 To WorksheetFunction.CountIf(currentSheet.Columns(fieldColumnLabel.Column), "#DataSectionStart")

            Set sectionStart = currentSheet.Columns(fieldColumnLabel.Column).Find(What:="#DataSectionStart", After:=sectionStart, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
            Set sectionEnd = currentSheet.Columns(fieldColumnLabel.Column).Find(What:="#DataSectionEnd", After:=sectionEnd, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
            Set currentSection = Range(sectionStart, sectionEnd)

            Set sheetIDLabel = currentSection.Find(What:="#DataSheetIDRow", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
            Set currentSheetIDRange = Range(sheetIDLabel, Cells(sheetIDLabel.row, FindLastColumn(currentSheet.Cells.Rows(sheetIDLabel.row))))
            Set numberColumn = currentSheet.Cells.Find(What:="#NumberColumn", After:=Range("a1"), LookIn:=xlValues, SearchDirection:=xlNext, MatchCase:=False)
            dataSheetArray = UniqueItems(currentSheetIDRange, False)

            For t = LBound(dataSheetArray) To UBound(dataSheetArray)
                currentSheetID = dataSheetArray(t)
                If (currentSheetID <> "" And IsNumeric(currentSheetID) = True) Then
                    insertString = ""
                    Set rFoundCell = currentSheetIDRange.Cells(1, 1)
                    'createDataset
                    fieldSetID = createDataset(mfi_id, CInt(currentSheetID), currentReportingMonth, 0)
                    dataSetCount = dataSetCount + 1
                    For fieldFindCount = 1 To WorksheetFunction.CountIf(currentSheetIDRange, currentSheetID)
                        Set rFoundCell = currentSheetIDRange.Find(What:=currentSheetID, After:=rFoundCell, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
                        For Each r In currentSection
                            If (r.Value = "#DataRow") Then
                                '(field_label_id, field_value, field_set_id, field_count)
                                If (insertString <> "") Then insertString = insertString & ","
                                insertString = insertString & "(" & rFoundCell.Offset(1, 0).Value & ",'" & escapeString(currentSheet.Cells(r.row, rFoundCell.Column)) & "', " & fieldSetID & ", " & currentSheet.Cells(r.row, numberColumn.Column) & ")"
                            End If
                        Next r
                    Next fieldFindCount
                    If (insertString <> "") Then
                        Range("a1").Value = "INSERT INTO reported_other (field_label_id, field_value, field_set_id, field_count) VALUES " & insertString
                        dbConnectionObj.Execute "INSERT INTO reported_other (field_label_id, field_value, field_set_id, field_count) VALUES " & insertString
                    End If
                End If
            Next t

        Next dataSectionCount

    End If
    handleOtherFinancialInformation = dataSetCount
End Function


Function createDataset(mfi_id As Integer, currentSheetID As Integer, currentDateString As String, auditLabelString As Integer) As Long
    If (checkDBConnection = False) Then Exit Function
    Dim rs As Recordset
    Dim deleteDuplicateString As String
    Dim newFieldSetString As String
    'First delete old
    deleteDuplicateString = "DELETE FROM reporting_field_set WHERE mfi_id = " & mfi_id & " AND sheet_id = " & currentSheetID & " AND field_set_date = '" & currentDateString & "' AND is_audited = " & auditLabelString
    'InputBox "", "", deleteDuplicateString
    dbConnectionObj.Execute deleteDuplicateString

    newFieldSetString = "INSERT INTO reporting_field_set (mfi_id,sheet_id,field_set_date,is_audited) VALUES (" & mfi_id & "," & currentSheetID & ", '" & currentDateString & "', " & auditLabelString & ")"
    'InputBox "", "", newFieldSetString
    dbConnectionObj.Execute newFieldSetString

    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dbConnectionObj
    rs.Open "SELECT * FROM reporting_field_set WHERE mfi_id = " & mfi_id & " AND sheet_id = " & currentSheetID & " AND field_set_date = '" & currentDateString & "' AND is_audited = " & auditLabelString
    rs.MoveFirst
    createDataset = rs.Fields("field_set_id").Value
End Function

Sub testOther()
    Call handleOtherFinancialInformation(ActiveSheet, 10, "1996-12-31")
End Sub

Sub dumpRPT()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Call setGlobalVariabls

    Dim m As Range

    Dim mfiWkbk As Workbook
    Dim theTemplate As Workbook
    Dim mfi_range As Range

    Dim mfiName As String
    Dim regionName As String
    Dim country As String
    Dim rep_currency As String
    Dim language As String
    Dim c1_name As String
    Dim c1_email As String
    Dim c2_name As String
    Dim c2_email As String

    Set mfiWkbk = Workbooks("MFI Contact List.xlsx")
    mfiWkbk.Sheets(1).Activate
    Set mfi_range = ActiveSheet.Range(Range("A3"), Range("a3").End(xlDown))


    For Each m In mfi_range

        mfiName = m.Offset(0, 1)
        rep_currency = m.Offset(0, 2)
        regionName = m.Offset(0, 3)
        country = m.Offset(0, 4)

        language = m.Offset(0, 5)
        c1_name = m.Offset(0, 6)
        c1_email = m.Offset(0, 7)
        c2_name = m.Offset(0, 8)
        c2_email = m.Offset(0, 9)
        rm1_name = m.Offset(0, 10)
        rm1_email = m.Offset(0, 11)
        rm2_name = m.Offset(0, 12)
        rm2_email = m.Offset(0, 13)

        Set theTemplate = Workbooks.Open("C:\Users\gerren\Desktop\NRT\Template 2012.xls")
        theTemplate.Sheets(1).Range("d8") = mfiName
        theTemplate.Sheets(1).Range("d9") = country
        theTemplate.Sheets(1).Range("d10") = regionName
        theTemplate.Sheets(1).Range("d11") = c1_name
        theTemplate.Sheets(1).Range("d13") = c2_name
        theTemplate.Sheets(1).Range("d16") = language
        theTemplate.Sheets(1).Range("d19") = rep_currency

        If (c1_email <> "") Then theTemplate.Sheets(1).Hyperlinks.Add Anchor:=theTemplate.Sheets(1).Range("d12"), Address:="mailto:" & c1_email, TextToDisplay:=c1_email
        If (c2_email <> "") Then theTemplate.Sheets(1).Hyperlinks.Add Anchor:=theTemplate.Sheets(1).Range("d14"), Address:="mailto:" & c2_email, TextToDisplay:=c2_email
        If (rm1_email <> "") Then theTemplate.Sheets(1).Hyperlinks.Add Anchor:=theTemplate.Sheets(1).Range("d20"), Address:="mailto:" & rm1_email, TextToDisplay:=rm1_name
        If (rm2_email <> "") Then theTemplate.Sheets(1).Hyperlinks.Add Anchor:=theTemplate.Sheets(1).Range("d21"), Address:="mailto:" & rm2_email, TextToDisplay:=rm2_name

        theTemplate.Sheets(1).Range("d12").Font.Size = 10
        theTemplate.Sheets(1).Range("d14").Font.Size = 10
        theTemplate.Sheets(1).Range("d20").Font.Size = 10
        theTemplate.Sheets(1).Range("d21").Font.Size = 10

        theTemplate.Sheets(1).Columns("D").AutoFit
        theTemplate.Sheets(1).Protect Password:="dwm"
        theTemplate.SaveAs "C:\Users\gerren\Desktop\NRT\files\DWM 2012 Reporting Template - " & mfiName & ".xls", FileFormat:=56
        theTemplate.Close
    Next m
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub loopThroughTheseEmails()
    Dim r As Range
    Dim c As Range
    Dim isMale As Boolean
    Set r = Selection
    For Each c In r
        isMale = False
        If (c.Offset(0, 4) = "Male") Then isMale = True
        Call send_2012_NRT_email(c.Offset(0, 1), c.Offset(0, 2), c.Value, c.Offset(0, 3), isMale)
    Next c
End Sub

Sub send_NRT_email(recipName As String, recipEmail As String, mfi_name As String, lang As String, Optional isMale As Boolean)

    ' Call setGlobalVariabls

    Dim details As Range
    Dim team As Range
    Dim requestType As Range
    Dim dynamicDetailLabel As Range
    Dim dynamicDetail As Range

    Dim requestBy As Range
    Dim projectName As Range
    Dim requestSheet As Worksheet
    Dim objol As New Outlook.Application
    Dim objmail As MailItem
    Dim requestFind As Range
    Dim contactSheet As Worksheet
    Dim fileName As String
    Dim Sigstring As String
    Dim Signature As String
    Dim Fname As String


    fileName = "C:\Users\nick\Desktop\NRT\DWM Reporting Template_" & mfi_name & ".xls"
    Sigstring = "C:\Users\nick\AppData\Roaming\Microsoft\Signatures\DWM2.htm"

    If (FileExists(fileName) = False) Then Exit Sub

    Dim recips As String

    Set objol = New Outlook.Application
    Set objmail = objol.CreateItem(olmailitem)


    If Dir(Sigstring) <> "" Then
    Signature = GetBoiler(Sigstring)
    Else
    Signature = ""
    End If


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' contactSheet.Visible = xlSheetVisible
    ' contactSheet.Activate

    With objmail
    .To = recipEmail 'enter in here the email address
    '.cc = ""
    .cc = "monitoring@dwmarkets.com" 'enter in here the email address
    If lang = "English" Then
    .Subject = "2011 DWM Reporting Template"
    ElseIf lang = "Español" Then
    .Subject = "DWM Hoja de Reporte del 2011"
    End If
    .Attachments.Add fileName
    If lang = "English" Then
    .HTMLBody = englishBody(recipName) & "<br><br>" & Signature
    ElseIf lang = "Español" Then
    .HTMLBody = spanishBody(recipName, isMale) & "<br><br>" & Signature
    End If

    'vbCrLf
    .NoAging = True
    .display
    End With
    Set objmail = Nothing
    Set objol = Nothing
    SendKeys "%{s}", True 'send the email without prompts

    ' sqlQuery = "INSERT INTO portfolio_request (`team`,`type_of_request`,`request_made_by`,`project_name`, `details`, `submission_date`) VALUES ('" & team & "','" & requestType & "','" & requestBy & "','" & dynamicDetail & "','" & details & "', #" & Strings.Format(Now, "MM/dd/yyyy") & "#)"
    ' dbConnectionObj.Execute sqlQuery
    ' requestSheet.Activate
    ' ActiveWorkbook.Close
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'MsgBox "Your request was sent successfully."
End Sub

Sub send_2012_NRT_email(recipName As String, recipEmail As String, mfi_name As String, lang As String, Optional isMale As Boolean)

    Dim details As Range
    Dim team As Range
    Dim requestType As Range
    Dim dynamicDetailLabel As Range
    Dim dynamicDetail As Range

    Dim requestBy As Range
    Dim projectName As Range
    Dim requestSheet As Worksheet
    Dim objol As New Outlook.Application
    Dim objmail As MailItem
    Dim requestFind As Range
    Dim contactSheet As Worksheet
    Dim fileName As String
    Dim Sigstring As String
    Dim Signature As String
    Dim Fname As String


    fileName = "C:\Users\nick\Desktop\NRT\DWM Reporting Template_" & mfi_name & ".xls"
    Sigstring = "C:\Users\gerren\AppData\Roaming\Microsoft\Signatures\Sig 3-11-2011.htm"

    If (FileExists(fileName) = False) Then Exit Sub

    Dim recips As String

    Set objol = New Outlook.Application
    Set objmail = objol.CreateItem(olmailitem)


    If Dir(Sigstring) <> "" Then
    Signature = GetBoiler(Sigstring)
    Else
    Signature = ""
    End If


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' contactSheet.Visible = xlSheetVisible
    ' contactSheet.Activate

    With objmail
        .To = recipEmail 'enter in here the email address
        '.cc = ""
        .cc = "monitoring@dwmarkets.com" 'enter in here the email address
        If lang = "English" Then
            .Subject = "2012 DWM Reporting Template"
        ElseIf lang = "Español" Then
            .Subject = "DWM Hoja de Reporte del 2012"
        End If

        .Attachments.Add fileName

        If lang = "English" Then
            .HTMLBody = englishBody(recipName) & "<br><br>" & Signature
        ElseIf lang = "Español" Then
            .HTMLBody = spanishBody(recipName, isMale) & "<br><br>" & Signature
        End If

        'vbCrLf
        .NoAging = True
        .display
    End With

    Set objmail = Nothing
    Set objol = Nothing
    SendKeys "%{s}", True 'send the email without prompts

    ' sqlQuery = "INSERT INTO portfolio_request (`team`,`type_of_request`,`request_made_by`,`project_name`, `details`, `submission_date`) VALUES ('" & team & "','" & requestType & "','" & requestBy & "','" & dynamicDetail & "','" & details & "', #" & Strings.Format(Now, "MM/dd/yyyy") & "#)"
    ' dbConnectionObj.Execute sqlQuery
    ' requestSheet.Activate
    ' ActiveWorkbook.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'MsgBox "Your request was sent successfully."
End Sub

Function englishBody(contactName)
    englishBody = "<font size=""3"" face=""Calibri"">Dear " & contactName & ",<br><br>" & _
    "Dear ,<br><br>We would like to reinitiate communications to start 2012 as well as re-introduce the Risk & Credit Management team.<br><br>Kathryn Barrios - Chief Credit Officer<br>Alex Dyakov - Credit Specialist<br>Sandra Osborne - Credit Specialist<br>Vivek Pradhan - Credit Specialist<br>Yrenilsa Lopez - Credit Specialist<br>Nicole Reyes - Portfolio Analyst<br>TBD - Portfolio Analyst<br><br>Please refer all communications to monitoring@dwmarkets.com. All emails sent to this address will be received by all members and will be reviewed by at least one team member.<br><br> Please find attached the 2012 DWM MFI Reporting Template. We have made some minor adjustments in 2012 as we have corrected errors in the template that we were made aware of over the course of 2011.  Please do not hesitate to contact us if the instructions for completing the new template are unclear or if you have any other questions. The reporting instructions are listed below.  Thank you. <br><br>" & _
    "Reporting Instructions<br><br>1. Please indicate the month being reported by changing cell D18.<br><br>2. Within 30 days of the close of each month, please submit completed monthly worksheets for the month indicated in cell D18.<br><br>3. Information can only be registered in highlighted green cells<br><br>4. Please input all figures in your country's local currency.<br><br>5. Please use the same template every month. At the beginning of the next calendar year, DWM will send an updated reporting template.<br><br>6. If anything is unclear, please do not hesitate to contact DWM.<br><br>7. Once completed please e-mail this file to monitoring@dwmarkets.com and your primary DWM contacts.<br><br>8.  In the Portfolio & Organizational Data Tab PLEASE INCLUDE the MONTHLY Write-Off and Recoveries amounts for 2011.<br><br><br>Warm Regards,<br>" & _
    "Best,<br> Monitoring Team"
End Function

Function spanishBody(contactName, isMale As Boolean)
Dim lola As String
If isMale = True Then
spanishBody = "Estimado " & contactName & ",<br><br>"
lola = "lo"
Else
spanishBody = "Estimada " & contactName & ",<br><br>"
lola = "la"
End If
spanishBody = "<font size=""3"" face=""Calibri"">" & spanishBody & "Por favor consulte todas las comunicaciones a monitoring@dwmarkets.com. Todos los correos electrónicos enviados a esta dirección serán revisados por uno de los miembros del equipo.<br><br>Por favor encuentre adjunto la hoja de reporte de DWM para el 2012. Hemos hecho unos pequeños ajustes y corregido algunos errores. Por favor no dude en contactarnos si las instrucciones para completar el nuevo reporte no son claras o si tiene alguna pregunta. Los cambios más relevantes entre la versión anterior y la versión actual están enumerados abajo en este mensaje, juntos con las instrucciones para llenar el reporte. <br><br>" & _
"POR FAVOR TAMBIEN INCLUYA los datos MENSUALES de Castigos y Prestamos Recuperados en la hoja de Cartera y Datos de la Organización.<br><br>Adiciones<br>• Nuevas líneas añadidas a las hojas del Balance General, Estado de Resultados, Calidad de Cartera , y Datos de la Organización<br>• Hoja de Fondos y Accionistas<br>Eliminaciones<br>• Requisitos de reporte trimestral<br>• Indicadores Sociales (proporcionaremos en un futuro el reporte con indicadores sociales por separado)<br><br>Instrucciones del Reporte<br>1. Por favor indicar el mes de reporte cambiando la celda D18<br>" & _
"2. Por favor enviar completadas las hojas de trabajo mensuales para el mes indicado en la celda D18, dentro de 30 días del fin de cada mes<br>3. Información solo puede ser ingresada en las celdas subrayadas verde <br>4. Por favor proporcione todos los datos en la moneda local de su país<br>5. Por favor use el mismo archivo cada mes. Al principio de cada año calendario, DWM enviara una plantilla de reporte actualizada<br>6. Si algo no está claro, por favor no dude en llamar a DWM<br>7. Una vez completado, por favor enve este archivo a monitoring@dwmarkets.com, y copiando sus contacto principal en DWM.<br>8. Por favor incluya los datos MENSUALES de Castigos y Prestamos Recuperados para el 2011 en la hoja de Cartera y Datos de la Organización." & _
"Atentamente,<br>Monitoring Team"
End Function

Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
Dim fso As Object
Dim ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
GetBoiler = ts.readall
ts.Close
End Function

Function FileExists(FullFileName As String) As Boolean
' returns TRUE if the file exists
FileExists = Len(Dir(FullFileName)) > 0
End Function
