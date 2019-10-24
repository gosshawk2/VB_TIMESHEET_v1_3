Module Module_TIMESHEET

    Public dic_Totals As Object
    Public dic_AnyChanges As Object
    Public LabourComments As String
    Public FormID As String = ""
    Public CPFormName As String = ""
    Public CountPopulate = 0

    Sub UpdateEmployees(ByRef ErrorList() As String, Optional ByVal DeleteAllEmployeesFirst As Boolean = False,
                        Optional ByRef TotalUpdates As Long = 0, Optional ByRef TotalInsertsNeeded As Long = 0, Optional AllowSave As Boolean = False,
                        Optional ByRef TotalDescNeeded As Long = 0, Optional ByVal IncludeTGW As Boolean = False)
        'UPDATE EMPLOYEES:
        'Extract employee details from the CSV or XML or EXCEL files and insert into the MYSQL TIMESHEET SERVER:
        'Get EXCEL DATA from selected Workbook:
        Dim ExcelFilename As String
        Dim SheetName As String
        Dim StartCell As String
        Dim EndCell As String
        Dim EmployeeTable As String = "tblEmployees"
        'Dim dt As DataTable
        Dim ExtractArr(,) As Object
        Dim Fieldnames As String
        Dim FieldValues As String
        Dim RowIDX As Long
        Dim ColIDX As Integer
        Dim TotalRows As Integer

        Dim TotalFields As Integer
        Dim STATUS As String = ""
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim FullName_FromSheet As String = ""
        Dim FullName_FromDB As String = ""
        Dim FullName_FromSheet_Reversed As String = ""
        Dim FullName_FromDB_Reversed As String = ""
        Dim EmpNo As String = ""
        Dim Description As String = ""
        Dim CreationDate As String = ""
        Dim ModifiedDate As String = ""
        Dim PasswordChangeDate As String = ""
        Dim StatusCol As Integer
        Dim UsernameCol As Integer
        Dim NameCol As Integer
        Dim DescCol As Integer
        Dim CreationDateCol As Integer
        Dim ModDateCol As Integer
        Dim LastUpdated As String = ""
        Dim ErrMessage As String = ""
        Dim ExcludeFields As String = ""
        Dim FieldnameArr() As String = Nothing
        Dim Extract As String = ""
        Dim SpacePos As Integer = 0
        Dim FullName As String = ""
        Dim UpdateCriteria As String = ""
        Dim UpdateRecord As Boolean
        Dim UpdateNeeded As Boolean = False
        Dim DescriptionChanged As Boolean = False
        Dim SavedOK As Boolean
        Dim ReplaceName As Boolean
        Dim Answer As Integer
        Dim FoundDicName As Boolean
        Dim FoundFirstName As Boolean
        Dim FoundLastName As Boolean
        Dim FoundFullName As Boolean
        Dim FoundNameReversed As Boolean
        Dim FoundEmpNo As Boolean
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllReturnedValues As Object()
        Dim AllFieldNames() As String
        Dim oldDBFirstname As String = ""
        Dim oldDBLastname As String = ""
        Dim oldDBEmpNO As String = ""
        Dim oldDBDescription As String = ""
        Dim ErrIDX As Long
        Dim Empno_7Digits As String = ""
        Dim dtDateNow As DateTime
        Dim ReplaceAll As Boolean = False
        Dim TotalUpdatesNeeded As Long
        Dim TotalNameChangesNeeded As Long
        Dim dtDateTime As DateTime
        Dim Percentage As Single = 0.0F
        Dim EmployeesArr As Object = Nothing
        Dim Dic_Employees As Object
        Dim strEmpCode As String = ""
        Dim OldEmployeeCode As String = ""
        Dim NewEmployeeCode As String = ""
        Dim WholeRow As String = ""
        Dim DBEmployeeArr() As String = Nothing
        Dim ExcelSheets() As String
        Dim ReturnColour As String = ""
        Dim AREA As String
        Dim SHIFT As String
        Dim Title As String
        Dim Criteria As String = ""
        Dim strSQL As String = ""
        Dim SortFields As String = ""
        Dim FoundErrMessage As String = ""

        ReDim ErrorList(1)

        'PLEASE WAIT ...
        'CALL FORM TO ALLOW USER TO SELECT sheet with Employees on:
        SheetName = frmMainGIForm.SelectedSheet
        ExcelFilename = frmMainGIForm.SelectedWorkbookFile
        'SheetName = "Employees"
        frmMainGIForm.txtMessages.Text = "Extracting Employees from Spreadsheet ..."
        ExtractArr = ExtractExcelRows(ExcelFilename, SheetName, 1, 1, 0, 0)
        If ExtractArr Is Nothing Then
            'MsgBox("No Employees")
            Exit Sub
        End If
        TotalRows = UBound(ExtractArr, 1)
        TotalInsertsNeeded = 0
        'TotalUpdatesNeeded = TotalUpdates
        'ExtractArr(1,1) = name , ExtractArr(1,2) = alias, ExtractArr(1,6) = description
        'ExtractArr(Rows,Fields) and ExtractArr(3,0) is the first ID value.
        'ExcludeFields = "ID"
        ExcludeFields = ""
        Fieldnames = GetMyFields(EmployeeTable, frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        FieldnameArr = Split(Fieldnames, ",")
        TotalFields = UBound(ExtractArr, 2)
        ErrIDX = 0
        For ColIDX = 1 To TotalFields
            'Extract Column Positions:
            Title = ExtractArr(2, ColIDX)
            If UCase(Title) = Nothing Then
                Continue For
            Else
                If UCase(Title) = "STATUS" Then
                    StatusCol = ColIDX
                    'yields: Not in use / In use
                ElseIf UCase(Title) = "USERNAME" Then
                    UsernameCol = ColIDX
                ElseIf UCase(Title) = "NAME" Then
                    NameCol = ColIDX
                ElseIf UCase(Title) = "DESCRIPTION" Then
                    DescCol = ColIDX
                ElseIf UCase(Title) = "CREATION DATE" Then
                    CreationDateCol = ColIDX
                Else
                    Continue For
                End If
            End If
        Next
        'TotalNameChangesNeeded = CheckEmployeeUpdates(ExcelFilename, SheetName, TotalDescNeeded)

        If TotalNameChangesNeeded > 0 Then
            Answer = MsgBox("SELECT YES to be prompted for EACH update - or NO to REPLACE ALL", vbYesNoCancel, "Total Updates Required: " & CStr(TotalNameChangesNeeded) & " Prompt for EACH Update ?")
            If Answer = vbYes Then
                ReplaceAll = False
            Else
                ReplaceAll = True
            End If
        End If
        If TotalDescNeeded > 0 And TotalNameChangesNeeded = 0 Then
            Answer = MsgBox("SELECT YES to be prompted for EACH Description (JOB TITLE) update - or NO to REPLACE ALL", vbYesNoCancel, "Total Description Updates Required: " & CStr(TotalDescNeeded) & " Prompt for EACH Update ?")
            If Answer = vbYes Then
                ReplaceAll = False
            Else
                ReplaceAll = True
            End If

        End If
        Dic_Employees = CreateObject("Scripting.Dictionary")
        Dic_Employees.removeall
        Dic_Employees.comparemode = vbTextCompare
        If DeleteAllEmployeesFirst Then
            Module_DanG_MySQL_Tools.DeleteMyRecord(EmployeeTable, frmMainGIForm.myConnString, "", ErrMessage)
        End If
        strSQL = "SELECT * FROM " & EmployeeTable
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
        If Len(SortFields) > 0 Then
            strSQL = strSQL & " ORDER BY " & SortFields
        End If
        EmployeesArr = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage, 2, Dic_Employees)
        Percentage = 0.0F
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks
        frmMainGIForm.pbarMain.Value = 0
        TotalUpdates = 0
        For RowIDX = 3 To TotalRows
            FieldValues = ""
            EmpNo = "0"
            STATUS = ""
            Firstname = ""
            Lastname = ""
            FullName = ""
            FullName_FromSheet = ""
            FullName_FromDB = ""
            Description = "none"
            CreationDate = "1970-01-01"
            ModifiedDate = "1970-01-01"
            LastUpdated = "1970-01-01"
            Empno_7Digits = "0"
            AllReturnedValues = Nothing
            DescriptionChanged = False
            UpdateNeeded = False
            'EXTRACT STATUS FROM NEW SPREADSHEET:
            If RowIDX > 0 And StatusCol > 0 Then
                Extract = ExtractArr(RowIDX, StatusCol) 'From Spreadsheet STATUS = IN USE / NOT IN USE
            End If
            If IsNothing(Extract) Then
                'Continue For
                STATUS = ""
            End If
            If InStr(Extract, "#") > 0 Then
                Continue For
            Else
                STATUS = Extract
            End If
            If RowIDX > 0 And UsernameCol > 0 Then
                'EXTRACT USERNAME / EMPNO FROM NEW SPREADSHEET:
                Extract = ExtractArr(RowIDX, UsernameCol) 'From Spreadsheet Username = Employee Number / AGENCY NO. (not 7-digits)
            End If
            If IsNothing(Extract) Then
                Continue For
            Else
                EmpNo = Extract
            End If
            'EXTRACT NAME FROM NEW SPREADSHEET:
            If RowIDX > 0 And NameCol > 0 Then
                Extract = ExtractArr(RowIDX, NameCol) 'From Spreadsheet NAME - Firstname and Lastname
            End If
            If InStr(Extract, "\") > 0 Or IsNothing(Extract) Or InStr(Extract, "#") > 0 Then
                'rejected - at moment BUT the name can be extracted after \\ in some instances.
            Else
                If Len(Extract) > 0 Then
                    SpacePos = InStr(Extract, " ")
                    If SpacePos > 0 Then
                        Firstname = Mid(Extract, 1, SpacePos - 1)
                        Lastname = Mid(Extract, SpacePos + 1, Len(Extract))
                        FullName_FromSheet = Firstname & " " & Lastname
                        FullName_FromSheet_Reversed = Lastname & " " & Firstname
                    Else
                        Firstname = Extract
                        FullName_FromSheet = Firstname
                        FullName_FromSheet_Reversed = Firstname
                    End If
                End If

            End If
            'EXTRACT DESCRIPTION FROM NEW SPREADSHEET:
            If RowIDX > 0 And DescCol > 0 Then
                Description = ExtractArr(RowIDX, DescCol)
            End If
            If IsNothing(Description) Then
                Description = "NONE"
            Else
                If InStr(Description, "TGW") > 0 Then
                    If IncludeTGW = False Then
                        Continue For
                    End If
                End If
            End If
            'EXTRACT CREATION DATE FROM NEW SPREADSHEET:
            If RowIDX > 0 And CreationDateCol > 0 Then
                Extract = ExtractArr(RowIDX, CreationDateCol)
            End If
            If IsNothing(Extract) Then
                CreationDate = ""
            Else
                If Len(Extract) > 0 Then
                    If IsDate(Extract) Then
                        dtDateTime = CDate(Extract)
                        CreationDate = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    Else
                        CreationDate = "1900-01-01 01:00:00"
                    End If
                End If
            End If


            'NOW CHECK IF EMPLOYEE NUMBER already in Database Table tblEmployees :
            SearchField = "EmpNo"
            ReturnField = "ID"
            ReturnValue = ""
            'ReDim AllReturnedValues(1)
            If Dic_Employees.Exists(EmpNo) Then
                FoundDicName = True
                'Beep()
            End If

            FoundEmpNo = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblEmployees", SearchField, EmpNo, "STRING", ReturnField, ReturnValue, AllReturnedValues, AllFieldNames)
            'EmpNo is from the spreadsheet being read in. AllReturnedValues can be used to extract the information in the record from the DATABASE.
            'How do we detect if an employee is no longer agency and is now with ECP ?
            '- When the same name appears twice (even if firstname and lastname could be reversed ?)
            If FoundEmpNo Then
                TotalUpdatesNeeded = TotalUpdatesNeeded + 1
                OldEmployeeCode = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFieldNames, "EmpNo", FoundErrMessage)
                If Len(FoundErrMessage) > 0 Then
                    'MsgBox("Found Error getting Emp No: " & FoundErrMessage)
                    ErrorList(ErrIDX) = FoundErrMessage
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                    Continue For
                End If
                oldDBFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFieldNames, "Firstname", FoundErrMessage)
                If Len(FoundErrMessage) > 0 Then
                    MsgBox("Found Error getting Firstname: " & FoundErrMessage)
                    ErrorList(ErrIDX) = FoundErrMessage
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                    Continue For
                End If
                oldDBLastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFieldNames, "Lastname", FoundErrMessage)
                If Len(FoundErrMessage) > 0 Then
                    MsgBox("Found Error getting Lastname: " & FoundErrMessage)
                    ErrorList(ErrIDX) = FoundErrMessage
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                    Continue For
                End If
                FullName_FromDB = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFieldNames, "FullName", FoundErrMessage)
                If Len(FoundErrMessage) > 0 Then
                    MsgBox("Found Error getting FullName: " & FoundErrMessage)
                    ErrorList(ErrIDX) = FoundErrMessage
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                    Continue For
                End If
                '
                If UCase(Left(EmpNo, 3)) = "AGY" Then
                    Call Convert_OldEmpNo_NewEmpNo("AGENCY", EmpNo, OldEmployeeCode, Empno_7Digits)
                    'Check if name exists more than once IN DATABASE - indicates if employee is now employed by ECP and not AGENCY anymore - so blank out name against AGENCY number.

                Else 'DEFAULT - as EmpNo not start with AGY - ECP no. starts with numbers
                    Call Convert_OldEmpNo_NewEmpNo("ECP", EmpNo, OldEmployeeCode, Empno_7Digits)

                End If
            Else
                'EMPLOYEE NO. NOT FOUND - INSERT INTO DATABASE:

            End If


            SearchField = "FullName"
            ReturnField = "ID"
            ReturnValue = ""

            'COMBINE FIRSTNAME and LASTNAME FROM NEW SPREADSHEET INTO FULLNAME AND SEARCH FULLNAME FIELD IN DATABASE:

            FoundFullName = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblEmployees", SearchField, FullName_FromSheet, "STRING", ReturnField, ReturnValue, AllReturnedValues, AllFieldNames)
            If FoundFullName Then
                'if Ucase(FullName_FromSheet) = 
                If UCase(Left(EmpNo, 3)) = "AGY" Then
                    Call Convert_OldEmpNo_NewEmpNo("AGENCY", EmpNo, OldEmployeeCode, Empno_7Digits)
                    'Check if name exists more than once IN DATABASE - indicates if employee is now employed by ECP and not AGENCY anymore - so blank out name against AGENCY number.

                Else 'DEFAULT - as EmpNo not start with AGY - ECP no. starts with numbers
                    Call Convert_OldEmpNo_NewEmpNo("ECP", EmpNo, OldEmployeeCode, Empno_7Digits)

                End If
            End If


            'Already have Employee Number,

            dtDateNow = Now()
            PasswordChangeDate = "1900-01-01 01:00:00"
            LastUpdated = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
            'LastUpdated = FormatDateTime(dtDateNow, "yyyy-MM-dd HH:mm:ss") 'Not Working
            FieldValues = STATUS
            FieldValues = FieldValues & "," & EmpNo
            FieldValues = FieldValues & "," & Firstname
            FieldValues = FieldValues & "," & Lastname
            FieldValues = FieldValues & "," & FullName_FromSheet
            FieldValues = FieldValues & "," & Description
            FieldValues = FieldValues & "," & Empno_7Digits
            FieldValues = FieldValues & "," & AREA
            FieldValues = FieldValues & "," & SHIFT
            FieldValues = FieldValues & "," & PasswordChangeDate
            FieldValues = FieldValues & "," & LastUpdated
            FieldValues = FieldValues & "," & CreationDate
            FieldValues = FieldValues & "," & ModifiedDate

            FullName = Firstname & " " & Lastname
            If FoundEmpNo Then 'EmpNo already in DB:

                'oldDBFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Firstname")
                WholeRow = Dic_Employees(EmpNo)
                If Len(WholeRow) = 0 Then
                    ErrorList(ErrIDX) = "ROW is BLANK (in Dictionary Script) " & Firstname & " " & Lastname & ",EmpNo: " & EmpNo
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                    'Continue For
                End If
                DBEmployeeArr = Split(WholeRow, ",")
                oldDBFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Firstname")
                oldDBLastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Lastname")
                oldDBEmpNO = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "EmpNo")
                oldDBDescription = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Description")

                'oldDBFirstname = DBEmployeeArr(1) 'GETTING INDEX OUT OF RANGE HERE.
                'oldDBLastname = DBEmployeeArr(2)
                'oldDBEmpNO = DBEmployeeArr(3)
                'oldDBDescription = DBEmployeeArr(4)

                FullName = Firstname & " " & Lastname
                DescriptionChanged = False
                If UCase(Firstname) & " " & UCase(Lastname) = UCase(oldDBFirstname) & " " & UCase(oldDBLastname) Then
                    'NO CHANGE IN EMPLOYEE NAME
                    'UpdateNeeded = False
                    If UCase(Description) = UCase(oldDBDescription) Then
                        'NO CHANGE in Description
                        DescriptionChanged = False
                    Else
                        'The description has CHANGED
                        DescriptionChanged = True
                        UpdateNeeded = True
                        TotalDescNeeded = TotalDescNeeded + 1
                    End If
                Else
                    'The EMPLOYEE NAMES are different:
                    UpdateNeeded = True
                    TotalUpdates = TotalUpdates + 1
                End If

                If AllowSave Then

                    UpdateCriteria = "ID = " & ReturnValue
                    'ASk user if they want to replace the existing name in the database with the new name.
                    If ReplaceAll = False Then
                        If UpdateNeeded Then
                            Answer = MsgBox("Name is different in database - do you want to replace with NEW name from spreadsheet ? : " & FullName & "(" & EmpNo & ")" & " ?", vbYesNoCancel, "REPLACE EXISTING NAME")
                            If Answer = vbYes Then
                                'UPDATE database with new name and details etc
                                'Fieldnames have two extra fields - also includes the ID.
                                'SavedOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, "tblEmployees", Fieldnames, FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                                SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "",
                                        EmployeeTable, Fieldnames, FieldValues, UpdateCriteria, "ID", ErrMessage, False, ",")
                                If Not SavedOK Then
                                    ErrorList(ErrIDX) = "Name in DB: " & oldDBFirstname & " " & oldDBLastname & ",EmpNo: " & oldDBEmpNO & " ,Did NOT UPDATE"
                                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                                    ErrIDX = ErrIDX + 1
                                End If
                            ElseIf Answer = vbCancel Then
                                Exit For
                            Else 'NO
                                'leave as is.
                            End If
                        End If
                        If DescriptionChanged Then
                            Answer = MsgBox("Description (JOB TITLE) is different in database - do you want to replace with NEW description from spreadsheet ? : " & Description & " " & FullName & "(" & EmpNo & ")" & " ?", vbYesNoCancel, "REPLACE EXISTING NAME")
                            If Answer = vbYes Then
                                'UPDATE database with new name and details etc
                                'SavedOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, "tblEmployees", Fieldnames, FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                                SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "",
                                        EmployeeTable, Fieldnames, FieldValues, UpdateCriteria, "ID", ErrMessage, False, ",")
                                If Not SavedOK Then
                                    ErrorList(ErrIDX) = "Name in DB: " & oldDBFirstname & " " & oldDBLastname & ",EmpNo: " & oldDBEmpNO & " ,Did NOT UPDATE"
                                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                                    ErrIDX = ErrIDX + 1
                                End If
                            ElseIf Answer = vbCancel Then
                                Exit For
                            Else 'NO
                                'leave as is.
                            End If
                        End If
                    Else
                        'SavedOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, "tblEmployees", Fieldnames, FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                        SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "",
                                        EmployeeTable, Fieldnames, FieldValues, UpdateCriteria, "ID", ErrMessage, False, ",")
                        If Not SavedOK Then
                            ErrorList(ErrIDX) = "Name in DB: " & oldDBFirstname & " " & oldDBLastname & ",EmpNo: " & oldDBEmpNO & " ,Did NOT UPDATE"
                            ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                            ErrIDX = ErrIDX + 1
                        End If
                    End If 'ReplaceAll
                End If 'AllowSave
            Else 'Not found emp no in tblEmployees:
                TotalInsertsNeeded = TotalInsertsNeeded + 1
                If AllowSave Then
                    'INSERT NEW NAME AND DETAILS INTO DATABASE TABLE:
                    'SavedOK = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, "tblEmployees", Fieldnames, FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                    SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, "",
                                        EmployeeTable, Fieldnames, FieldValues, UpdateCriteria, "ID", ErrMessage, False, ",")
                End If
            End If
            'Update Progress control - either a progress bar or lblProgress label:
            If AllowSave = False Then
                frmMainGIForm.txtMessages.Text = "Analysing ... Rows: " & CStr(RowIDX) & "/" & CStr(TotalRows)
            Else
                frmMainGIForm.txtMessages.Text = "Updating ... Rows: " & CStr(RowIDX) & "/" & CStr(TotalRows)
            End If
            Percentage = (RowIDX / TotalRows) * 100
            frmMainGIForm.pbarMain.Value = CInt(Percentage)
            Application.DoEvents()
            'For Each cf As Form In formsopen(frmGI_RP_Userform)

            'Next
        Next

        'MsgBox("FINISHED EMPLOYEE UPDATE")
        If AllowSave Then
            frmMainGIForm.txtMessages.Text = "FINISHED EMPLOYEE UPDATE"
        Else
            frmMainGIForm.txtMessages.Text = "FINISHED EMPLOYEE CHECK"
        End If
        If Check_ImportDate_IsToday(EmployeeTable, "LastUpdated", frmMainGIForm.myConnString, ReturnColour, ErrMessage) Then
            'MsgBox("EMPLOYEES OK")
        End If
        'MsgBox("Colour = " & ReturnColour)
        If UCase(ReturnColour) = "GREEN" Then
            frmGI_RP_Userform.btnUpdateEmployees.BackColor = Color.Lime
        ElseIf UCase(ReturnColour) = "AMBER" Then
            frmGI_RP_Userform.btnUpdateEmployees.BackColor = Color.Orange
        Else
            frmGI_RP_Userform.btnUpdateEmployees.BackColor = Color.Red
        End If


        If UBound(ErrorList) > 0 Then
            'Display form to show all errors generated from above process:
            'Dim cf As New frmScrollList

            'cf.MdiParent = frmMainGIForm
            'cf.StartPosition = FormStartPosition.CenterParent
            'cf.Text = "Error List" & CStr(frmMainGIForm.MdiChildren.Count)
            'cf.Name = "Error List"
            'cf.Show()
        End If

    End Sub

    Sub UpdateNames_Using_TABLES(ByRef ErrorList() As String, Optional ByVal DeleteAllEmployeesFirst As Boolean = False,
                        Optional ByRef TotalUpdates As Long = 0, Optional ByRef TotalInsertsNeeded As Long = 0, Optional AllowSave As Boolean = False,
                        Optional ByRef TotalDescNeeded As Long = 0, Optional ByVal IncludeTGW As Boolean = False)
        'Put selected EXCEL SHEET into temporary table immediately and compare with existing employee table for differences - to get updates
        Dim ExcelFilename As String
        Dim SheetName As String
        Dim StartCell As String
        Dim EndCell As String
        Dim EmployeeTable As String = "tblEmployees"
        'Dim dt As DataTable
        Dim ExtractArr(,) As Object
        Dim Fieldnames As String
        Dim FieldValues As String
        Dim RowIDX As Long
        Dim ColIDX As Integer
        Dim TotalRows As Integer

        Dim TotalFields As Integer
        Dim STATUS As String = ""
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim FullName_FromSheet As String = ""
        Dim FullName_FromDB As String = ""
        Dim FullName_FromSheet_Reversed As String = ""
        Dim FullName_FromDB_Reversed As String = ""
        Dim EmpNo As String = ""
        Dim Description As String = ""
        Dim CreationDate As String = ""
        Dim ModifiedDate As String = ""
        Dim PasswordChangeDate As String = ""
        Dim StatusCol As Integer
        Dim UsernameCol As Integer
        Dim NameCol As Integer
        Dim DescCol As Integer
        Dim CreationDateCol As Integer
        Dim ModDateCol As Integer
        Dim LastUpdated As String = ""
        Dim ErrMessage As String = ""
        Dim ExcludeFields As String = ""
        Dim FieldnameArr() As String = Nothing
        Dim Extract As String = ""
        Dim SpacePos As Integer = 0
        Dim FullName As String = ""
        Dim UpdateCriteria As String = ""
        Dim UpdateRecord As Boolean
        Dim UpdateNeeded As Boolean = False
        Dim DescriptionChanged As Boolean = False
        Dim SavedOK As Boolean
        Dim ReplaceName As Boolean
        Dim Answer As Integer
        Dim FoundDicName As Boolean
        Dim FoundFirstName As Boolean
        Dim FoundLastName As Boolean
        Dim FoundFullName As Boolean
        Dim FoundNameReversed As Boolean
        Dim FoundEmpNo As Boolean
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllReturnedValues As Object()
        Dim AllFieldNames() As String
        Dim oldDBFirstname As String = ""
        Dim oldDBLastname As String = ""
        Dim oldDBEmpNO As String = ""
        Dim oldDBDescription As String = ""
        Dim ErrIDX As Long
        Dim Empno_7Digits As String = ""
        Dim dtDateNow As DateTime
        Dim ReplaceAll As Boolean = False
        Dim TotalUpdatesNeeded As Long
        Dim TotalNameChangesNeeded As Long
        Dim dtDateTime As DateTime
        Dim Percentage As Single = 0.0F
        Dim EmployeesArr As Object = Nothing
        Dim Dic_Employees As Object
        Dim strEmpCode As String = ""
        Dim OldEmployeeCode As String = ""
        Dim NewEmployeeCode As String = ""
        Dim WholeRow As String = ""
        Dim DBEmployeeArr() As String = Nothing
        Dim ExcelSheets() As String
        Dim ReturnColour As String = ""
        Dim AREA As String
        Dim SHIFT As String
        Dim Title As String
        Dim Criteria As String = ""
        Dim strSQL As String = ""
        Dim SortFields As String = ""
        Dim FoundErrMessage As String = ""
        Dim NewNamesTable As String = "tblNEWEmployees"
        Dim Rejected As Integer = 0

        ReDim ErrorList(1)

        'PLEASE WAIT ...
        'CALL FORM TO ALLOW USER TO SELECT sheet with Employees on:
        SheetName = frmMainGIForm.SelectedSheet
        ExcelFilename = frmMainGIForm.SelectedWorkbookFile
        'SheetName = "Employees"
        frmMainGIForm.txtMessages.Text = "Extracting Employees from Spreadsheet ..."
        ExtractArr = ExtractExcelRows(ExcelFilename, SheetName, 1, 1, 0, 0)
        If ExtractArr Is Nothing Then
            'MsgBox("No Employees")
            Exit Sub
        End If
        TotalRows = UBound(ExtractArr, 1)
        TotalInsertsNeeded = 0
        'TotalUpdatesNeeded = TotalUpdates
        'ExtractArr(1,1) = name , ExtractArr(1,2) = alias, ExtractArr(1,6) = description
        'ExtractArr(Rows,Fields) and ExtractArr(3,0) is the first ID value.
        'ExcludeFields = "ID"
        ExcludeFields = ""
        Fieldnames = GetMyFields(EmployeeTable, frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        FieldnameArr = Split(Fieldnames, ",")
        TotalFields = UBound(ExtractArr, 2)
        ErrIDX = 0
        For ColIDX = 1 To TotalFields
            'Extract Column Positions:
            Title = ExtractArr(2, ColIDX)
            If UCase(Title) = Nothing Then
                Continue For
            Else
                If UCase(Title) = "STATUS" Then
                    StatusCol = ColIDX
                    'yields: Not in use / In use
                ElseIf UCase(Title) = "USERNAME" Then
                    UsernameCol = ColIDX
                ElseIf UCase(Title) = "NAME" Then
                    NameCol = ColIDX
                ElseIf UCase(Title) = "DESCRIPTION" Then
                    DescCol = ColIDX
                ElseIf UCase(Title) = "CREATION DATE" Then
                    CreationDateCol = ColIDX
                Else
                    Continue For
                End If
            End If
        Next
        'TotalNameChangesNeeded = CheckEmployeeUpdates(ExcelFilename, SheetName, TotalDescNeeded)

        If TotalNameChangesNeeded > 0 Then
            Answer = MsgBox("SELECT YES to be prompted for EACH update - or NO to REPLACE ALL", vbYesNoCancel, "Total Updates Required: " & CStr(TotalNameChangesNeeded) & " Prompt for EACH Update ?")
            If Answer = vbYes Then
                ReplaceAll = False
            Else
                ReplaceAll = True
            End If
        End If
        If TotalDescNeeded > 0 And TotalNameChangesNeeded = 0 Then
            Answer = MsgBox("SELECT YES to be prompted for EACH Description (JOB TITLE) update - or NO to REPLACE ALL", vbYesNoCancel, "Total Description Updates Required: " & CStr(TotalDescNeeded) & " Prompt for EACH Update ?")
            If Answer = vbYes Then
                ReplaceAll = False
            Else
                ReplaceAll = True
            End If
        End If

        Dic_Employees = CreateObject("Scripting.Dictionary")
        Dic_Employees.removeall
        Dic_Employees.comparemode = vbTextCompare
        If DeleteAllEmployeesFirst Then
            Module_DanG_MySQL_Tools.DeleteMyRecord(EmployeeTable, frmMainGIForm.myConnString, "", ErrMessage)
        End If
        strSQL = "SELECT * FROM " & EmployeeTable
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
        If Len(SortFields) > 0 Then
            strSQL = strSQL & " ORDER BY " & SortFields
        End If
        'EmployeesArr = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage, 2, Dic_Employees)
        Percentage = 0.0F
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks
        frmMainGIForm.pbarMain.Value = 0
        TotalUpdates = 0
        For RowIDX = 3 To TotalRows
            FieldValues = ""
            EmpNo = "0"
            STATUS = ""
            Firstname = ""
            Lastname = ""
            FullName = ""
            FullName_FromSheet = ""
            FullName_FromDB = ""
            Description = "none"
            CreationDate = "1970-01-01"
            ModifiedDate = "1970-01-01"
            LastUpdated = "1970-01-01"
            Empno_7Digits = "0"
            AllReturnedValues = Nothing
            DescriptionChanged = False
            UpdateNeeded = False
            'EXTRACT STATUS FROM NEW SPREADSHEET:
            If RowIDX > 0 And StatusCol > 0 Then
                Extract = ExtractArr(RowIDX, StatusCol) 'From Spreadsheet STATUS = IN USE / NOT IN USE
            End If
            If IsNothing(Extract) Then
                'Continue For
                STATUS = ""
            End If
            If InStr(Extract, "#") > 0 Then
                Continue For
            Else
                STATUS = Extract
            End If
            If RowIDX > 0 And UsernameCol > 0 Then
                'EXTRACT USERNAME / EMPNO FROM NEW SPREADSHEET:
                Extract = ExtractArr(RowIDX, UsernameCol) 'From Spreadsheet Username = Employee Number / AGENCY NO. (not 7-digits)
            End If
            If IsNothing(Extract) Then
                Continue For
            Else
                EmpNo = Extract
            End If
            'EXTRACT NAME FROM NEW SPREADSHEET:
            If RowIDX > 0 And NameCol > 0 Then
                Extract = ExtractArr(RowIDX, NameCol) 'From Spreadsheet NAME - Firstname and Lastname
            End If
            If InStr(Extract, "\") > 0 Or IsNothing(Extract) Or InStr(Extract, "#") > 0 Then
                'rejected - at moment BUT the name can be extracted after \\ in some instances.
            Else
                If Len(Extract) > 0 Then
                    SpacePos = InStr(Extract, " ")
                    If SpacePos > 0 Then
                        Firstname = Mid(Extract, 1, SpacePos - 1)
                        Lastname = Mid(Extract, SpacePos + 1, Len(Extract))
                        FullName_FromSheet = Firstname & " " & Lastname
                        FullName_FromSheet_Reversed = Lastname & " " & Firstname
                    Else
                        Firstname = Extract
                        FullName_FromSheet = Firstname
                        FullName_FromSheet_Reversed = Firstname
                    End If
                End If

            End If
            'EXTRACT DESCRIPTION FROM NEW SPREADSHEET:
            If RowIDX > 0 And DescCol > 0 Then
                Description = ExtractArr(RowIDX, DescCol)
            End If
            If IsNothing(Description) Then
                Description = "NONE"
            Else
                If InStr(Description, "TGW") > 0 Then
                    If IncludeTGW = False Then
                        Continue For
                    End If
                End If
            End If
            'EXTRACT CREATION DATE FROM NEW SPREADSHEET:
            If RowIDX > 0 And CreationDateCol > 0 Then
                Extract = ExtractArr(RowIDX, CreationDateCol)
            End If
            If IsNothing(Extract) Then
                CreationDate = ""
            Else
                If Len(Extract) > 0 Then
                    If IsDate(Extract) Then
                        dtDateTime = CDate(Extract)
                        CreationDate = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    Else
                        CreationDate = "1900-01-01 01:00:00"
                    End If
                End If
            End If

            dtDateNow = Now()
            PasswordChangeDate = "1900-01-01 01:00:00"
            LastUpdated = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
            'LastUpdated = FormatDateTime(dtDateNow, "yyyy-MM-dd HH:mm:ss") 'Not Working
            FieldValues = STATUS
            FieldValues = FieldValues & "," & EmpNo
            FieldValues = FieldValues & "," & Firstname
            FieldValues = FieldValues & "," & Lastname
            FieldValues = FieldValues & "," & FullName_FromSheet
            FieldValues = FieldValues & "," & Description
            FieldValues = FieldValues & "," & Empno_7Digits
            FieldValues = FieldValues & "," & AREA
            FieldValues = FieldValues & "," & SHIFT
            FieldValues = FieldValues & "," & PasswordChangeDate
            FieldValues = FieldValues & "," & LastUpdated
            FieldValues = FieldValues & "," & CreationDate
            FieldValues = FieldValues & "," & ModifiedDate

            FullName = Firstname & " " & Lastname

            SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, "",
                                        NewNamesTable, Fieldnames, FieldValues, UpdateCriteria, "ID", ErrMessage, False, ",")
            If Not SavedOK Then
                Rejected = Rejected + 1
            End If

        Next

    End Sub
    Sub UpdateSuppliers()
        'dont need. read in from Chetans spreadsheet.
    End Sub

    Sub Populate_DeliveryDateCombo(Formname As String, strStartDeliveryDate As String, strEndDeliveryDate As String)
        'read from tblDeliveryInfo - DeliveryDates to populate the combo on the form:
        Dim myDeliveryRefs As Object(,)
        Dim DBTable As String
        Dim strSQL As String
        Dim ErrMessage As String = ""
        Dim DeliveryRef As String = ""
        Dim RowIDX As Long
        Dim Criteria As String = ""
        Dim IDX As Long
        Dim comboArr() As String
        Dim dtStartDeliveryDate As DateTime
        Dim dtEndDeliveryDate As DateTime

        DBTable = "tblDeliveryInfo"
        If Len(strStartDeliveryDate) > 0 Then
            dtStartDeliveryDate = CDate(strStartDeliveryDate)
            If Len(strEndDeliveryDate) > 0 Then
                dtEndDeliveryDate = CDate(strEndDeliveryDate)
                Criteria = "DeliveryDate between " & Chr(34) & dtStartDeliveryDate.ToString("yyyy-MM-dd") & " AND " & Chr(34) & dtEndDeliveryDate.ToString("yyyy-MM-dd") & Chr(34)
            Else
                Criteria = "DeliveryDate = " & Chr(34) & dtStartDeliveryDate.ToString("yyyy-MM-dd") & Chr(34)
            End If
            strSQL = "SELECT DeliveryReference FROM " & DBTable
            If Len(Criteria) > 0 Then
                strSQL = strSQL & " WHERE " & Criteria
            End If
            'NOT GETTING all the Deliveryreferences - just 2 nothings. ?????????????????
            myDeliveryRefs = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage)
            If myDeliveryRefs IsNot Nothing Then
                'start at (0,0)
                ReDim comboArr(UBound(myDeliveryRefs, 2))
                For RowIDX = 0 To UBound(myDeliveryRefs, 2)
                    DeliveryRef = myDeliveryRefs(0, RowIDX)
                    If DeliveryRef IsNot Nothing Then
                        comboArr(RowIDX) = DeliveryRef

                        'Me.comDeliveryRef.Items.Add(DeliveryRef)
                    End If
                Next
                'Need to sort the comboArr() before assigning to COMBO BOX:
                Array.Sort(comboArr)
                frmMainGIForm.InsertValueIntoForm(Formname, "comDeliveryRef", "", comboArr)

            Else
                MsgBox("Delivery Date Not in database")
            End If
        Else
            Criteria = ""
        End If


    End Sub

    Sub Populate_ASNCombo(Formname As String, strStartDeliveryDate As String, strEndDeliveryDate As String)
        'read from tblDeliveryInfo - DeliveryDates to populate the combo on the form:
        Dim myASNNumbers As Object(,)
        Dim DBTable As String
        Dim strSQL As String
        Dim ErrMessage As String = ""
        Dim ASNNo As String = ""
        Dim RowIDX As Long
        Dim Criteria As String = ""
        Dim IDX As Long
        Dim comboArr() As String
        Dim dtStartDeliveryDate As DateTime
        Dim dtEndDeliveryDate As DateTime

        DBTable = "tblDeliveryInfo"
        If Len(strStartDeliveryDate) > 0 Then
            dtStartDeliveryDate = CDate(strStartDeliveryDate)
            If Len(strEndDeliveryDate) > 0 Then
                dtEndDeliveryDate = CDate(strEndDeliveryDate)
                Criteria = "DeliveryDate between " & Chr(34) & dtStartDeliveryDate.ToString("yyyy-MM-dd") & " AND " & Chr(34) & dtEndDeliveryDate.ToString("yyyy-MM-dd") & Chr(34)
            Else
                Criteria = "DeliveryDate = " & Chr(34) & dtStartDeliveryDate.ToString("yyyy-MM-dd") & Chr(34)
            End If
            strSQL = "SELECT ASN_Number FROM " & DBTable
            If Len(Criteria) > 0 Then
                strSQL = strSQL & " WHERE " & Criteria
            End If
            'NOT GETTING all the Deliveryreferences - just 2 nothings. ?????????????????
            myASNNumbers = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage)
            If myASNNumbers IsNot Nothing Then
                'start at (0,0)
                ReDim comboArr(UBound(myASNNumbers, 2))
                For RowIDX = 0 To UBound(myASNNumbers, 2)
                    ASNNo = myASNNumbers(0, RowIDX)
                    If ASNNo IsNot Nothing Then
                        comboArr(RowIDX) = ASNNo

                        'Me.comDeliveryRef.Items.Add(DeliveryRef)
                    End If
                Next
                Array.Sort(comboArr)
                frmMainGIForm.InsertValueIntoForm(Formname, "comASNNo", "", comboArr)
            Else
                MsgBox("Delivery Date not in database")
            End If
        Else
            Criteria = ""
        End If


    End Sub

    Function Populate_Name_Dropdown(Optional ByRef ComboResult As Control = Nothing) As ComboBox
        Dim Employees As Object(,)
        Dim DBTable As String
        Dim strSQL As String
        Dim ErrMessage As String = ""
        Dim Name As String = ""
        Dim RowIDX As Long
        Dim ComboIDX As Long
        Dim Criteria As String = ""
        Dim IDX As Long
        Dim comboArr As String()
        Dim dtStartDeliveryDate As DateTime
        Dim dtEndDeliveryDate As DateTime
        Dim SortField As String
        Dim Firstname As String
        Dim Lastname As String
        Dim CountNothing As Integer
        Dim ResultCombo As New ComboBox

        Populate_Name_Dropdown = Nothing
        DBTable = "tblEmployees"
        Criteria = ""
        SortField = "Firstname"
        strSQL = "SELECT * FROM " & DBTable
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
        strSQL = strSQL & " ORDER By " & SortField
        'NOT GETTING all the Deliveryreferences - just 2 nothings. ?????????????????
        Employees = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage)
        ResultCombo.Items.Clear()

        If Employees IsNot Nothing Then
            'start at (0,0)
            'ReDim comboArr(UBound(Employees, 2))
            ReDim comboArr(0)
            ComboIDX = 0
            For RowIDX = 0 To UBound(Employees, 2)
                Firstname = Employees(3, RowIDX)
                Lastname = Employees(4, RowIDX)
                If Len(Firstname) > 3 Then
                    Name = Firstname
                End If
                If Len(Lastname) > 3 Then
                    Name = Name & " " & Lastname
                End If
                If Len(Firstname) < 4 And Len(Lastname) > 3 Then
                    Name = Lastname
                End If
                If Name IsNot Nothing And Len(Name) > 0 Then
                    comboArr(ComboIDX) = Name
                    ResultCombo.Items.Add(Name)
                    ReDim Preserve comboArr(UBound(comboArr) + 1)
                    ComboIDX = ComboIDX + 1
                    'Me.comDeliveryRef.Items.Add(DeliveryRef)
                Else
                    'Nothing:
                    CountNothing = CountNothing + 1
                End If
            Next
            Array.Sort(comboArr)
            ComboResult = DirectCast(ResultCombo, System.Windows.Forms.ComboBox)
            'ResultCombo.Items.AddRange(comboArr)
            Populate_Name_Dropdown = ResultCombo
            'frmMainGIForm.InsertValueIntoForm(Formname, "comASNNo", "", comboArr)
        Else
            MsgBox("Delivery Date not in database")
        End If



    End Function



    Function GetWorksheet(ExcelSheets() As Object) As String
        Dim SelectedWorksheet As String = ""
        Dim NewForm As New frmSelectSheet
        Dim SheetName As String

        GetWorksheet = ""

        Call frmMainGIForm.ShowForms("frmSelectSheet", ExcelSheets)
        SelectedWorksheet = frmMainGIForm.SelectedSheet
        'SelectedWorksheet = NewForm.SelectedWorksheet
        GetWorksheet = SelectedWorksheet
    End Function

    Function CheckEmployeeUpdates(ByVal ExcelFilename As String, ByVal Sheetname As String, Optional ByRef NumDescUpdates As Long = 0) As Long
        'UPDATE EMPLOYEES:
        'Extract employee details from the CSV or XML or EXCEL files and insert into the MYSQL TIMESHEET SERVER:
        'Get EXCEL DATA from selected Workbook:
        Dim ExtractArr(,) As Object
        Dim Fieldnames As String
        Dim FieldValues As String
        Dim RowIDX As Long
        Dim TotalRows As Long
        Dim TotalFields As Long
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim EmpNo As String = ""
        Dim description As String = ""
        Dim LastPWChange As String = ""
        Dim LastUpdated As String = ""
        Dim ErrMessage As String = ""
        Dim ExcludeFields As String = ""
        Dim FieldnameArr() As String = Nothing
        Dim Extract As String = ""
        Dim SpacePos As Integer = 0
        Dim FullName As String = ""
        Dim UpdateCriteria As String = ""
        Dim FoundName As Boolean
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllReturnedValues As Object()
        Dim oldDBFirstname As String = ""
        Dim oldDBLastname As String = ""
        Dim oldDBEmpNO As String = ""
        Dim oldDBDescription As String = ""
        Dim ErrIDX As Long
        Dim Empno_7Digits As String = ""
        Dim dtDateNow As DateTime
        Dim ReplaceAll As Boolean = False
        Dim NumNameUpdates As Long
        Dim dtDateTime As DateTime
        Dim Percentage As Single = 0.0F


        CheckEmployeeUpdates = 0
        Percentage = 0.0F
        NumNameUpdates = 0
        NumDescUpdates = 0
        'PLEASE WAIT ...
        ExtractArr = ExtractExcelRows(ExcelFilename, Sheetname, 1, 1, 0, 0)
        TotalRows = UBound(ExtractArr, 1)

        'ExtractArr(1,1) = name , ExtractArr(1,2) = alias, ExtractArr(1,6) = description
        'ExtractArr(Rows,Fields) and ExtractArr(3,0) is the first ID value.
        ExcludeFields = "ID"
        Fieldnames = GetMyFields("tblEmployees", frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        FieldnameArr = Split(Fieldnames, ",")
        TotalFields = UBound(ExtractArr, 2)
        ErrIDX = 0
        frmMainGIForm.pbarMain.Value = 0
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Continuous
        For RowIDX = 2 To TotalRows
            FieldValues = ""
            EmpNo = "0"
            Firstname = ""
            Lastname = ""
            description = "none"
            LastPWChange = "1970-01-01"
            LastUpdated = "1970-01-01"
            Empno_7Digits = "0"
            AllReturnedValues = Nothing
            Extract = ExtractArr(RowIDX, 1) 'NAME = EMPNO
            If InStr(Extract, "#") > 0 Then
                Continue For
            Else
                EmpNo = Extract
            End If
            Extract = ExtractArr(RowIDX, 2) 'alias = FULL NAME
            If InStr(Extract, "\") > 0 Or IsNothing(Extract) Then
                'rejected
            Else
                If Len(Extract) > 0 Then
                    SpacePos = InStr(Extract, " ")
                    If SpacePos > 0 Then
                        Firstname = Mid(Extract, 1, SpacePos - 1)
                        Lastname = Mid(Extract, SpacePos + 1, Len(Extract))
                    Else
                        Firstname = Extract
                    End If
                End If
            End If
            description = ExtractArr(RowIDX, 6)
            If IsNothing(description) Then
                description = "none"
            End If
            Extract = ExtractArr(RowIDX, 3)
            Extract = Mid(Extract, 1, Len(Extract) - 4)
            If Len(Extract) > 0 Then
                If IsDate(Extract) Then
                    dtDateTime = CDate(Extract)
                    LastPWChange = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                Else
                    LastPWChange = ""
                End If
            End If
            SearchField = "EmpNo"
            ReturnField = "ID"
            ReturnValue = ""
            'ReDim AllReturnedValues(1)
            FoundName = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblEmployees", SearchField, EmpNo, "STRING", ReturnField, ReturnValue, AllReturnedValues)

            If FoundName Then
                oldDBFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Firstname")
                oldDBLastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Lastname")
                oldDBEmpNO = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "EmpNo")
                oldDBDescription = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, FieldnameArr, "Description")

                If UCase(Firstname) & " " & UCase(Lastname) = UCase(oldDBFirstname) & " " & UCase(oldDBLastname) Then
                    'No CHANGE to NAME required.
                Else
                    'The NAME is different in the database:
                    NumNameUpdates = NumNameUpdates + 1
                End If

                If UCase(description) = UCase(oldDBDescription) Then
                    'No CHANGE to DESCRIPTION.
                Else
                    'Description is different in the database:
                    NumDescUpdates = NumDescUpdates + 1
                End If

                'ASk user if they want to replace the existing name in the database with the new name.
            End If

            frmMainGIForm.txtMessages.Text = "Extracting Total Updates Needed ..."
            Percentage = (RowIDX / TotalRows) * 100
            frmMainGIForm.pbarMain.Value = CInt(Percentage)
            Application.DoEvents()


        Next
        frmMainGIForm.pbarMain.Style = ProgressBarStyle.Blocks

        CheckEmployeeUpdates = NumNameUpdates

    End Function

    Function Check_ImportDataUpdates(ExtractArray As Object(,)) As Long
        Dim TotalRows As Long
        Dim RowIDX As Long
        Dim Percentage As Single = 0.0F
        Dim strDeliveryRef As String
        Dim TotalUpdateRows As Long
        Dim SearchField As String = ""
        Dim SearchValue As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim FoundDeliveryRef As Boolean
        Dim AllReturnedValues As Object()
        Dim tblDeliveryInfoArr As Object(,)
        Dim strSQL As String
        Dim ErrMessage As String = ""
        Dim Criteria As String
        Dim dtDeliveryDate As DateTime

        Check_ImportDataUpdates = 0
        ReDim tblDeliveryInfoArr(1, 1)
        strSQL = "SELECT * FROM tblDeliveryInfo"
        tblDeliveryInfoArr = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage)
        TotalRows = UBound(ExtractArray, 1)
        Percentage = 0
        TotalUpdateRows = 0
        For RowIDX = 2 To TotalRows
            strDeliveryRef = ExtractArray(RowIDX, 7)
            If IsNothing(strDeliveryRef) Then
                Continue For
            End If
            SearchField = "DeliveryReference"
            SearchValue = strDeliveryRef
            ReturnField = "ID"
            ReturnValue = ""
            'ReDim AllReturnedValues(1)
            'Search for Delivery Date and Delivery Reference together here - use criteria:
            FoundDeliveryRef = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblDeliveryInfo", SearchField, SearchValue, "STRING", ReturnField, ReturnValue, AllReturnedValues)
            If FoundDeliveryRef Then
                TotalUpdateRows = TotalUpdateRows + 1
            End If

            Percentage = (RowIDX / TotalRows) * 100
            frmMainGIForm.pbarMain.Value = CInt(Percentage)
            frmMainGIForm.txtMessages.Text = "Extracting Total Updates Needed ... Row " & CStr(RowIDX) & "/" & CStr(TotalRows) & " : " & CStr(Percentage) & "%"
            Application.DoEvents()
        Next

        Check_ImportDataUpdates = TotalUpdateRows

    End Function

    Function MatchArray(obj1 As Object(,), obj2 As Object(,)) As Boolean


        MatchArray = False

    End Function

    Sub Sort2DArray(SortbyDim1 As Boolean, ByRef str2DArray(,) As Object, SortIDX As Long, Optional SortFieldType As String = "STRING")
        Dim tempDate1 As Date
        Dim tempDate2 As Date
        Dim tempStr As String
        Dim SortDim As Integer

        If SortbyDim1 Then
            SortDim = 2 'Used if str2DArray(sortIDX,LoopIDX) where for loopidx = 0 to ubound(str2DArray,sortdim)
        Else
            SortDim = 1 'Used if str2DArray(LoopIDX,sortIDX) where for loopidx = 0 to ubound(str2DArray,sortdim)
        End If
        If UCase(SortFieldType) = "DATE" Then
            For outer = 0 To (UBound(str2DArray, SortDim) - 1)
                For inner = 0 To UBound(str2DArray, SortDim)
                    tempDate1 = CDate(str2DArray(inner, SortIDX))
                    tempDate2 = CDate(str2DArray(outer, SortIDX))
                    If tempDate1 > tempDate2 Then
                        tempStr = str2DArray(outer, SortIDX)
                        str2DArray(outer, SortIDX) = str2DArray(inner, SortIDX)
                        str2DArray(inner, SortIDX) = tempStr
                    End If
                Next
            Next
        End If

    End Sub

    Sub BinarySearch_2DArray(arr As Object, ByVal Search As String)
        Dim ii As Integer
        Dim lb As Integer
        Dim ub As Integer
        Dim Found As Boolean = False
        Dim Middle As Integer
        Dim First As Integer
        Dim Last As Integer
        Dim FirstItem As Integer
        Dim LastItem As Integer
        Dim numMatching As Integer


        Do
            Middle = CInt((First + Last) / 2) 'calcuate the middle position of the scope

            If arr(Middle, 1).ToLower.StartsWith(Search) Then 'if the middle name starts with the search String
                Found = True 'name was found
            ElseIf arr(Middle, 1).ToLower < Search Then 'if the search name comes after the middle position's value
                First = Middle + 1 'move the first position to be 1 below the middle
            ElseIf arr(Middle, 1).ToLower > Search Then 'if the search name comes before the middle position's value
                Last = Middle - 1 'move the last position to be 1 above the middle
            End If

        Loop Until First > Last Or Found = True 'loop until the name is not found or the name is found

        If Found = True Then
            lb = 0
            ub = UBound(arr)
            If Middle > lb Then
                ii = Middle
                While arr(ii, 1).ToLower.StartsWith(Search)
                    ii = ii - 1
                    If ii < lb Then Exit While
                End While
                FirstItem = ii + 1
            Else
                FirstItem = lb
            End If

            ii = Middle + 1
            If Middle <= ub Then
                ii = Middle
                While arr(ii, 1).ToLower.StartsWith(Search)
                    ii = ii + 1
                    If ii > ub Then Exit While
                End While
                LastItem = ii - 1
            Else
                LastItem = ub
            End If

            numMatching = LastItem - FirstItem + 1

            Dim matchingItems(1, 1)
            ReDim matchingItems(0 To numMatching, 0 To 4)

            For ii = 1 To numMatching
                For jj = 0 To 4
                    matchingItems(ii, jj) = arr(FirstItem + ii - 1, jj).padLeft(15)
                Next jj
            Next ii
        End If

        'It broke at matchingContacts(ii, jj) = contacts(firstContact + ii - 1, jj).padLeft(15) and jj exceeded bounds. 
        'I would also change ReDim matchingContacts(1 To numMatching, 0 To 4) to ReDim matchingContacts(numMatching - 1,5) because there Is an error there. –

    End Sub

    Sub ImportData(ByRef ErrorList() As String, ByVal ImportDate As String, Optional ByVal LastImportDate As String = "")
        'UPDATE EMPLOYEES:
        'Extract employee details from the CSV or XML or EXCEL files and insert into the MYSQL TIMESHEET SERVER:
        'Get EXCEL DATA from selected Workbook:
        Dim ExcelFilename As String
        Dim SheetName As String
        Dim StartCell As String
        Dim EndCell As String
        Dim StartRow As Long
        Dim StartCol As Long
        Dim SortCol1 As Long
        Dim SortCol2 As Long
        Dim CheckCol As Long 'First FULL column in range thats filled to calc total Rows.
        Dim CheckRow As Long 'First Full ROW in range thats filled to calc total Columns.
        Dim ExtractArr(,) As Object
        Dim DeliveryInfoFieldnames As String
        Dim DeliveryInfoFieldsArr() As String = Nothing
        Dim DailySheetFieldnames As String
        Dim DailySheetFieldsArr() As String = Nothing
        Dim DeliveryInfo_FieldValues As String
        Dim Daily_FieldValues As String
        Dim InboundFieldnames As String
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim ASN As String
        Dim strSupplierCode As String
        Dim strSupplierName As String
        Dim strDeliveryType As String
        Dim strPalletsDue As String
        Dim strCartonsDue As String
        Dim strOrigin As String
        Dim strShift As String
        Dim strDestination As String
        Dim strSpecialInstructions As String
        Dim strDueTime As String
        Dim strLines As String 'no 14
        Dim strCases As String 'no 15
        Dim strActualCases As String 'no. 16
        Dim strExpectedTotes As String
        Dim strExpectedCages As String
        Dim strExpectedPallets As String
        Dim strReadyLabel As String
        Dim strPalletise As String
        Dim strCollar As String
        Dim strWrapStrap As String
        Dim strBagging As String
        Dim strManHR As String
        Dim RowIDX As Long
        Dim ColIDX As Long
        Dim TotalRows As Long
        Dim TotalFields As Long
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim EmpNo As String = ""
        Dim description As String = ""
        Dim LastPWChange As String = ""
        Dim LastUpdated As String = ""
        Dim ErrMessage As String = ""
        Dim ExcludeFields As String = ""
        Dim FieldnameArr() As String = Nothing
        Dim Extract As String = ""
        Dim SpacePos As Integer = 0
        Dim FullName As String = ""
        Dim UpdateCriteria As String = ""
        Dim UpdateRecord As Boolean = False
        Dim UpdateNeeded As Boolean = False
        Dim DescriptionChanged As Boolean = False
        Dim SavedOK As Boolean
        Dim ReplaceName As Boolean = False
        Dim Answer As Integer
        Dim FoundDeliveryDate As Boolean = False
        Dim FoundDeliveryRef As Boolean = False
        Dim SearchField As String
        Dim SearchText As String = ""
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllReturnedValues As Object()
        Dim oldDBDeliveryDate As String
        Dim oldDBDeliveryReference As String
        Dim ErrIDX As Long
        Dim Empno_7Digits As String = ""
        Dim dtDateNow As DateTime
        Dim ReplaceAll As Boolean = False
        Dim TotalUpdatesNeeded As Long = 0
        Dim TotalDailyChangesNeeded As Long = 0
        Dim TotalDeliveryInfoChangesNeeded As Long = 0
        Dim TotalDescNeeded As Long
        Dim dtDateTime As DateTime
        Dim Percentage As Single = 0.0F
        Dim strASNNo As String
        Dim strExpectedCases As String
        Dim strExpectedLines As String
        Dim strEstimatedPallets As String
        Dim strEstimatedCages As String
        Dim strEstimatedTotes As String
        Dim SortFieldIDX As Long 'Position within field list of field we want to sort by
        Dim strDeliveryComments As String
        Dim TotalUpdatesFound As Long = 0
        Dim Updates As Long = 0
        Dim TempFilename As String
        Dim dtDueTime As DateTime
        Dim dblDueTime As Double

        ReDim ErrorList(1)
        'Get user to select location of DAILY spreadsheet within KPI INBOUND workbook:
        'Needs sorting in reverse Delivery Date order to catch the LATEST Delivery record REF as any Delivery not unloaded before and turned away will have exactly the same 
        '      details again but on a later date.
        ExcelFilename = BrowseFilename(1)
        If Len(ExcelFilename) < 2 Then
            Exit Sub
        End If
        TempFilename = Application.StartupPath & "\TempFilename.xls"
        'PLEASE WAIT ...
        'Get the row on which the required date starts:

        SheetName = "Daily"
        StartRow = 2
        StartCol = 1
        SortCol1 = 6
        SortCol2 = 7
        CheckRow = 2
        CheckCol = 7
        frmMainGIForm.txtMessages.Text = "Please Wait ... Importing Data"
        ReDim ExtractArr(1, 1)
        If Len(ImportDate) > 0 Then
            SearchText = ImportDate
        Else
            SearchText = Now().ToString("yyyy-MM-dd")
        End If
        'what if a DATE RANGE has been selected ?
        'Sheet is sorted form latest date to earliest.
        'Call DanG_EXCEL_Module.SortSheet(ExcelFilename, SheetName, StartRow, StartCol, SortCol1, ExtractArr, CheckRow, CheckCol, SortCol2, True, True, TempFilename)
        ExtractArr = ExtractExcelRange(ExcelFilename, SheetName, SearchText, 6, "DATE", TempFilename, True, 1, 2, 0, 0, Nothing, LastImportDate)
        'MsgBox("OK - First row extracted: " & ExtractArr(1, 1))
        If ExtractArr Is Nothing Then
            MsgBox("No records Returned - Date Not Found")
            Exit Sub
        End If
        frmMainGIForm.txtMessages.Text = "Please Wait ... Data Extracted"
        TotalRows = UBound(ExtractArr, 1)
        ExcludeFields = ""
        DailySheetFieldnames = GetMyFields("tblDailySheet", frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        DailySheetFieldsArr = Split(DailySheetFieldnames, ",")
        DeliveryInfoFieldnames = GetMyFields("tblDeliveryInfo", frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        DeliveryInfoFieldsArr = Split(DeliveryInfoFieldnames, ",")
        InboundFieldnames = DeliveryInfoFieldsArr(1)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(2)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(32)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(3)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(4)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(5)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(6)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(7)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(8)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(9)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(10)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(11)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(12)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(13)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(14)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(15)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(16)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(17)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(18)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(22)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(25)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(26)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(27) 'SHIFT = 28 for value position.
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(28)
        TotalFields = UBound(ExtractArr, 2)
        SortFieldIDX = 6 'DeliveryDate as ID=0 
        'Call Sort2DArray(False, ExtractArr, SortFieldIDX)
        ErrIDX = 0
        'TotalDailyChangesNeeded  = CheckEmployeeUpdates(ExcelFilename, SheetName, TotalDescNeeded)
        'TotalUpdatesFound = Check_ImportDataUpdates(ExtractArr)
        If TotalUpdatesFound > 0 Then
            Answer = MsgBox("FOUND " & CStr(TotalUpdatesFound) & " Updates Found. SELECT YES to be prompted for EACH update - or NO to REPLACE ALL", vbYesNoCancel, "Total Updates Required: " & CStr(TotalUpdatesFound) & " Prompt for EACH Update ?")
            If Answer = vbYes Then
                ReplaceAll = False
            ElseIf Answer = vbNo Then
                ReplaceAll = True
            Else
                Exit Sub
            End If

        End If
        Percentage = 0
        Updates = 0
        frmMainGIForm.txtMessages.Text = "Please Wait ... Updating database"
        For RowIDX = 2 To TotalRows
            DeliveryInfo_FieldValues = ""
            Daily_FieldValues = ""
            strDeliveryComments = ""
            AllReturnedValues = Nothing
            DescriptionChanged = False
            UpdateNeeded = False
            Extract = ExtractArr(RowIDX, 7) 'Delivery Ref
            If IsNothing(Extract) Then
                Continue For
            End If
            strDeliveryRef = Extract
            Extract = ExtractArr(RowIDX, 6) 'Delivery Date
            If IsNothing(Extract) Then
                strDeliveryDate = "1970-01-01"
            Else
                If Len(Extract) > 0 Then
                    If IsDate(Extract) Then
                        dtDateTime = CDate(Extract)
                        'Could combine with DUE TIME ?????
                        strDeliveryDate = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    Else
                        strDeliveryDate = "1970-01-01"
                    End If
                Else
                    strDeliveryDate = "1970-01-01"
                End If
            End If
            strDeliveryType = ExtractArr(RowIDX, 1)
            If UCase(strDeliveryType) = "PRE-LABELLED" Then
                strReadyLabel = "YES"
            Else
                strReadyLabel = "NO"
            End If
            strWrapStrap = "NO"
            strPalletise = "NO"
            strCollar = "NO"
            strBagging = "NO"
            strSupplierCode = ExtractArr(RowIDX, 2)
            strSupplierName = ExtractArr(RowIDX, 3)
            strPalletsDue = ExtractArr(RowIDX, 4)
            strCartonsDue = ExtractArr(RowIDX, 5)

            strOrigin = ExtractArr(RowIDX, 8)
            strSpecialInstructions = ExtractArr(RowIDX, 9)
            strASNNo = ExtractArr(RowIDX, 10)
            strShift = ExtractArr(RowIDX, 11)
            strDestination = ExtractArr(RowIDX, 12)
            strDueTime = ExtractArr(RowIDX, 13)
            If IsNumeric(strDueTime) Then
                dblDueTime = CDbl(strDueTime)
                strDueTime = DateTime.FromOADate(ExtractArr(RowIDX, 13)).ToString("HH:mm:ss")
                'DueTime = dblDueTime.ToString("HH:mm:ss")
            ElseIf IsDate(strDueTime) Then
                dtDueTime = CDate(strDueTime)
                strDueTime = dtDueTime.ToString("HH:mm:ss")
            Else
                'leave as STRING
                'leave as is - usually a message like UNBOOKED entered - even though supposed to be a TIME column ???
            End If
            strExpectedLines = ExtractArr(RowIDX, 14)
            strExpectedCases = ExtractArr(RowIDX, 15)
            strActualCases = ExtractArr(RowIDX, 16)
            strEstimatedTotes = ExtractArr(RowIDX, 17)
            strEstimatedCages = ExtractArr(RowIDX, 18)
            strEstimatedPallets = ExtractArr(RowIDX, 19)
            strManHR = ExtractArr(RowIDX, 20)

            SearchField = "DeliveryReference"
            SearchValue = strDeliveryRef
            ReturnField = "ID"
            ReturnValue = ""
            'ReDim AllReturnedValues(1)
            'Search for Delivery Date and Delivery Reference together here - use criteria:
            FoundDeliveryRef = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblDeliveryInfo", SearchField, SearchValue, "STRING", ReturnField, ReturnValue, AllReturnedValues)

            Empno_7Digits = EmpNo
            dtDateNow = Now()
            LastUpdated = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
            'LastUpdated = FormatDateTime(dtDateNow, "yyyy-MM-dd HH:mm:ss") 'Not Working
            'FOR tblDeliveryInfo first:
            DeliveryInfo_FieldValues = Chr(34) & strDeliveryDate & Chr(34) '1
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strDeliveryRef & Chr(34) '2
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strSupplierCode & Chr(34) '32
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strSupplierName & Chr(34) '4
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strASNNo & Chr(34) '5
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strExpectedCases & Chr(34) '6
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strExpectedLines & Chr(34) '7
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strEstimatedPallets & Chr(34) '8
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strEstimatedCages & Chr(34) '9
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strEstimatedTotes & Chr(34) '10
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strCartonsDue & Chr(34) '11
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strPalletsDue & Chr(34) '12
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strReadyLabel & Chr(34) '13

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strPalletise & Chr(34) '14
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strWrapStrap & Chr(34) '15
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strCollar & Chr(34) '16
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strBagging & Chr(34) '17

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strActualCases & Chr(34) '18
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strManHR & Chr(34) '19
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & LastUpdated & Chr(34)

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strOrigin & Chr(34) '26
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strDueTime & Chr(34) '27
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strShift & Chr(34) '28
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & Chr(34) & strDeliveryComments & Chr(34) '29 

            If FoundDeliveryRef Then 'DeliveryRef already in DB tblDeliveryInfo:

                oldDBDeliveryDate = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "DeliveryDate")
                oldDBDeliveryReference = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "DeliveryReference")

                If UCase(oldDBDeliveryReference) = UCase(strDeliveryRef) Then
                    'Delivery Date already exists - record already exists in MySQL database:
                    'BUT does it need updating ? with what fields from the spreadsheet ?
                    'Maybe something might have changed on that row of the spreadsheet to an earlier import from a previous spreadsheet 
                    ' - FOR THAT SAME DAY ?
                    '- maybe could complare each field to see if any changes have taken place between the two ???
                    UpdateNeeded = True
                    'OK so the date could be changed to the current date - but what if the SUPPLIER is different ???
                    'Replace DATE and Update SAME record ?
                    'Replace Supplier and Update SAME record (IF A NUMBER) ?
                    'INSERT AS A NEW RECORD with TODAYs date and SUPPLIER as per spreadsheet ?
                    Updates = Updates + 1
                Else
                    'The EMPLOYEE NAMES are different:
                    UpdateNeeded = False
                End If
                UpdateCriteria = "ID = " & ReturnValue
                'ASk user if they want to replace the existing name in the database with the new name.
                If ReplaceAll = False Then
                    If UpdateNeeded Then
                        'Answer = MsgBox("Already Exists in DB - Do you want to replace with NEW row from spreadsheet ?  " & strDeliveryDate & ", " & strDeliveryRef & " ?", vbYesNoCancel, "REPLACE EXISTING ?")
                        Answer = vbNo
                        If Answer = vbYes Then
                            'UPDATE database with new name and details etc
                            SavedOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, "tblDeliveryInfo", InboundFieldnames, DeliveryInfo_FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                            If Not SavedOK Then
                                ErrorList(ErrIDX) = "Delivery Reference in DB: " & oldDBDeliveryDate & ", " & strDeliveryRef & " ,Did NOT UPDATE"
                                ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                                ErrIDX = ErrIDX + 1
                            End If
                        ElseIf Answer = vbCancel Then
                            Exit For
                        Else 'NO
                            'leave as is.
                        End If
                    End If
                Else
                    'maybe need extra condition - if UpdateNeeded before using SAVEDOK. ???
                    SavedOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, "tblDeliveryInfo", InboundFieldnames, DeliveryInfo_FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                    If Not SavedOK Then
                        ErrorList(ErrIDX) = "Delivery Reference in DB: " & oldDBDeliveryDate & ": " & strDeliveryRef & " ,Did NOT UPDATE"
                        ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                        ErrIDX = ErrIDX + 1
                    End If
                End If
            Else
                SavedOK = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, "tblDeliveryInfo", InboundFieldnames, DeliveryInfo_FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)
                If Not SavedOK Then
                    ErrorList(ErrIDX) = "Delivery Reference in DB: " & strDeliveryDate & ": " & strDeliveryRef & " ,Did NOT INSERT"
                    ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                    ErrIDX = ErrIDX + 1
                End If
            End If
            'IF NOT SAVED OK - then store offending rows in the ERROR LIST.

            'Update Progress control - either a progress bar or lblProgress label:
            Percentage = (RowIDX / TotalRows) * 100
            frmMainGIForm.pbarMain.Value = CInt(Percentage)
            frmMainGIForm.txtMessages.Text = "Please Wait ... Importing Data: Row " & CStr(RowIDX) & "/" & CStr(TotalRows) & " : " & CInt(Percentage) & "%"
            Application.DoEvents()
        Next
        'MsgBox("Finished. TOTAL UPDATES: " & CStr(Updates))
        frmMainGIForm.txtMessages.Text = "Finished. TOTAL UPDATES: " & CStr(Updates)
    End Sub

    Sub ImportData2(ByRef ErrorList() As String, ByVal ImportDate As String,
                    Optional ByVal LastImportDate As String = "",
                    Optional ByVal Confirm As Boolean = False)
        'UPDATE EMPLOYEES:
        'Extract employee details from the CSV or XML or EXCEL files and insert into the MYSQL TIMESHEET SERVER:
        'Get EXCEL DATA from selected Workbook:
        Dim ExcelFilename As String
        Dim SheetName As String
        Dim StartCell As String
        Dim EndCell As String
        Dim StartRow As Long
        Dim StartCol As Long
        Dim SortCol1 As Long
        Dim SortCol2 As Long
        Dim CheckCol As Long 'First FULL column in range thats filled to calc total Rows.
        Dim CheckRow As Long 'First Full ROW in range thats filled to calc total Columns.
        Dim ExtractArr(,) As Object
        Dim DeliveryInfoFieldnames As String
        Dim DeliveryInfoFieldsArr() As String = Nothing
        Dim DailySheetFieldnames As String
        Dim DailySheetFieldsArr() As String = Nothing
        Dim DeliveryInfo_FieldValues As String
        Dim Daily_FieldValues As String
        Dim InboundFieldnames As String
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim ASN As String
        Dim strSupplierCode As String
        Dim strSupplierName As String
        Dim strDeliveryType As String
        Dim strPalletsDue As String
        Dim strCartonsDue As String
        Dim strOrigin As String
        Dim strShift As String
        Dim strDestination As String
        Dim strSpecialInstructions As String
        Dim strDueTime As String
        Dim strLines As String 'no 14
        Dim strCases As String 'no 15
        Dim strActualCases As String 'no. 16
        Dim strExpectedTotes As String
        Dim strExpectedCages As String
        Dim strExpectedPallets As String
        Dim strReadyLabel As String
        Dim strPalletise As String
        Dim strCollar As String
        Dim strWrapStrap As String
        Dim strBagging As String
        Dim strManHR As String
        Dim RowIDX As Long
        Dim ColIDX As Long
        Dim TotalRows As Long
        Dim TotalFields As Long
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim EmpNo As String = ""
        Dim description As String = ""
        Dim LastPWChange As String = ""
        Dim LastUpdated As String = ""
        Dim ErrMessage As String = ""
        Dim ExcludeFields As String = ""
        Dim FieldnameArr() As String = Nothing
        Dim Extract As String = ""
        Dim SpacePos As Integer = 0
        Dim FullName As String = ""
        Dim UpdateCriteria As String = ""
        Dim UpdateRecord As Boolean = False
        Dim UpdateNeeded As Boolean = False
        Dim DescriptionChanged As Boolean = False
        Dim SavedOK As Boolean
        Dim ReplaceName As Boolean = False
        Dim Answer As Integer
        Dim AnswerDuplicate As Integer
        Dim FoundDeliveryDate As Boolean = False
        Dim FoundDeliveryRef As Boolean = False
        Dim SearchField As String
        Dim SearchText As String = ""
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllReturnedValues As Object()
        Dim oldDBDeliveryDate As String
        Dim oldDBDeliveryReference As String
        Dim oldDBSupplierCode As String
        Dim oldDBSupplierName As String
        Dim ErrIDX As Long
        Dim Empno_7Digits As String = ""
        Dim dtDateNow As DateTime
        Dim ReplaceAll As Boolean = False
        Dim TotalUpdatesNeeded As Long = 0
        Dim TotalDailyChangesNeeded As Long = 0
        Dim TotalDeliveryInfoChangesNeeded As Long = 0
        Dim TotalDescNeeded As Long
        Dim dtDateTime As DateTime
        Dim Percentage As Single = 0.0F
        Dim strASNNo As String
        Dim strExpectedCases As String
        Dim strExpectedLines As String
        Dim strEstimatedPallets As String
        Dim strEstimatedCages As String
        Dim strEstimatedTotes As String
        Dim SortFieldIDX As Long 'Position within field list of field we want to sort by
        Dim strDeliveryComments As String
        Dim TotalUpdatesFound As Long = 0
        Dim Updates As Long = 0
        Dim TempFilename As String
        Dim dtDueTime As DateTime
        Dim dblDueTime As Double
        Dim ErrSaveMessage As String

        ReDim ErrorList(1)
        'Get user to select location of DAILY spreadsheet within KPI INBOUND workbook:
        'Needs sorting in reverse Delivery Date order to catch the LATEST Delivery record REF as any Delivery not unloaded before and turned away will have exactly the same 
        '      details again but on a later date.
        ExcelFilename = BrowseFilename(1)
        If Len(ExcelFilename) < 2 Then
            Exit Sub
        End If
        TempFilename = Application.StartupPath & "\TempFilename.xls"
        'PLEASE WAIT ...
        'Get the row on which the required date starts:

        SheetName = "Daily"
        StartRow = 2
        StartCol = 1
        SortCol1 = 6
        SortCol2 = 7
        CheckRow = 2
        CheckCol = 7
        frmMainGIForm.txtMessages.Text = "Please Wait ... Importing Data"
        ReDim ExtractArr(1, 1)
        If Len(ImportDate) > 0 Then
            SearchText = ImportDate
        Else
            SearchText = Now().ToString("yyyy-MM-dd")
        End If
        'what if a DATE RANGE has been selected ?
        'Sheet is sorted form latest date to earliest.
        'Call DanG_EXCEL_Module.SortSheet(ExcelFilename, SheetName, StartRow, StartCol, SortCol1, ExtractArr, CheckRow, CheckCol, SortCol2, True, True, TempFilename)
        ExtractArr = ExtractExcelRange(ExcelFilename, SheetName, SearchText, 6, "DATE", TempFilename, True, 1, 2, 0, 0, Nothing, LastImportDate)
        'MsgBox("OK - First row extracted: " & ExtractArr(1, 1))
        If ExtractArr Is Nothing Then
            MsgBox("No records Returned - Date Not Found")
            Exit Sub
        End If
        frmMainGIForm.txtMessages.Text = "Please Wait ... Data Extracted"
        TotalRows = UBound(ExtractArr, 1)
        ExcludeFields = ""
        DailySheetFieldnames = GetMyFields("tblDailySheet", frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        DailySheetFieldsArr = Split(DailySheetFieldnames, ",")
        DeliveryInfoFieldnames = GetMyFields("tblDeliveryInfo", frmMainGIForm.myConnString, ErrMessage, ExcludeFields)
        DeliveryInfoFieldsArr = Split(DeliveryInfoFieldnames, ",")
        InboundFieldnames = DeliveryInfoFieldsArr(1)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(2)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(32)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(3)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(4)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(5)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(6)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(7)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(8)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(9)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(10)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(11)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(12)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(13)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(14)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(15)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(16)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(17)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(18)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(22)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(25)

        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(26)
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(27) 'SHIFT = 28 for value position.
        InboundFieldnames = InboundFieldnames & "," & DeliveryInfoFieldsArr(28)
        TotalFields = UBound(ExtractArr, 2)
        SortFieldIDX = 6 'DeliveryDate as ID=0 
        'Call Sort2DArray(False, ExtractArr, SortFieldIDX)
        ErrIDX = 0
        'TotalDailyChangesNeeded  = CheckEmployeeUpdates(ExcelFilename, SheetName, TotalDescNeeded)
        'TotalUpdatesFound = Check_ImportDataUpdates(ExtractArr)
        ReplaceAll = True
        If TotalUpdatesFound > 0 Then
            If Confirm Then
                Answer = MsgBox("FOUND " & CStr(TotalUpdatesFound) & " Updates Found. SELECT YES to be prompted for EACH update - or NO to REPLACE ALL", vbYesNoCancel, "Total Updates Required: " & CStr(TotalUpdatesFound) & " Prompt for EACH Update ?")
                If Answer = vbYes Then
                    ReplaceAll = False
                ElseIf Answer = vbNo Then
                    ReplaceAll = True
                Else
                    Exit Sub
                End If
            Else
                ReplaceAll = True
            End If
        End If
        Percentage = 0
        Updates = 0
        frmMainGIForm.txtMessages.Text = "Please Wait ... Updating database"
        For RowIDX = 1 To TotalRows
            DeliveryInfo_FieldValues = ""
            Daily_FieldValues = ""
            strDeliveryComments = ""
            AllReturnedValues = Nothing
            DescriptionChanged = False
            UpdateNeeded = False
            Extract = ExtractArr(RowIDX, 7) 'Delivery Ref
            If IsNothing(Extract) Then
                Continue For
            End If
            strDeliveryRef = Extract
            Extract = ExtractArr(RowIDX, 6) 'Delivery Date
            If IsNothing(Extract) Then
                strDeliveryDate = "1970-01-01"
            Else
                If Len(Extract) > 0 Then
                    If IsDate(Extract) Then
                        dtDateTime = CDate(Extract)
                        'Could combine with DUE TIME ?????
                        strDeliveryDate = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    Else
                        strDeliveryDate = "1970-01-01"
                    End If
                Else
                    strDeliveryDate = "1970-01-01"
                End If
            End If
            strDeliveryType = ExtractArr(RowIDX, 1)
            If UCase(strDeliveryType) = "PRE-LABELLED" Then
                strReadyLabel = "YES"
            Else
                strReadyLabel = "NO"
            End If
            strWrapStrap = "NO"
            strPalletise = "NO"
            strCollar = "NO"
            strBagging = "NO"
            strSupplierCode = ExtractArr(RowIDX, 2)
            strSupplierName = ExtractArr(RowIDX, 3)
            strPalletsDue = ExtractArr(RowIDX, 4)
            strCartonsDue = ExtractArr(RowIDX, 5)

            strOrigin = ExtractArr(RowIDX, 8)
            strSpecialInstructions = ExtractArr(RowIDX, 9)
            strASNNo = ExtractArr(RowIDX, 10)
            strShift = ExtractArr(RowIDX, 11)
            strDestination = ExtractArr(RowIDX, 12)
            strDueTime = ExtractArr(RowIDX, 13)
            If IsNumeric(strDueTime) Then
                dblDueTime = CDbl(strDueTime)
                strDueTime = DateTime.FromOADate(ExtractArr(RowIDX, 13)).ToString("HH:mm:ss")
                'DueTime = dblDueTime.ToString("HH:mm:ss")
            ElseIf IsDate(strDueTime) Then
                dtDueTime = CDate(strDueTime)
                strDueTime = dtDueTime.ToString("HH:mm:ss")
            Else
                'leave as STRING
                'leave as is - usually a message like UNBOOKED entered - even though supposed to be a TIME column ???
            End If
            strExpectedLines = ExtractArr(RowIDX, 14)
            strExpectedCases = ExtractArr(RowIDX, 15)
            strActualCases = ExtractArr(RowIDX, 16)
            strEstimatedTotes = ExtractArr(RowIDX, 17)
            strEstimatedCages = ExtractArr(RowIDX, 18)
            strEstimatedPallets = ExtractArr(RowIDX, 19)
            strManHR = ExtractArr(RowIDX, 20)

            SearchField = "DeliveryReference"
            SearchValue = strDeliveryRef
            ReturnField = "ID"
            ReturnValue = ""
            'ReDim AllReturnedValues(1)
            'Search for Delivery Date and Delivery Reference together here - use criteria:
            FoundDeliveryRef = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblDeliveryInfo", SearchField, SearchValue, "STRING", ReturnField, ReturnValue, AllReturnedValues)

            Empno_7Digits = EmpNo
            dtDateNow = Now()
            LastUpdated = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
            'LastUpdated = FormatDateTime(dtDateNow, "yyyy-MM-dd HH:mm:ss") 'Not Working
            'FOR tblDeliveryInfo first:


            If FoundDeliveryRef Then 'DeliveryRef already in DB tblDeliveryInfo:

                oldDBDeliveryDate = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "DeliveryDate")
                oldDBDeliveryReference = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "DeliveryReference")
                oldDBSupplierCode = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "Supplier_Code")
                oldDBSupplierName = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, DeliveryInfoFieldsArr, "Supplier_Name")

                If UCase(oldDBDeliveryReference) = UCase(strDeliveryRef) Then
                    'Delivery Date already exists - record already exists in MySQL database:
                    'BUT does it need updating ? with what fields from the spreadsheet ?
                    'Maybe something might have changed on that row of the spreadsheet to an earlier import from a previous spreadsheet 
                    ' - FOR THAT SAME DAY ?
                    '- maybe could complare each field to see if any changes have taken place between the two ???
                    If UCase(oldDBSupplierName) = UCase(strSupplierName) Then
                        'OK Suppliers match.
                    Else
                        Answer = MsgBox("Duplicate Reference will be created for today. Select YES to use NEW Supplier or NO to keep existing", vbYesNo, "WARNING: Supplier Names DO NOT MATCH")
                        If Answer = vbNo Then
                            strSupplierName = oldDBSupplierName
                        End If
                    End If
                    AnswerDuplicate = MsgBox("Reference Already Exists. Answer YES to keep old record and create a NEW duplicated reference for today OR Answer No to OVERWRITE the old reference with TODAYs date ?", vbYesNo, "REFERENCE EXISTS")
                    If AnswerDuplicate = vbYes Then
                        UpdateNeeded = False
                        strDeliveryDate = Now().ToString("yyyy-MM-dd HH:mm:ss") 'currently from spreadsheet.
                    Else
                        strDeliveryDate = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
                        UpdateNeeded = True
                    End If

                    'OK so the date could be changed to the current date - but what if the SUPPLIER is different ???
                    'Replace DATE and Update SAME record ?
                    'Replace Supplier and Update SAME record (IF A NUMBER) ?
                    'INSERT AS A NEW RECORD with TODAYs date and SUPPLIER as per spreadsheet ?
                    Updates = Updates + 1
                Else
                    'The EMPLOYEE NAMES are different:
                    UpdateNeeded = False
                End If
                UpdateCriteria = "ID = " & ReturnValue
                'ASk user if they want to replace the existing name in the database with the new name.
            Else
                UpdateNeeded = False
                strDeliveryDate = dtDateNow.ToString("yyyy-MM-dd HH:mm:ss")
            End If
            'IF NOT SAVED OK - then store offending rows in the ERROR LIST.

            DeliveryInfo_FieldValues = strDeliveryDate '1
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strDeliveryRef  '2
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strSupplierCode  '32
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strSupplierName  '4
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strASNNo '5
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strExpectedCases  '6
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strExpectedLines  '7
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strEstimatedPallets  '8
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strEstimatedCages  '9
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strEstimatedTotes  '10
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strCartonsDue  '11
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strPalletsDue  '12
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strReadyLabel  '13

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strPalletise  '14
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strWrapStrap  '15
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strCollar  '16
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strBagging  '17

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strActualCases  '18
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strManHR  '19
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & LastUpdated

            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strOrigin  '26
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strDueTime '27
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strShift  '28
            DeliveryInfo_FieldValues = DeliveryInfo_FieldValues & "," & strDeliveryComments  '29 

            'SavedOK = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, "tblDeliveryInfo", InboundFieldnames, DeliveryInfo_FieldValues, ErrMessage, UpdateCriteria, ExcludeFields)

            SavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, UpdateNeeded, "", "tblDeliveryInfo",
                                                           InboundFieldnames, DeliveryInfo_FieldValues, UpdateCriteria,
                                                           ExcludeFields, ErrSaveMessage)
            If Not SavedOK Then
                ErrorList(ErrIDX) = "Delivery Reference in DB: " & strDeliveryDate & ": " & strDeliveryRef & " ,Did NOT INSERT"
                ReDim Preserve ErrorList(UBound(ErrorList) + 1)
                ErrIDX = ErrIDX + 1
            End If
            'Update Progress control - either a progress bar or lblProgress label:
            Percentage = (RowIDX / TotalRows) * 100
            frmMainGIForm.pbarMain.Value = CInt(Percentage)
            frmMainGIForm.txtMessages.Text = "Please Wait ... Importing Data: Row " & CStr(RowIDX) & "/" & CStr(TotalRows) & " : " & CInt(Percentage) & "%"
            Application.DoEvents()
        Next
        'MsgBox("Finished. TOTAL UPDATES: " & CStr(Updates))
        frmMainGIForm.txtMessages.Text = "Finished. TOTAL UPDATES: " & CStr(Updates)
    End Sub 'IMPORTDATA2

    Function Check_ImportDate_IsToday(CheckTable As String, CheckDateField As String, ConnString As String, ByRef ReturnColour As String, ByRef ErrMessage As String) As Boolean
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim FoundTodayDate As Boolean
        Dim FoundYesterdayDate As Boolean
        Dim strSQL As String
        Dim ExtractArr(,) As Object
        Dim LastDate As String
        Dim dtLastDate As DateTime
        Dim DaysOut As Integer
        Dim MessageTitle As String
        Dim FriendlyMessage As String

        Check_ImportDate_IsToday = False


        'check last date in tblDeliveryInfo to see if it matches TODAY. - TURN GREEN.
        'if within this week - turn AMBER.
        'if more than 1 week - turn RED.
        SearchField = "DeliveryDate"
        SearchValue = Now().ToString("yyyy-MM-dd")
        ReturnField = "ID"
        ReturnValue = ""
        ReturnColour = ""
        If UCase(CheckTable) = UCase("tblEmployees") Then
            MessageTitle = "Update Employees"
            FriendlyMessage = "Employees Table"
        ElseIf UCase(CheckTable) = UCase("tblDeliveryInfo") Then
            MessageTitle = "Import Data"
            FriendlyMessage = "Delivery Info Table"
        Else
            MessageTitle = "TIMESHEET"
            FriendlyMessage = "Unknown"
        End If
        If Len(CheckTable) = 0 Then
            ErrMessage = "Error: CheckTable passed is empty"
            Exit Function
        End If
        If Len(ConnString) = 0 Then
            ErrMessage = "Error: No Connection String Passed"
            Exit Function
        End If

        strSQL = "SELECT MAX(" & CheckDateField & ") FROM " & CheckTable
        ExtractArr = Module_DanG_MySQL_Tools.MySQLToArray(ConnString, strSQL, ErrMessage)
        'ReDim AllReturnedValues(1)
        LastDate = ""
        If Not IsNothing(ExtractArr) Then
            LastDate = ExtractArr(0, 0)
            If IsDate(LastDate) Then
                dtLastDate = CDate(LastDate)
                DaysOut = Now().Subtract(LastDate).Days
                If DaysOut = 0 Then
                    ReturnColour = "GREEN"
                ElseIf DaysOut < 3 Then
                    ReturnColour = "AMBER"
                Else
                    ReturnColour = "RED"
                End If
                'MsgBox("DAYS OUTSTANDING: " & CStr(DaysOut), vbOK, MessageTitle)
            Else
                'MsgBox("Cannot Find Date - " & FriendlyMessage & " Empty")
                ReturnColour = "RED"
            End If
            FoundTodayDate = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, "tblDeliveryInfo", SearchField, SearchValue, "STRING", ReturnField, ReturnValue)
        End If
        Check_ImportDate_IsToday = FoundTodayDate
    End Function

    Function Get_TotalHours(ScrollControlFrame As ScrollableControl, strDeliveryDate As String, strDeliveryRef As String,
                            Optional ByRef TimeSheetHours() As String = Nothing,
                            Optional DelimChar As String = "|",
                            Optional FormName As String = "GI_TIMESHEET_") As Double
        Dim ctrl As Control
        Dim ControlName As String
        Dim OpIDX As Long
        Dim TotalOps As Long
        Dim strStartTime As String = ""
        Dim strEndTime As String = ""
        Dim FLMdtStartTime As DateTime
        Dim FLMdtEndTime As DateTime
        Dim FLMdblStartTime As Double
        Dim FLMdblEndTime As Double
        Dim FLMdblTotalTime As Double
        Dim SPAN As TimeSpan
        Dim dblStartTime As Double = 0.0F
        Dim dblEndTime As Double = 0.0F
        Dim dblTotalTime As Double = 0.0F
        Dim dtStartTime As DateTime
        Dim dtEndTime As DateTime
        Dim VarKeyTotals As String = ""
        Dim VarKeyControls As String = ""
        Dim dblTotalHours As Double
        Dim tempProperty As clsControls
        Dim TAGNum As String
        Dim IsError As Boolean = False
        Dim Totals As clsTotals
        Dim dblTotalMinutes As Double
        Dim dblTotalSeconds As Double
        Dim dblSeconds As Double
        Dim intHours As Integer
        Dim intMins As Integer
        Dim intSecs As Integer
        Dim strHours As String
        Dim strMins As String
        Dim strSecs As String
        Dim intDays2 As Integer
        Dim intHours2 As Integer
        Dim intMins2 As Integer
        Dim intSecs2 As Integer
        Dim intTotalHours As Integer
        Dim intTotalMins As Integer
        Dim intTotalSecs As Integer
        Dim CalcFLM As Boolean = False
        Dim strTotalHours As String
        Dim strTotalMins As String
        Dim strTotalSecs As String
        Dim ErrMessage As String = ""
        Dim intDays As Integer
        Dim TotalSPAN As TimeSpan
        Dim strTime As String
        Dim strOpName As String
        Dim strOpActivity As String
        Dim strFLMName As String
        Dim strRow As String
        Dim strComments As String
        Dim FinalTime As String
        Dim strDays As String
        Dim TotalSpanHours As TimeSpan
        Dim PointPos As Integer

        '==============================================================================

        'Span = dtDateLoggedOut.Subtract(dtDateLoggedIn)
        'dtMinutesDiff = DateDiff(DateInterval.Minute, dtDateLoggedIn, dtDateLoggedOut)
        'DecHours = Math.Abs(dtMinutesDiff) / 60
        'DecMinutes = Math.Abs(dtMinutesDiff) Mod 60
        'If DecHours < 1 Then
        'strHours = "0"
        'Else
        'Dim theRounded = Math.Sign(DecHours) * Math.Floor(Math.Abs(DecHours) * 100) / 100.0
        'intHours = CInt(theRounded)
        'strHours = CStr(intHours)
        'hmmm hours still ending up as 1.34 and 5.6 etc.
        'End If
        'strMins = CStr(DecMinutes)

        'strLoggedInDuration = Span.Hours & "h " & Span.Minutes & "m " & Span.Seconds & "s "

        '==============================================================================


        'what if either time is 1970-01-01 or undefined ?
        FLMdblTotalTime = 0.0F
        Get_TotalHours = 0.0F
        dblTotalHours = 0.0F
        dblTotalSeconds = 0.0F
        TotalOps = Get_TotalRows("tblOperatives", strDeliveryRef)

        VarKeyTotals = CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef
        Totals = dic_Totals(VarKeyTotals)
        If Not IsNothing(Totals) Then
            Totals.TotalOpHours = 0
            Totals.TotalFLMHours = 0
        End If
        ReDim Preserve TimeSheetHours(TotalOps + 2)
        'TotalOps = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1
        OpIDX = 1 'Pick first row
        IsError = False
        VarKeyControls = CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & "43"
        tempProperty = dic_Controls(VarKeyControls)
        If tempProperty IsNot Nothing Then
            tempProperty.ControlOPstrTotalHrs = ""
            tempProperty.ControlOPdblTotalHrs = 0.0F
        End If
        For Each ctrl In ScrollControlFrame.Controls
            ControlName = ctrl.Name
            TAGNum = ctrl.Tag
            VarKeyControls = CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & TAGNum
            tempProperty = dic_Controls(VarKeyControls)
            OpIDX = Get_NumericPartOfString(ControlName)
            'NEED to save the times to the DB HERE. - comments.

            'strStartTime = tempProperty.ControlDeliveryRef
            If IsNumeric(TAGNum) Then
                If CLng(TAGNum) = 441 Then 'wont work - wrong FRAME !!!
                    If Not IsNothing(tempProperty) Then
                        tempProperty.ControlFLMdblTotalHrs = 0.0F
                        tempProperty.ControlFLMstrTotalHrs = ""
                        FLMdtStartTime = tempProperty.ControlFLMStartDateTime
                        'tempProperty.ControlFLMdblTotalHrs = FLMdtStartTime.ToOADate()
                        CalcFLM = True
                        tempProperty = dic_Controls(CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & CStr(CLng(TAGNum) - 401))
                        strFLMName = tempProperty.ControlValue
                        tempProperty = dic_Controls(VarKeyControls)
                    End If
                End If
                If CLng(TAGNum) = 442 Then
                    If Not IsNothing(tempProperty) Then
                        FLMdtEndTime = tempProperty.ControlFLMEndDateTime
                        If FLMdtEndTime > FLMdtStartTime Then
                            Span = CalculateHours(FLMdtStartTime, FLMdtEndTime, strTotalHours, strTotalMins, False, False, strTotalSecs)
                            intDays = Span.Days
                            intHours = Span.Hours
                            intMins = Span.Minutes
                            intSecs = Span.Seconds
                            dblTotalSeconds = Span.TotalSeconds
                            tempProperty.ControlFLMdblTotalHrs = Span.TotalHours
                            tempProperty.ControlFLMstrTotalHrs = strTotalHours & ":" & strTotalMins & ":" & strTotalSecs

                            Totals.TotalFLMHours = Span.TotalHours
                            If intDays = 0 Then
                                TimeSheetHours(0) = strFLMName & DelimChar & CStr(intHours) & "H " & CStr(intMins) & "M " & CStr(intSecs) & "S"
                            Else
                                TimeSheetHours(0) = strFLMName & DelimChar & CStr(intDays) & "D " & CStr(intHours) & "H " & CStr(intMins) & "M " & CStr(intSecs) & "S"
                            End If
                        End If
                        CalcFLM = True
                    End If
                    If UCase(ControlName) = UCase("txtOperativeStartTime:") & CStr(OpIDX) Then
                    'do we use the date from the control collection or trust the single time string here ?
                    dtStartTime = tempProperty.ControlOpStartDateTime
                    strStartTime = dtStartTime.ToString("HH:mm:ss")
                    tempProperty = dic_Controls(CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & CStr(CLng(TAGNum) - 402))
                    strOpName = tempProperty.ControlValue
                    tempProperty = dic_Controls(CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & CStr(CLng(TAGNum) - 401))
                    strOpActivity = tempProperty.ControlValue
                    tempProperty = dic_Controls(VarKeyControls)
                End If
                If UCase(ControlName) = UCase("txtOperativeFinishTime:") & CStr(OpIDX) Then
                    dtEndTime = tempProperty.ControlOpEndDateTime
                    strEndTime = dtEndTime.ToString("HH:mm:ss")
                    If dtEndTime >= dtStartTime Then
                        Span = CalculateHours(dtStartTime, dtEndTime, strTotalHours, strTotalMins, False, False, strTotalSecs)
                        dblSeconds = Span.TotalSeconds
                        dblTotalSeconds = dblTotalSeconds + Span.TotalSeconds
                        dblTotalHours = Span.TotalHours
                        strTotalHours = dblTotalHours.ToString
                        'intTotalHours = Math.Floor(Span.TotalHours)
                        'intTotalMins = Math.Floor(Span.TotalMinutes) 'may need ceiling to add 1 ???
                        'intTotalSecs = Math.Floor(Span.TotalSeconds)
                        intTotalHours = CInt(strTotalHours)
                        intTotalMins = CInt(strTotalMins)
                        intTotalSecs = CInt(strTotalSecs)
                        intDays = Math.Floor(Span.TotalDays)

                        'INSERT THE TIME INTO THE COMMENTS TEXTBOX HERE:


                        'strDays = CStr(intDays)
                        'strHours = CStr(intTotalHours)
                        'strMins = CStr(intTotalMins)
                        'strSecs = CStr(intTotalSecs)
                        strDays = CStr(intDays)
                        'strHours = strTotalHours 'IF DECIMAL NEEDED.
                        strHours = CStr(intTotalHours)
                        strMins = strTotalMins
                        strSecs = strTotalSecs
                        If intTotalHours < 10 Then
                            strHours = "0" & strHours
                        End If
                        If intTotalMins < 10 Then
                            strMins = "0" & strMins
                        End If
                        If intTotalSecs < 10 Then
                            strSecs = "0" & strSecs
                        End If
                        If intDays < 10 Then
                            strDays = "0" & strDays
                        End If
                        If intDays > 0 Then
                            strDays = strDays & "d:"
                        Else
                            strDays = ""
                        End If
                        strTime = strHours & ":" & strMins & ":" & strSecs
                        strRow = CStr(OpIDX) & DelimChar & strOpName & DelimChar & strOpActivity & DelimChar & strTime
                        TimeSheetHours(OpIDX) = strRow
                        'Insert into OpComments:

                        tempProperty = dic_Controls(VarKeyControls)
                        tempProperty.ControlOPstrTotalHrs = strTime
                        tempProperty.ControlOpdblTotalSecs = Span.TotalSeconds
                        tempProperty.ControlOPdblTotalHrs = Span.TotalHours
                        strComments = "txtOpComments:" & CStr(OpIDX)
                        Call frmMainGIForm.InsertValueIntoForm(FormName & FormID, strComments, strTime)
                        'dic_Controls(CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & TAGNum) = tempProperty
                        'tempProperty = dic_Controls(CDate(strDeliveryDate).ToString("dd/MM/yyyy") & "_" & strDeliveryRef & "_" & CStr(CLng(TAGNum) - 399))
                        'tempProperty.ControlValue = strTime
                    End If
                End If
                    dic_Controls(VarKeyControls) = tempProperty
                End If
            End If
        Next
        FLMdblStartTime = 0.0F
        FLMdblEndTime = 0.0F
        dblStartTime = 0.0F
        dblEndTime = 0.0F
        FLMdblStartTime = FLMdtStartTime.ToOADate()
        FLMdblEndTime = FLMdtEndTime.ToOADate()
        dblStartTime = dtStartTime.ToOADate()
        dblEndTime = dtEndTime.ToOADate()
        intHours = 0
        intMins = 0
        'Need to test if greater than dblStartTime for default hrs for 1970-01-01
        'Need to save to TotalHrs in tblLabourHours:
        If CalcFLM Then
            Span = CalculateHours(FLMdtStartTime, FLMdtEndTime, strTotalHours, strTotalMins, True, False, strTotalSecs)
        Else
            Span = CalculateHours(dtStartTime, dtEndTime, strTotalHours, strTotalMins, True, False, strTotalSecs)
        End If
        If Len(ErrMessage) > 0 Then
            MsgBox(ErrMessage)
            IsError = True
        Else

            'intHours = CInt(dblTotalSeconds / 3600)
            'intMins = CInt(System.Math.IEEERemainder((dblTotalSeconds / 60), 60))
            'intSecs = CInt(System.Math.IEEERemainder(dblTotalSeconds, 60))
            TotalSPAN = TimeSpan.FromSeconds(dblTotalSeconds)
            dblTotalHours = TotalSPAN.TotalHours
            strTotalHours = CStr(dblTotalHours)
            PointPos = InStr(strTotalHours, ".")
            If PointPos > 0 Then
                strTotalHours = Mid(strTotalHours, 1, PointPos - 1)
            Else
                strTotalHours = "0"
            End If
            TotalSpanHours = TimeSpan.FromHours(dblTotalHours)
            'intDays = CInt((dblTotalSeconds / 3600) / 24)

            dblTotalTime = Span.TotalHours
            intDays = TotalSPAN.Days
            intHours = CInt(strTotalHours)
            'intHours = TotalSPAN.Hours
            intMins = TotalSPAN.Minutes
            intSecs = TotalSPAN.Seconds
            'intMins = Span.Minutes
        End If

        If Not IsNothing(Totals) Then
            If CalcFLM Then
                Totals.TotalFLMHours = Span.TotalHours
            End If
            Totals.TotalOpHours = Totals.TotalOpHours + Span.TotalHours
            strSecs = CStr(intSecs)
            strMins = CStr(intMins)
            strHours = CStr(intHours)
            strDays = CStr(intDays)
            If intHours < 10 Then
                strHours = "0" & strHours
            End If
            If intMins < 10 Then
                strMins = "0" & strMins
            End If
            If intSecs < 10 Then
                strSecs = "0" & strSecs
            End If
            If intDays < 10 Then
                strDays = "0" & strDays
            End If
            strDays = strDays & "d"
            If intDays > 0 Then
                'FinalTime = strDays & " " & strHours & "H " & strMins & "m " & strSecs & "s"
                FinalTime = strHours & "H " & strMins & "m " & strSecs & "s"
            Else
                FinalTime = strHours & "h " & strMins & "m " & strSecs & "s"

            End If

            TimeSheetHours(OpIDX + 1) = FinalTime
            Call frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalHours", FinalTime)

            'SaveOK(4) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate(RowIDX - 1), "", DBTAbles(4), Fieldnames(4), FieldValues(4),
            'UpdateCriteria(4), ExcludeFields(4), ErrMessages(4), False, ",")


            If CalcFLM Then
                TimeSheetHours(0) = FinalTime
            End If
            'NEED TO SAVE TO PROPERTY ControlOpdblTotalSecs
            'Then Add a SAVE code in savecontrols to use ControlOpDblTotalSecs - and calculate total time from there as a string.
            dic_Totals(VarKeyTotals) = Totals
            Get_TotalHours = dblTotalHours
        Else
            Get_TotalHours = 0.0F
        End If


    End Function

    Function CalculateHours(dtStartDateTime As DateTime, dtEndDateTime As DateTime, ByRef strTotalHours As String, ByRef strTotalMinutes As String,
                            Optional RoundUpSeconds As Boolean = True,
                            Optional RoundUpToNearest15Mins As Boolean = False,
                            Optional ByRef strTotalSeconds As String = "",
                            Optional PreviousHours As Integer = 0,
                            Optional PreviousMins As Integer = 0,
                            Optional PreviousSecs As Integer = 0,
                            Optional ByRef NewDateTime As DateTime = #1970-01-01#,
                            Optional ErrMessage As String = "") As TimeSpan
        Dim dblStartTime As Double
        Dim dblEndTime As Double
        Dim dblTotalTime As Double
        Dim dblSeconds As Double
        Dim intHours As Integer
        Dim intMins As Integer
        Dim intSecs As Integer
        Dim strSeconds As String
        Dim strMins As String
        Dim strHours As String
        Dim dblExtraHoursResult As Double
        Dim intExtraHoursResult As Integer
        Dim Span As TimeSpan = Nothing
        Dim IsError As Boolean = False
        Dim ResultMins As Integer
        Dim PreviousDateTime As DateTime

        CalculateHours = Nothing

        Span = dtEndDateTime.Subtract(dtStartDateTime)
        PreviousDateTime = NewDateTime
        NewDateTime = PreviousDateTime.AddHours(PreviousHours)
        NewDateTime = NewDateTime.AddMinutes(PreviousMins)
        NewDateTime = NewDateTime.AddSeconds(PreviousSecs)
        dblStartTime = 0.0F
        dblEndTime = 0.0F
        dblStartTime = dtStartDateTime.ToOADate()
        dblEndTime = dtEndDateTime.ToOADate()
        strTotalHours = "0"
        strTotalMinutes = "0"
        strTotalSeconds = "0"
        'If dblStartTime > 0 And dblEndTime > 0 Then
        'If dtEndDateTime > dtStartDateTime Then
        Span = dtEndDateTime.Subtract(dtStartDateTime)
                dblTotalTime = Span.TotalHours
                dblSeconds = Span.TotalSeconds
                intHours = Span.Hours
                intMins = Span.Minutes
                intSecs = Span.Seconds
                If RoundUpSeconds Then
                    If intSecs > 29 Then
                        intMins = intMins + 1
                        intSecs = 0
                    End If
                End If
        If RoundUpToNearest15Mins Then
            ResultMins = intMins Mod 15
            'If intMins > 45 Then
            'intMins = 0
            'intHours = intHours + 1
            'Else
            '
            'End If
            If ResultMins > 7 Then
                intMins = intMins + ResultMins
            Else
                intMins = intMins - ResultMins
            End If
        End If
        strTotalHours = CStr(Span.Hours)
        strTotalMinutes = CStr(Span.Minutes)
        strTotalSeconds = CStr(Span.Seconds)
        strHours = intHours.ToString
        strMins = intMins.ToString
        strSeconds = intSecs.ToString
        'Else
        'MsgBox("Error: Start Time is Greater Than the Finish Time")
        IsError = True
        'End If
        'End If

        CalculateHours = Span

    End Function

    Function Get_DropdownItems(ConnString As String, strSQL As String, Criteria As String, Messages As String,
                               Optional ByRef SingleFieldArr() As String = Nothing,
                               Optional ByVal SingleFieldINdex As Integer = -1) As Object
        Dim ItemArr(,) As String = {{0}, {"NONE"}}
        Dim IDX As Long
        Dim TotalRows As Long

        Get_DropdownItems = Nothing
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
        ItemArr = MySQLToArray(ConnString, strSQL, Messages)
        If ItemArr IsNot Nothing Then
            If UBound(ItemArr) > 0 Then
                Get_DropdownItems = ItemArr
            End If
        End If
        If SingleFieldINdex > -1 Then
            ReDim SingleFieldArr(UBound(ItemArr, 2) - 1)
            IDX = 0
            TotalRows = UBound(ItemArr, 2)
            Do While IDX < TotalRows
                SingleFieldArr(IDX) = ItemArr(SingleFieldINdex, IDX)
                IDX = IDX + 1
            Loop

        End If

    End Function

    Function Get_StartTAG(TimeTAG As Long, RowNum As Long, TotalControlsPerRow As Long, ConstantStartOffset As Long) As Long
        Get_StartTAG = 0
        Get_StartTAG = TimeTAG + ((RowNum - 1) * TotalControlsPerRow) + ConstantStartOffset

    End Function

    Sub AddNewOperatives(ByRef TotalOps As Long, ByRef OpID As Long, FrameScrollControl As ScrollableControl, ByRef TagID As Long, ByRef btnTAGID As Long,
                         ByRef StartTABIndex As Long,
                         strDeliveryDate As String, strDeliveryRef As String, ASN As String, LowerTAG As Long, UpperTAG As Long, ByVal TimeTAGStart As Long,
                         Fieldnames As String, TotalRows As Long, Optional StartFieldIndex As Long = 5, Optional DontIncrementOP As Boolean = False,
                         Optional ByRef ErrMessage As String = "", Optional ByRef NewIndex As Long = 0, Optional OpName As String = "", Optional OpActivity As String = "",
                         Optional dtStartTime As Date = #1970-01-01#, Optional dtFinishTime As Date = #1970-01-01#,
                         Optional OpComment As String = "", Optional TotalOpTEXTControlsPerRow As Long = 6,
                         Optional DisableStartTimeButton As Boolean = False, Optional DisableFinishTimeButton As Boolean = False,
                         Optional DisableStartTimeEntry As Boolean = True, Optional DisableFinishTimeEntry As Boolean = True)
        Dim RowGap As Long
        Dim TopPos As Integer
        Dim StartTopPos As Integer
        Dim ScrollBarHeight As Integer
        Dim ComboArray() As String = Nothing
        Dim ListArray(,) As String = {{0}, {"NONE"}}
        Dim ComboIDX As Long
        Dim TimeTAGID As Long
        Dim ControlText As String
        Dim ControlType As String
        Dim ControlTAG As String
        Dim ControlDate As Date
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlDeliveryDate As Date
        Dim ControlDeliveryRef As String
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim Dic_Collection As Object
        Dim ControlRowNumber As Long
        Dim ControlTotalRows As Long
        Dim MakeVisible As Boolean
        Dim ControlBackColor As Integer
        Dim ControlForeColor As Integer
        Dim ControlLeftMargin As Boolean
        Dim ControlFieldname As String
        Dim FieldnameArr() As String
        Dim LoadFieldsOK As Boolean
        Dim ControlFieldsTable As String
        Dim SearchCriteria As String
        Dim ControlDBTable As String
        Dim ControlFontName As String
        Dim ControlFontSize As Integer
        Dim ControlStyle As String
        Dim FrameRowNumberField As String = ""
        Dim IDX As Long
        Dim RowIDX As Long
        Dim Firstname As String
        Dim Lastname As String
        Dim FullName As String
        Dim strActivity As String
        Dim ControlTABIndex As Long

        ReDim FieldnameArr(1)
        If Len(Fieldnames) > 0 Then
            'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        Else
            ReDim FieldnameArr(20)
            Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
            Beep()
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        End If
        FrameRowNumberField = "FrameRowNumber"
        ControlTABIndex = StartTABIndex
        StartTopPos = 2
        RowGap = 28
        If OpID < 2 Then
            TopPos = StartTopPos
        Else
            TopPos = ((OpID - 1) * RowGap) + StartTopPos
        End If

        TagID = Get_StartTAG(43, OpID, TotalOpTEXTControlsPerRow, 0)

        ControlDBTable = "tblOperatives"
        ControlFontName = "Cambria"


        If IsDate(strDeliveryDate) Then
            ControlDeliveryDate = CDate(strDeliveryDate)
        Else
            'MsgBox ("Need to pass proper delivery date")
            'Exit Sub
            ControlDeliveryDate = CDate("01/01/1970")
        End If
        ControlDeliveryRef = strDeliveryRef

        MakeVisible = True
        ControlType = "TEXTBOX"
        ControlText = CStr(OpID)
        ControlLeft = 7
        ControlTop = TopPos
        ControlHeight = 23
        ControlWidth = 29
        ControlFontSize = 10
        ControlFontName = "Cambria"
        ControlStyle = ""
        ControlTAG = CStr(TagID)
        ControlDate = Now()
        ControlDeliveryRef = strDeliveryRef
        ControlASN = ASN
        ControlOBJCount = OpID
        ControlStartTAG = CStr(LowerTAG)
        ControlEndTAG = CStr(UpperTAG)
        'BackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlForeColor = RGB(255, 245, 60) 'yellow text
        ControlLeftMargin = False
        ControlFieldname = FieldnameArr(StartFieldIndex)
        ControlRowNumber = OpID
        ControlTotalRows = TotalRows

        ComboArray = Nothing

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtOpRow:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER")
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlLeft = 41
        ControlTop = TopPos
        ControlWidth = 160
        ControlHeight = 23
        ControlTAG = CStr(TagID)
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0
        Firstname = ""
        Lastname = ""
        FullName = ""
        'ListArray = Get_DropdownItems(frmMainGIForm.myConnString, "SELECT Firstname,Lastname FROM tblEmployees ORDER BY Firstname", "", ErrMessage)
        'ReDim ComboArray(UBound(ListArray, 2) + 1)
        'For RowIDX = 0 To UBound(ListArray, 2)
        'If Not IsNothing(ListArray(0, RowIDX)) Then
        'Firstname = ListArray(0, RowIDX)
        'End If
        'If Not IsNothing(ListArray(1, RowIDX)) Then
        'Lastname = ListArray(1, RowIDX)
        'End If
        'FullName = Firstname & " " & Lastname
        'If Len(FullName) > 1 Then
        ''put into combo array
        'ComboArray(ComboIDX) = FullName
        'ComboIDX = ComboIDX + 1
        'End If
        'Next

        ControlType = "TEXTBOX" 'OP Name - combobox
        If Len(OpName) = 0 Then
            ControlText = "Select Employee"
        Else
            ControlText = OpName
        End If
        ControlStyle = ""
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlFieldname = FieldnameArr(StartFieldIndex + 1)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "comOperativeName:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        'ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
        ComboArray = Nothing 'call PopulateDowndrop.
        ComboIDX = 0
        strActivity = ""

        ListArray = Get_DropdownItems(frmMainGIForm.myConnString, "SELECT Activity FROM tblActivity", "", ErrMessage)
        ReDim ComboArray(UBound(ListArray, 2) + 1)
        If Not IsNothing(ListArray) Then
            For RowIDX = 0 To UBound(ListArray, 2)
                If Len(ListArray(0, RowIDX)) > 0 And Not IsNothing(ListArray(0, RowIDX)) Then
                    strActivity = ListArray(0, RowIDX).ToString
                    ComboArray(ComboIDX) = strActivity
                    ComboIDX = ComboIDX + 1
                End If
            Next
        End If
        ControlType = "COMBOBOX"
        If Len(OpActivity) = 0 Then
            ControlText = "Select Activity"
        Else
            ControlText = OpActivity
        End If
        ControlTAG = CStr(TagID)
        ControlStyle = ""
        ControlLeft = 206
        ControlWidth = 130
        ControlTop = TopPos
        ControlHeight = 23
        ControlFontName = "CAMBRIA"
        ControlFontSize = 10
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 2)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "comOperativeActivity:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlType = "BTN"
        ControlText = "@S"
        ControlTAG = "BTN_ST" & CStr(OpID)
        ControlLeft = 342
        ControlWidth = 40
        ControlTop = TopPos
        ControlHeight = 25
        ControlFontName = "CAMBRIA"
        ControlFontSize = 11
        ControlStyle = "BOLD,ITALIC"
        ControlBackColor = RGB(255, 215, 0) 'GOLD
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 3) 'SET to OpStartTime

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "btnOperativeStartTime:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER",
                                 "MIDDLE CENTER", DisableStartTimeButton, False, #1/1/1970 12:00:00 AM#)

        btnTAGID = btnTAGID + 1
        ControlTABIndex = ControlTABIndex + 1

        'txtOperativeTimeStart :

        TimeTAGID = TagID + TimeTAGStart
        ControlType = "TEXTBOX"
        If dtStartTime > #1970-01-01# Then
            ControlText = dtStartTime.ToString("HH:mm:ss")
        Else
            ControlText = "00:00:00"
        End If
        ControlTAG = CStr(TimeTAGID)
        ControlLeft = 392
        ControlWidth = 60
        ControlTop = TopPos
        ControlHeight = 23
        ControlStyle = ""
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 3) 'OpStartTime

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtOperativeStartTime:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor,
                                 "CENTER", "", DisableStartTimeEntry, False, dtStartTime, dtFinishTime)

        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlType = "BTN"
        ControlText = "@F"
        ControlTAG = "BTN_ET" & CStr(OpID)
        ControlLeft = 462
        ControlWidth = 40
        ControlTop = TopPos
        ControlHeight = 25
        ControlFontName = "Cambria"
        ControlFontSize = 11
        ControlStyle = "BOLD,ITALIC"
        ControlBackColor = RGB(255, 215, 0) 'GOLD
        ControlForeColor = RGB(0, 0, 0) 'BLACK
        ControlFieldname = FieldnameArr(StartFieldIndex + 4) 'OpFinishTime

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "btnOperativeFinishTime:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor,
                                 "CENTER", "MIDDLE CENTER", DisableFinishTimeButton)

        btnTAGID = btnTAGID + 1
        ControlTABIndex = ControlTABIndex + 1

        TimeTAGID = TagID + TimeTAGStart
        'TAGID = TAGID + 1
        'txtOperativeTimeEnd :

        ControlType = "TEXTBOX"
        If dtFinishTime > #1970-01-01# Then
            ControlText = dtFinishTime.ToString("HH:mm:ss")
        Else
            ControlText = "00:00:00"
        End If
        ControlTAG = CStr(TimeTAGID)
        ControlLeft = 512
        ControlWidth = 60
        ControlTop = TopPos
        ControlHeight = 23
        ControlStyle = ""
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 4) 'OpFinishTime

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtOperativeFinishTime:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor,
                                 "CENTER", "", DisableFinishTimeEntry, False, dtStartTime, dtFinishTime)

        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlType = "TEXTBOX"
        If Len(OpComment) = 0 Then
            ControlText = "None."
        Else
            ControlText = OpComment
        End If
        ControlTAG = CStr(TagID)
        ControlLeft = 582
        ControlWidth = 100
        ControlTop = TopPos
        ControlHeight = 23
        ControlStyle = ""
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlBackColor = RGB(240, 248, 253) 'ALICEBLUE -almost
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 5) 'Operative Comments

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtOpComments:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)

        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        If DontIncrementOP = False Then
            OpID = OpID + 1
        End If

        FrameScrollControl.AutoScroll = True
        'ScrollBarHeight = FrameScrollControl.VerticalScroll.Value
        'If OpID > 1 Then
        'roughly 100 = 5 rows
        'ScrollBarHeight = ScrollBarHeight + (100 / 5)
        'ScrollBarHeight = ScrollBarHeight + (OpID * 20)
        'FrameScrollControl.VerticalScroll.Value = ScrollBarHeight
        'End If

        'Set Dic_Collection = CreateObject("Scripting.Dictionary")
        'Dic_Collection.CompareMode = vbTextCompare

        TotalOps = OpID
        'TextTAGID = TagID
        StartTABIndex = ControlTABIndex

    End Sub

    Function InsertOperatives(AddNewRow As Boolean, strDeliveryDate As String, strDeliveryRef As String, strASNNo As String, OpFrameControl As ScrollableControl,
                         ByVal FrameRowNumber As Long,
                         Optional ByVal TimeStartTAG As Long = 400, Optional ByVal LowerTag As Long = 43,
                         Optional ByVal ControlsPerRow As Long = 6, Optional ByVal StartFieldIndex As Long = 4,
                         Optional ByVal OpName As String = "",
                         Optional ByVal OpActivity As String = "",
                         Optional ByVal dtStartTime As Date = #1970-01-01#, Optional dtFinishTime As Date = #1970-01-01#,
                         Optional ByVal OpComment As String = "None.",
                         Optional ByRef StartTABIndex As Long = 0,
                         Optional ByVal ResetFieldIndexTo4 As Boolean = True,
                         Optional ByVal DisableStartButton As Boolean = False,
                         Optional ByVal DisableFinishButton As Boolean = False,
                         Optional ByVal DisableStartTimeEntry As Boolean = True,
                         Optional ByVal DisableFinishTimeEntry As Boolean = True,
                         Optional ByVal FormName As String = "GI_TIMESHEET_") As Long
        Dim UpperTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim StartBtnTAG As Long = 0
        Dim HighestOpTag As Long = 0
        Dim TotalControlsInFrame As Long = 0
        Dim SearchText As String = ""
        Dim SearchField As String = ""
        Dim FieldType As String = "STRING"
        Dim ReturnField As String = "ID"
        Dim ReturnValue As String
        Dim SearchCriteria As String = ""
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String
        Dim LabourFieldsArr() As String
        Dim OpLoadedOK As Boolean = False
        Dim LabLoadedOK As Boolean = False
        Dim ControlDBTable As String = "tblOperatives"
        Dim OpDBTable As String = "tblOperatives"
        Dim LabourDBTable As String = "tblLabourHours"
        Dim LabourRecID As String = ""
        Dim OpRecID As String = ""
        Dim LabAllValues() As Object = Nothing
        Dim LabAllFields() As String = Nothing
        Dim OpAllValues() As Object = Nothing
        Dim OpAllFields() As String = Nothing
        Dim OpSaveCriteria As String
        Dim LabSaveCriteria As String
        Dim LabourFields As String
        Dim LabourFieldValues As String
        Dim OpFields As String
        Dim OpFieldValues As String
        Dim ControlTotalRows As Long = 0
        Dim dtNow As DateTime
        Dim dtDeliveryDate As DateTime
        Dim LabourSavedOK As Boolean = False
        Dim OpSavedOK As Boolean = False
        Dim NewFrameRowNumber As Long = 0
        Dim localOPID As Long
        Dim ExtractTotals As New clsTotals
        Dim tempControl As New clsControls
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim strSavedDeliveryDate As String
        Dim strUsername As String
        Dim strNAME As String
        Dim strEMPNO As String
        Dim strControlKey As String
        Dim strTotalRows As String
        Dim strFLMTotalRows As String
        Dim lngFLMTotalRows As Long
        Dim ControlTABIndex As Long
        Dim strSQL As String
        Dim ReturnFromSql As Object
        Dim OpTotalRows As Long
        Dim ChangeProperty As String
        Dim ControlName As String
        Dim testCollection As Object
        Dim OpControl As Control
        Dim TotalRows As Long
        Dim FieldIndex As Long

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:
        InsertOperatives = 0

        UpperTAG = 0
        'TotalRows = 0
        StartBtnTAG = 0
        HighestOpTag = 0
        ControlTABIndex = StartTABIndex

        If Len(strDeliveryDate) = 0 Then
            MsgBox("Delivery Date not supplied")
            Exit Function
        End If

        If ResetFieldIndexTo4 Then
            FieldIndex = 4
        Else
            FieldIndex = StartFieldIndex
        End If

        strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")

        'tblLabourHours search:
        If Len(strDeliveryRef) > 0 Then
            'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            strSavedDeliveryDate = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            SearchText = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = ""
            SortFields = "DeliveryDate"
            Reversed = True
        Else
            strSavedDeliveryDate = CDate("1970-01-01 00:00:00")
        End If

        SortFields = ""
        Reversed = False
        ExcludeFields = ""

        'LabourFieldsArr = strToStringArray(frmMainGIForm., ",", 0, False, False, False, "_", False, 34, 39)

        OpTotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)

        LabLoadedOK = Find_myQuery(frmMainGIForm.myConnString, LabourDBTable, SearchField, SearchText, FieldType, ReturnField, LabourRecID, LabAllValues,
            LabAllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        strFLMTotalRows = GetMYValuebyFieldname(LabAllValues, LabAllFields, "Total_Rows") 'FROM FLM DETAILS TABLE: tblLabourHours
        If IsNumeric(strFLMTotalRows) Then
            'SO TotalRows will be comming from the tblLabourHours table - THIS will be the current amount of Operatives.
            lngFLMTotalRows = CLng(strFLMTotalRows)
        Else
            lngFLMTotalRows = 0
        End If
        ' - OK first we have strFLMTotalRows -> lngFLMTotalRows = Total amount of rows in the table that has just loaded;
        ' - This procedure is repeated while the correct number of rows are being inserted inside the loop until lngFLMTotalRows.
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        If ExtractTotals IsNot Nothing Then 'Which should now be 0 if this is the first row.
            If AddNewRow Then
                'TotalRows = lngFLMTotalRows + 1
                TotalRows = OpTotalRows + 1
            Else
                TotalRows = OpTotalRows
            End If
            ExtractTotals.Total_Operatives = TotalRows
            'TotalRows = ExtractTotals.Total_Operatives - 1
            HighestOpTag = ExtractTotals.HighestOpTAGID
            StartBtnTAG = ExtractTotals.HighestOpBtnTAGID
        Else
            'TotalRows = lngFLMTotalRows + 1 'SO if both variables start at 0 then TotalsRows should be = 1.
            TotalRows = OpTotalRows + 1
            HighestOpTag = LowerTag
            StartBtnTAG = 10
        End If
        'Is this correct ????
        NewFrameRowNumber = TotalRows

        strControlKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(40)
        tempControl = dic_Controls(strControlKey)
        strUsername = ""
        strNAME = ""
        strEMPNO = ""
        If tempControl IsNot Nothing Then
            strUsername = tempControl.ControlUpdatedByUsername
            strNAME = tempControl.ControlUpdatedByName
            strEMPNO = tempControl.ControlUpdatedByEmpNo
        End If

        TotalControlsInFrame = Math.Abs((TotalRows * ControlsPerRow))
        ControlTotalRows = TotalRows

        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTag
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTag
        End If
        'First Parameter = Total Operatives value
        'Second Parameter = NEW ROW NUMBER
        'so in theory Total Operatives will be 0 for the first INSERT but the second paramter will be 1 for the FIRST ROW:
        ' - NOW FrameRowNumber is the value passed into this function.
        ' - if called from PopulateControls it will be a LOOP counter variable for EACH row that exists in tblOperatives matching the chosen reference.
        ' - if called from ADD Operative or from SearchAndInsertName() this will be the NEXT / NEW ROW to be added.
        localOPID = FrameRowNumber
        AddNewOperatives(localOPID, localOPID, OpFrameControl, HighestOpTag, StartBtnTAG, ControlTABIndex,
                         strDeliveryDate, strDeliveryRef, strASNNo, LowerTag, UpperTAG, TimeStartTAG, OpFieldnames, TotalRows,
                         FieldIndex, False, ErrMessage, 0, OpName, OpActivity, dtStartTime, dtFinishTime, OpComment, 6,
                         DisableStartButton, DisableFinishButton, DisableStartTimeEntry, DisableFinishTimeEntry)
        'ExtractTotals.Total_Operatives = localOPID
        ExtractTotals.HighestOpTAGID = HighestOpTag
        ExtractTotals.HighestOpBtnTAGID = StartBtnTAG
        ExtractTotals.OpStartTAG = LowerTag
        ExtractTotals.OpFinishTAG = HighestOpTag
        ExtractTotals.HighestOpTabIndex = ControlTABIndex
        'TotalRows = localOPID - 1
        dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = ExtractTotals

        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalOps", CStr(TotalRows))
        ControlTotalRows = TotalRows

        'INSERT NEW RECORD if not exists:
        'NEED TO LOCATE THE FLM RECORD THAT ALREADY EXISTS:

        If Len(strDeliveryDate) > 0 Then
            'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            strSavedDeliveryDate = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            SearchText = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = ""
            SortFields = "DeliveryDate"
            Reversed = True
        Else
            strSavedDeliveryDate = CDate("1970-01-01 00:00:00")
        End If

        SortFields = ""
        Reversed = False
        ExcludeFields = ""

        'LabourFieldsArr = strToStringArray(frmMainGIForm., ",", 0, False, False, False, "_", False, 34, 39)

        LabLoadedOK = Find_myQuery(frmMainGIForm.myConnString, LabourDBTable, SearchField, SearchText, FieldType, ReturnField, LabourRecID, LabAllValues,
                                   LabAllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        If Len(strDeliveryDate) > 0 Then
            'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            SearchText = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SortFields = "DeliveryDate"
            Reversed = True
            SearchCriteria = "FrameRowNumber = " & NewFrameRowNumber
        End If

        'SortFields = ""
        'Reversed = False
        'ExcludeFields = ""


        OpLoadedOK = Find_myQuery(frmMainGIForm.myConnString, OpDBTable, SearchField, SearchText, FieldType, ReturnField, OpRecID, OpAllValues,
                                   OpAllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If OpLoadedOK Then
            OpSaveCriteria = "ID = " & OpRecID
            'Need TAG number or control name to insert primary key value into Dic_Controls():
            '************** OK we have the NewFrameRowNumber - just need to pick a control name to alter: *****************
            ControlName = "txtOpRow:" & CStr(FrameRowNumber)
            OpControl = frmMainGIForm.FindFrameControls(FormName & FormID, ControlName)
            If OpControl IsNot Nothing Then
                ChangeProperty = "PRIMARYKEY"
                testCollection = UpdateCollection(dic_Controls, Nothing, OpRecID, CDate(strDeliveryDate), "",
                                                      strDeliveryRef, OpControl.Tag, ChangeProperty)

                If testCollection IsNot Nothing Then
                    dic_Controls = testCollection
                End If
            End If
        Else
            OpSaveCriteria = ""
        End If
        LabSaveCriteria = ""
        'tblLabourHours: BLANK record setup:
        If LabLoadedOK Then
            'UPDATE LABOUR record:
            strLastSaved = GetMYValuebyFieldname(OpAllValues, OpAllFields, "LastSaved")
            If IsDate(strLastSaved) Then
                dtLastSaved = CDate(strLastSaved)
            Else
                dtLastSaved = Now()
            End If
            'dtLastSaved = Now()
            LabourFields = "Total_Rows,Start_TAGID,END_TAGID,LastSaved"
            'LabourFields = "Total_Rows,Start_TAGID,END_TAGID,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"
            'LabourFieldValues = Chr(34) & ControlTotalRows & Chr(34)
            LabourFieldValues = ControlTotalRows & "," & LowerTag & "," & UpperTAG
            'LabourFieldValues = LabourFieldValues & "," & Chr(34) & strUsername & Chr(34) & "," & Chr(34) & strEMPNO & Chr(34) & "," & Chr(34) & strNAME & Chr(34)
            'LabourFieldValues = LabourFieldValues & "," & Chr(39) & dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39)
            LabourFieldValues = LabourFieldValues & "," & dtLastSaved
            If Len(LabourRecID) > 0 Then
                LabSaveCriteria = "ID = " & LabourRecID
            End If
            'ARE WE LOADING IN EXISTING OPERATIVES HERE ALSO ?????
            ' - IF THIS PROCEDURE IS CALLED DURING POPULATECONTROLS - MAY HAVE VALUES ALREADY FOR THE ROW NUMBER,comNAME,comACTIVITY,txtSTARTTIME,txtENDTIME,txtCOMMENT
            OpFields = "LinkID,DeliveryDate,DeliveryReference,FrameRowNumber,OpComments,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"

            OpFieldValues = CLng(LabourRecID) & "," & CDate(strSavedDeliveryDate) & "," & strDeliveryRef
            OpFieldValues = OpFieldValues & "," & NewFrameRowNumber & "," & OpComment
            OpFieldValues = OpFieldValues & "," & strUsername & "," & strEMPNO & "," & strNAME
            OpFieldValues = OpFieldValues & "," & dtLastSaved

        Else 'NO FLM RECORD EXISTS FOR THIS REF:
            dtDeliveryDate = CDate(strSavedDeliveryDate)
            dtNow = Now()

            LabourFields = "DeliveryDate,DeliveryReference,Total_Rows,Start_TAGID,END_TAGID,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"

            LabourFieldValues = CDate(strSavedDeliveryDate) & "," & strDeliveryRef & "," & ControlTotalRows
            LabourFieldValues = LabourFieldValues & "," & LowerTag & "," & UpperTAG
            LabourFieldValues = LabourFieldValues & "," & strUsername & "," & strEMPNO & "," & strNAME
            LabourFieldValues = LabourFieldValues & "," & dtNow

            ExcludeFields = ""



            If IsNumeric(LabourRecID) Then
                OpFields = "DeliveryDate,DeliveryReference,FrameRowNumber,LinkID,OpComments,LastSaved"
                OpFieldValues = CDate(strSavedDeliveryDate) & "," & strDeliveryRef
                OpFieldValues = OpFieldValues & "," & NewFrameRowNumber & "," & CLng(LabourRecID) & "," & OpComment
                OpFieldValues = OpFieldValues & "," & dtNow
            Else
                OpFields = "DeliveryDate,DeliveryReference,FrameRowNumber,OpComments,LastSaved"
                OpFieldValues = CDate(strSavedDeliveryDate) & "," & strDeliveryRef
                OpFieldValues = OpFieldValues & "," & NewFrameRowNumber & "," & OpComment
                OpFieldValues = OpFieldValues & "," & dtNow
            End If

        End If
        'SHOULD BE NEW FRAME ROW NUMBER = TOTAL RECORDS AS WE ARE INSERTING A NEW OP ROW - NOT UPDATING IT !

        If Len(strDeliveryRef) > 0 And AddNewRow Then
            'LabourSavedOK = InsertUpdateMyRecord(LabLoadedOK, frmMainGIForm.myConnString, LabourDBTable, LabourFields, LabourFieldValues, ErrMessage, LabSaveCriteria,
            'ExcludeFields,)

            LabourSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, LabLoadedOK, "", LabourDBTable, LabourFields, LabourFieldValues,
                                                                    LabSaveCriteria, ExcludeFields, ErrMessage, False, ",")

            If Not LabourSavedOK Then
                'MsgBox("Labour Hours did NOT initialise/Update OK: " & ErrMessage)
                frmMainGIForm.logger.LogError("GI_ERRORS_" & frmMainGIForm.myVersion & ".log", Application.StartupPath, ErrMessage,
                                              "LabourSavedOK: InsertOperatives()",
                                              frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername,
                                              frmMainGIForm.myUsernameSuffix)
                Exit Function
            End If
            'OpSavedOK = InsertUpdateMyRecord(OpLoadedOK, frmMainGIForm.myConnString, OpDBTable, OpFields, OpFieldValues, ErrMessage, OpSaveCriteria,
            'ExcludeFields,)

            OpSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, OpLoadedOK, "", OpDBTable, OpFields, OpFieldValues,
                                                                    OpSaveCriteria, ExcludeFields, ErrMessage, False, ",")

            If Not OpSavedOK Then
                'ERROR - MsgBox("Operatives did NOT initialise OK: " & ErrMessage)
                frmMainGIForm.logger.LogError("GI_ERRORS_" & frmMainGIForm.myVersion & ".log", Application.StartupPath,
                                              ErrMessage, "not OpSAVEDOK: InsertOperatives()",
                                              frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername,
                                              frmMainGIForm.myUsernameSuffix)
                Exit Function
            Else
                SearchCriteria = "FrameRowNumber = " & NewFrameRowNumber
                OpLoadedOK = Find_myQuery(frmMainGIForm.myConnString, OpDBTable, SearchField, SearchText, FieldType, ReturnField, OpRecID, OpAllValues,
                                   OpAllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                If OpLoadedOK Then
                    'OpSaveCriteria = "ID = " & OpRecID
                    'Need TAG number or control name to insert primary key value into Dic_Controls():
                    '************** OK we have the NewFrameRowNumber - just need to pick a control name to alter: *****************
                    ControlName = "txtOpRow:" & CStr(NewFrameRowNumber)
                    OpControl = frmMainGIForm.FindFrameControls(FormName & FormID, ControlName)
                    If OpControl IsNot Nothing Then
                        ChangeProperty = "PRIMARYKEY"
                        testCollection = UpdateCollection(dic_Controls, Nothing, OpRecID, CDate(strDeliveryDate), "",
                                                              strDeliveryRef, OpControl.Tag, ChangeProperty)

                        If testCollection IsNot Nothing Then
                            dic_Controls = testCollection
                        End If
                    End If
                End If
            End If
        End If
        'tblOperataves blank record setup:

        InsertOperatives = TotalRows

    End Function

    Sub InsertOperatives_Into_Grid(strDeliveryDate As String, strDeliveryRef As String, strASNNo As String, OpFrameControl As ScrollableControl,
                         ByVal FrameRowNumber As Long,
                         Optional ByVal FormName As String = "GI_TIMESHEET_",
                         Optional ByVal TimeStartTAG As Long = 400, Optional ByVal LowerTag As Long = 43,
                         Optional ByVal ControlsPerRow As Long = 6, Optional ByVal StartFieldIndex As Long = 4,
                         Optional ByVal OpName As String = "",
                         Optional ByVal OpActivity As String = "",
                         Optional ByVal dtStartTime As Date = #1970-01-01#, Optional dtFinishTime As Date = #1970-01-01#,
                         Optional ByVal OpComment As String = "None.",
                         Optional ByRef StartTABIndex As Long = 1,
                         Optional ByVal ResetFieldIndexTo4 As Boolean = True,
                         Optional ByVal DisableStartButton As Boolean = False,
                         Optional ByVal DisableFinishButton As Boolean = False,
                         Optional ByVal DisableStartTimeEntry As Boolean = True,
                         Optional ByVal DisableFinishTimeEntry As Boolean = True,
                         Optional ByVal AddNewRow As Boolean = False)

        'AMENDED: 03-APR-2019 at 09:20
        'Change: LabourRecID and AddNewRow were not declared ???? !!!
        'Happend after Procedure to change the STATUS by user - was amended to set LAST_SAVED to the current date and time.

        Dim OpSavedOK As Boolean = False
        Dim OpLoadedOK As Boolean = False
        Dim NewFrameRowNumber As Long = 0
        Dim localOPID As Long
        Dim ExtractTotals As New clsTotals
        Dim tempControl As New clsControls
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim strSavedDeliveryDate As String
        Dim strUsername As String
        Dim strNAME As String
        Dim strEMPNO As String
        Dim strControlKey As String
        Dim strTotalRows As String
        Dim strFLMTotalRows As String
        Dim lngFLMTotalRows As Long
        Dim ControlTABIndex As Long
        Dim strSQL As String
        Dim ReturnFromSql As Object
        Dim OpTotalRows As Long
        Dim ChangeProperty As String
        Dim ControlName As String
        Dim testCollection As Object
        Dim OpControl As Control
        Dim TotalRows As Long
        Dim FieldIndex As Long
        Dim SearchText As String
        Dim SearchField As String
        Dim FieldType As String
        Dim ReturnField = "ID"
        Dim ReturnValue As String
        Dim SearchCriteria As String = ""
        Dim SortFields As String = "DeliveryDate"
        Dim Reversed As Boolean = True
        Dim ExcludeFields As String = ""
        Dim LabLoadedOK As Boolean = False
        Dim LabourDBTable As String = "tblLabourHours"
        Dim LabAllValues As Object() = Nothing
        Dim LabAllFields() As String = Nothing
        Dim HighestOpTag As Long = 0
        Dim ErrMessage As String = ""
        Dim StartBtnTAG As Long
        Dim TotalControlsInFrame As Long
        Dim ControlTotalRows As Long
        Dim UpperTAG As Long
        Dim OpDBTable = "tblOperatives"
        Dim OpAllFields() As String = Nothing
        Dim OpRecID As String
        Dim OpAllValues() As Object = Nothing
        Dim OpSaveCriteria As String = ""
        Dim LabSaveCriteria As String = ""
        Dim LabourFields As String
        Dim LabourFieldValues As String
        Dim LabourRecID As String


        '1) Get total rows in DB TABLE: tblOperatives
        '2) Initialise tempControls variable and set to TAG=43
        '3) Get entries from each of the controls in clsControls / Dic_Controls()
        '4) A NAME.
        '5) An Activity.
        '6) Optional - A START TIME.
        '7) TEST if NAME AND / OR ACTIVITY already exists + test start TIME and End TIME if in DB TABLE.
        '8) Insert into DB Table : tblOperatives and tblLabourHours AND hours into tblDeliveryInfo

        'JUST SAVE TO tblOperatives !!!

        '1) TEST if operative already exists ??? - Check OpName already entered. Check Activity . Check Start Time. check End Time.
        '2) Total Times not calculating correctly - seems to be adding 1 hour to total for EACH Operative - BUT overall Total hrs seem ok.

        If Len(strDeliveryRef) > 0 Then
            'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            strSavedDeliveryDate = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
            SearchText = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = ""
            SortFields = "DeliveryDate"
            Reversed = True
        Else
            strSavedDeliveryDate = CDate("1970-01-01 00:00:00")
        End If

        SortFields = ""
        Reversed = False
        ExcludeFields = ""

        'LabourFieldsArr = strToStringArray(frmMainGIForm., ",", 0, False, False, False, "_", False, 34, 39)

        OpTotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)

        LabLoadedOK = Find_myQuery(frmMainGIForm.myConnString, LabourDBTable, SearchField, SearchText, FieldType, ReturnField, LabourRecID, LabAllValues,
            LabAllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        strFLMTotalRows = GetMYValuebyFieldname(LabAllValues, LabAllFields, "Total_Rows") 'FROM FLM DETAILS TABLE: tblLabourHours
        If IsNumeric(strFLMTotalRows) Then
            'SO TotalRows will be comming from the tblLabourHours table - THIS will be the current amount of Operatives.
            lngFLMTotalRows = CLng(strFLMTotalRows)
        Else
            lngFLMTotalRows = 0
        End If
        ' - OK first we have strFLMTotalRows -> lngFLMTotalRows = Total amount of rows in the table that has just loaded;
        ' - This procedure is repeated while the correct number of rows are being inserted inside the loop until lngFLMTotalRows.
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        If ExtractTotals IsNot Nothing Then 'Which should now be 0 if this is the first row.
            If AddNewRow Then
                'TotalRows = lngFLMTotalRows + 1
                TotalRows = OpTotalRows + 1
            Else
                TotalRows = OpTotalRows
            End If
            ExtractTotals.Total_Operatives = TotalRows
            'TotalRows = ExtractTotals.Total_Operatives - 1
            HighestOpTag = ExtractTotals.HighestOpTAGID
            StartBtnTAG = ExtractTotals.HighestOpBtnTAGID
        Else
            'TotalRows = lngFLMTotalRows + 1 'SO if both variables start at 0 then TotalsRows should be = 1.
            TotalRows = OpTotalRows + 1
            HighestOpTag = LowerTag
            StartBtnTAG = 10
        End If
        'Is this correct ????
        NewFrameRowNumber = TotalRows

        strControlKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(40)
        tempControl = dic_Controls(strControlKey)
        strUsername = ""
        strNAME = ""
        strEMPNO = ""
        If tempControl IsNot Nothing Then
            strUsername = tempControl.ControlUpdatedByUsername
            strNAME = tempControl.ControlUpdatedByName
            strEMPNO = tempControl.ControlUpdatedByEmpNo
        End If

        TotalControlsInFrame = Math.Abs((TotalRows * ControlsPerRow))
        ControlTotalRows = TotalRows

        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTag
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTag
        End If

    End Sub

    Function Get_TotalRows(TableName As String, strDeliveryRef As String, Optional Criteria As String = "") As Long
        Dim strSQL As String
        Dim ReturnFromSQL As Object
        Dim ReturnValue As String
        Dim TotalRows As Long = 0
        Dim ErrMessage As String = ""

        strSQL = "SELECT DeliveryReference,Count(ID) FROM " & TableName & " WHERE DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34)
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " AND " & Criteria
        End If
        ReturnFromSQL = MySQLToArray(frmMainGIForm.myConnString, strSQL, ErrMessage)
        ReturnValue = ReturnFromSQL(1, 0)
        If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
            TotalRows = CLng(ReturnValue)
        End If

        Get_TotalRows = TotalRows

    End Function

    Function Get_Reference_Lock_Info(FromLockTable As Boolean, strDeliveryDate As String, strDeliveryRef As String, ByRef CompleteSTATUS As String,
                                     Optional ByRef RecLockSTATUS As String = "",
                                     Optional ByRef Username As String = "",
                                     Optional ByRef UserEmpNo As String = "",
                                     Optional ByRef UserFullName As String = "",
                                     Optional ByRef StartEditTime As DateTime = #1970-01-01#, Optional ByRef EndEditTime As DateTime = #1970-01-01#) As Boolean
        Dim SearchText As String = ""
        Dim SearchField As String = ""
        Dim FieldType As String = "STRING"
        Dim ReturnField As String = "ID"
        Dim ReturnValue As String
        Dim SearchCriteria As String = ""
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String
        Dim RefFieldsArr() As String
        Dim LoadedOK As Boolean = False
        Dim ErrMessage As String = ""
        Dim refLockTableName As String = "tblReferenceLock"
        Dim DelInfoTableName As String = "tblDeliveryInfo"
        Dim DBTable As String
        Dim RecID As String
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim strStartEditTime As String
        Dim strEndEditTime As String

        Get_Reference_Lock_Info = False
        If FromLockTable = False Then
            DBTable = DelInfoTableName
        Else
            DBTable = refLockTableName
        End If

        If Len(strDeliveryDate) = 0 Then
            MsgBox("NO DELIVERY DATE given")
            Exit Function
        End If
        If Len(strDeliveryRef) = 0 Then
            MsgBox("NO DELIVERY REF")
            Exit Function
        End If
        'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")
        'SearchText = UCase("DeliveryReference = " & Chr(39) & strDeliveryRef & Chr(39))
        SearchText = strDeliveryRef
        SearchField = "DeliveryReference"
        FieldType = "STRING"
        ReturnField = "ID"
        ReturnValue = ""
        SearchCriteria = ""
        'SearchCriteria = "DATE(DeliveryDate) = " & CDate(strDeliveryDate)
        AllValues = Nothing
        AllFields = Nothing
        SortFields = "DeliveryDate"
        Reversed = True
        RecID = ""

        LoadedOK = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, FieldType, ReturnField, RecID, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        If LoadedOK Then
            'OK Reference exists in DB.
            'Is Someone is EDITING THIS REFERENCE ?
            Get_Reference_Lock_Info = False
            CompleteSTATUS = GetMYValuebyFieldname(AllValues, AllFields, "Status")
            RecLockSTATUS = GetMYValuebyFieldname(AllValues, AllFields, "Lock_Status")
            Username = GetMYValuebyFieldname(AllValues, AllFields, "UpdatedByUsername")
            UserFullName = GetMYValuebyFieldname(AllValues, AllFields, "UpdatedByName")
            UserEmpNo = GetMYValuebyFieldname(AllValues, AllFields, "UpdatedByEmpNo")
            If FromLockTable = True Then
                strStartEditTime = GetMYValuebyFieldname(AllValues, AllFields, "StartEditTime")
                If IsDate(strStartEditTime) Then
                    StartEditTime = CDate(strStartEditTime)
                End If
                strEndEditTime = GetMYValuebyFieldname(AllValues, AllFields, "FinishEditTime")
                If IsDate(strEndEditTime) Then
                    EndEditTime = CDate(strEndEditTime)
                End If
            End If
            Get_Reference_Lock_Info = True
        Else
            'Cannot Find Delivery Reference or Delivery Date for TABLE PASSED:

            Get_Reference_Lock_Info = False


        End If

    End Function

    Sub AddNewShorts(ByRef TotalShorts As Long, FrameScrollControl As ScrollableControl, ByRef ShortID As Long, ByRef TagID As Long, ByRef btnTAGID As Long,
                         strDeliveryDate As String, strDeliveryRef As String, ASN As String, LowerTAG As Long, UpperTAG As Long, ByVal TimeTAGStart As Long,
                         Fieldnames As String, TotalRows As Long, Optional StartFieldIndex As Long = 3, Optional DontIncrement As Boolean = False,
                         Optional ByRef ErrMessage As String = "", Optional ByRef NewIndex As Long = 0,
                         Optional ByVal PartNo As String = "", Optional ByVal PartQty As String = "", Optional ByRef StartTABIndex As Long = 1,
                         Optional ByVal TotalOpTEXTControlsPerRow As Long = 3)
        Dim RowGap As Long
        Dim TopPos As Integer
        Dim StartTopPos As Integer
        Dim ScrollBarHeight As Integer
        Dim ComboArray() As String
        Dim TimeTAGID As Long
        Dim ControlText As String
        Dim ControlType As String
        Dim ControlTAG As String
        Dim ControlDate As Date
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlDeliveryDate As Date
        Dim ControlDeliveryRef As String
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlRowNumber As Long
        Dim ControlTotalRows As Long
        Dim MakeVisible As Boolean
        Dim ControlBackColor As Integer
        Dim ControlForeColor As Integer
        Dim ControlLeftMargin As Boolean
        Dim ControlFieldname As String
        Dim FieldnameArr() As String
        Dim LoadFieldsOK As Boolean
        Dim ControlFieldsTable As String
        Dim SearchCriteria As String
        Dim ControlDBTable As String
        Dim ControlFontName As String
        Dim ControlFontSize As Integer
        Dim ControlStyle As String
        Dim FrameRowNumberField As String = ""
        Dim IDX As Long
        Dim RowIDX As Long
        Dim Firstname As String
        Dim Lastname As String
        Dim FullName As String
        Dim strActivity As String
        Dim ControlTABIndex As Long

        ReDim FieldnameArr(1)
        ControlTABIndex = StartTABIndex
        ControlDBTable = "tblshortsandextraparts"
        If Len(Fieldnames) > 0 Then 'wrong fields !
            'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        Else
            ReDim FieldnameArr(20)
            Fieldnames = GetMyFields(ControlDBTable, frmMainGIForm.myConnString, ErrMessage)
            Beep()
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        End If
        FrameRowNumberField = "FrameRowNumber"
        StartTopPos = 2
        RowGap = 28
        If ShortID < 2 Then
            TopPos = StartTopPos
        Else
            TopPos = ((ShortID - 1) * RowGap) + StartTopPos
        End If

        TagID = Get_StartTAG(TimeTAGStart, ShortID, TotalOpTEXTControlsPerRow, 1)

        ControlFontName = "Cambria"

        If IsDate(strDeliveryDate) Then
            ControlDeliveryDate = CDate(strDeliveryDate)
            strDeliveryDate = ControlDeliveryDate.ToString("dd/MM/yyyy")
        Else
            ControlDeliveryDate = CDate("01/01/1970")
            strDeliveryDate = ControlDeliveryDate.ToString("dd/MM/yyyy")
        End If
        ControlDeliveryRef = strDeliveryRef

        'ControlTAG getting the wrong TAG no. - highest Tag is being passed - for row 1 - should be at 1001 NOT 1003.

        MakeVisible = True
        ControlType = "TEXTBOX"
        ControlText = CStr(ShortID)
        ControlLeft = 6
        ControlTop = TopPos
        ControlHeight = 23
        ControlWidth = 35
        ControlFontSize = 11
        ControlStyle = "BOLD"
        ControlTAG = CStr(TagID)
        ControlDate = Now()
        ControlDeliveryRef = strDeliveryRef
        ControlASN = ASN
        ControlOBJCount = ShortID
        ControlStartTAG = CStr(LowerTAG)
        ControlEndTAG = CStr(UpperTAG)
        ControlTotalRows = TotalRows ' need to know lowerTAG and number of fields in frame_Operatives.
        'BackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlForeColor = RGB(255, 245, 60) 'yellow text
        ControlLeftMargin = False
        ControlFieldname = FrameRowNumberField
        ControlRowNumber = ShortID
        ControlTotalRows = TotalShorts
        ComboArray = Nothing

        Call AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtShortRow:" & CStr(ShortID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", True)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlLeft = 48
        ControlTop = TopPos
        ControlWidth = 120
        ControlHeight = 23
        ControlTAG = CStr(TagID)
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing

        ControlType = "TEXTBOX"
        ControlText = PartNo
        ControlStyle = "BOLD"
        ControlFontName = "Cambria"
        ControlFontSize = 11
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE, needs lt grey
        ControlFieldname = FieldnameArr(StartFieldIndex + 1)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtShortPartNo:" & CStr(ShortID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", False)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        'ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
        ComboArray = Nothing
        strActivity = ""

        ControlType = "TEXTBOX"
        ControlText = ""
        ControlTAG = CStr(TagID)
        ControlStyle = "BOLD"
        ControlLeft = 174
        ControlWidth = 50
        ControlTop = TopPos
        ControlHeight = 23
        ControlFontName = "CAMBRIA"
        ControlFontSize = 11
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 2)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtShortQty:" & CStr(ShortID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", False)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        If DontIncrement = False Then
            ShortID = ShortID + 1
        End If

        FrameScrollControl.AutoScroll = True

        TotalShorts = ShortID
        StartTABIndex = ControlTABIndex

    End Sub

    Function InsertExtraParts2(AddNewRow As Boolean, strDeliveryDate As String, strDeliveryRef As String, strASNNo As String,
                               ExtraFrameControl As ScrollableControl,
                         ByVal FrameRowNumber As Long,
                         Optional ByVal TimeStartTAG As Long = 2000,
                         Optional ByVal LowerTag As Long = 2001,
                         Optional ByVal ControlsPerRow As Long = 3,
                         Optional ByVal StartFieldIndex As Long = 5,
                         Optional ByVal PartNo As String = "",
                         Optional ByVal intPartQty As Integer = 0,
                         Optional ByRef StartTABIndex As Long = 1,
                         Optional ByVal ResetFieldIndex As Boolean = True,
                         Optional ByVal FormName As String = "GI_TIMESHEET_") As Long

        Dim ExtraID As Long = 0
        Dim UpperTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim StartBtnTAG As Long = 0
        Dim HighestExtraTag As Long = 0
        Dim TotalControlsInFrame As Long = 0
        Dim tempTotals As clsTotals
        Dim ControlDBTable As String
        Dim dtDeliveryDate As DateTime
        Dim ExtraTotalRows As Long
        Dim searchtext = strDeliveryRef
        Dim SearchField As String = "DeliveryReference"
        Dim FieldType As String = "STRING"
        Dim ReturnField As String = "ID"
        Dim ReturnValue As String = ""
        Dim SearchCriteria As String = ""
        Dim TotalRowCriteria As String = ""
        Dim SortFields As String = ""
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String = ""
        Dim strExtraTotalRows As String
        Dim lngExtraTotalRows As Long
        Dim ExtraLoadedOK As Boolean = False
        Dim ExtraSavedOK As Boolean = False
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim TotalRowsFromQuery As Long
        Dim NewFrameRowNumber As Long
        Dim strExtraPartNo As String
        Dim strExtraQTY As String
        Dim ExtraRecID As String
        Dim ControlTotalRows As Long
        Dim dtLastSaved As DateTime
        Dim ExtraFields As String
        Dim ExtraValues As String
        Dim SaveCriteria As String
        Dim SaveDeliveryDate As String
        Dim strUsername As String
        Dim strEMPNO As String
        Dim strName As String
        Dim strControlKey As String
        Dim tempControl As clsControls
        Dim ControlTabIndex As Long
        Dim FieldIndex As Long
        Dim ExtractTotals As clsTotals
        Dim TotalRows As Long

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:
        'THIS IS THE PROC THAT GETS CALLED FROM THE MAIN CONTROL PANEL !!!

        InsertExtraParts2 = 0

        LowerTag = 2001
        UpperTAG = 0
        'TotalRows = 0
        StartBtnTAG = 0
        HighestExtraTag = 0
        TimeStartTAG = 2000
        ControlTabIndex = StartTABIndex
        If Len(strDeliveryDate) = 0 Then
            MsgBox("Delivery Date Not Supplied")
            Exit Function
        End If

        If ResetFieldIndex Then
            FieldIndex = 4
        Else
            FieldIndex = StartFieldIndex
        End If

        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        ControlDBTable = "tblshortsandextraparts"
        SaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")

        If Len(strDeliveryRef) > 0 Then
            searchtext = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = "ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
            SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & FrameRowNumber
            SortFields = ""
            Reversed = False
            ExcludeFields = ""
            AllValues = Nothing
            AllFields = Nothing
        Else
            SaveDeliveryDate = CDate("1970-01-01 00:00:00")
        End If

        ExtraRecID = ""

        TotalRowCriteria = "ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
        ExtraTotalRows = Get_TotalRows("tblShortsAndExtraParts", strDeliveryRef, TotalRowCriteria)

        If dic_Totals(strDeliveryDate & "_" & strDeliveryRef) IsNot Nothing Then
            ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If ExtractTotals IsNot Nothing Then
                If AddNewRow Then
                    TotalRows = ExtraTotalRows + 1
                Else
                    TotalRows = ExtraTotalRows
                End If
                ExtractTotals.Total_ExtraParts = TotalRows
                HighestExtraTag = ExtractTotals.HighestExtraTAGID
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = ExtractTotals
            Else
                'LOADED FROM DB:
                TotalRows = 1 'reset - default to 1.
                HighestExtraTag = LowerTag
            End If
        End If

        NewFrameRowNumber = TotalRows

        strControlKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(1)
        tempControl = dic_Controls(strControlKey)
        strUsername = ""
        strName = ""
        strEMPNO = ""
        If tempControl IsNot Nothing Then
            strUsername = tempControl.ControlUpdatedByUsername
            strName = tempControl.ControlUpdatedByName
            strEMPNO = tempControl.ControlUpdatedByEmpNo
        End If

        ExtraLoadedOK = Find_myQuery(frmMainGIForm.myConnString, ControlDBTable, SearchField, searchtext, FieldType, ReturnField, ExtraRecID,
                                         AllValues, AllFields, SearchCriteria, SortFields, Reversed, ErrMessage, "=", TotalRowsFromQuery)

        strExtraTotalRows = ""
        strExtraPartNo = ""
        strExtraQTY = ""
        If ExtraLoadedOK Then
            'A Extra PART record exists - get total rows and get the PART NO and QTY:
            strExtraTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
            strExtraPartNo = GetMYValuebyFieldname(AllValues, AllFields, "PartNo")
            strExtraQTY = GetMYValuebyFieldname(AllValues, AllFields, "Qty")
            HighestExtraTag = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")

            If IsNumeric(strExtraTotalRows) Then
                lngExtraTotalRows = CLng(strExtraTotalRows) 'TotalRows from DB
            Else
                lngExtraTotalRows = 0
            End If

        End If


        'when totalrows = 0 then Extraid will have to be 1 as it is the row number.
        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTag
            ExtraID = 1
            HighestExtraTag = 2001
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTag
            ExtraID = FrameRowNumber
        End If
        TotalControlsInFrame = Math.Abs(TotalRows * ControlsPerRow)
        ControlTotalRows = TotalRows

        AddNewExtra(TotalRows, ExtraFrameControl, ExtraID, HighestExtraTag, StartBtnTAG,
                            strDeliveryDate, strDeliveryRef, strASNNo, LowerTag, UpperTAG, TimeStartTAG, ShortAndExtraFieldnames,
                        TotalRows, FieldIndex, False, ErrMessage, 0, PartNo, intPartQty, StartTABIndex)
        tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        If tempTotals IsNot Nothing Then
            tempTotals.HighestExtraTAGID = HighestExtraTag
            tempTotals.HighestExtraTABIndex = StartTABIndex
        End If
        'frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef) = StartBtnTAG
        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalExtras", CStr(TotalRows - 1))
        If Len(CStr(FrameRowNumber)) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtExtraRow:" & CStr(ExtraID - 1), CStr(FrameRowNumber))
        End If
        If Len(PartNo) > 0 Then
            strExtraPartNo = PartNo
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtExtraPartNo:" & CStr(ExtraID - 1), CStr(strExtraPartNo))
        ElseIf Len(strExtraPartNo) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtExtraPartNo:" & CStr(ExtraID - 1), CStr(strExtraPartNo))
        Else
            '
        End If

        If Len(CStr(intPartQty)) Then
            strExtraQTY = CStr(intPartQty)
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtExtraQty:" & CStr(ExtraID - 1), CStr(strExtraQTY))
        ElseIf Len(strExtraQTY) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtExtraQty:" & CStr(ExtraID - 1), CStr(strExtraQTY))
        Else
            '
        End If

        'TOTALS is nothing ?????

        If Len(strDeliveryRef) > 0 And AddNewRow Then

            dtLastSaved = Now()
            ExtraFields = "DeliveryDate,DeliveryReference,FrameRowNumber,PartNo,Qty,ShortOrExtra,Start_TAGID,END_TAGID,TotalRows"
            ExtraFields = ExtraFields & ",UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"
            'ExtraFields = "Total_Rows,Start_TAGID,END_TAGID,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"

            'Getting error - cannot convert STRING to type LONG ???????????

            ExtraValues = CDate(SaveDeliveryDate) & "," & strDeliveryRef & "," & NewFrameRowNumber
            ExtraValues = ExtraValues & "," & PartNo & "," & intPartQty & "," & "EXTRA" & "," & LowerTag & "," & UpperTAG & "," & ControlTotalRows
            ExtraValues = ExtraValues & "," & strUsername & "," & strEMPNO & "," & strName
            ExtraValues = ExtraValues & "," & dtLastSaved

            If Len(ExtraRecID) > 0 Then
                SaveCriteria = "ID = " & ExtraRecID
            Else
                SaveCriteria = ""
            End If

            If Len(ErrMessage) > 0 Then
                MsgBox("Error;" & ErrMessage)
                Exit Function
            End If

            'ExtraSavedOK = InsertUpdateMyRecord(ExtraLoadedOK, frmMainGIForm.myConnString, ControlDBTable, ExtraFields, ExtraValues, ErrMessage, SaveCriteria, ExcludeFields)
            ExtraSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, ExtraLoadedOK, "", ControlDBTable, ExtraFields, ExtraValues,
                        SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            If Not ExtraSavedOK Then
                frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, ErrMessage, "ExtraSavedOK: InsertExtraParts()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
            Else
                'MsgBox("OK INSERTED Extra record")
                frmMainGIForm.txtMessages.Text = "OK INSERTED Extra record"
            End If
            'INSERT UPPER AND LOWER TAGS HERE TOO:
            If ExtraSavedOK Then
                ExtraFields = "TotalRows"
                ExtraValues = ControlTotalRows
                SaveCriteria = "DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34) & " AND ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
                ExtraSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "", ControlDBTable, ExtraFields, ExtraValues,
                            SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            End If

        End If

        InsertExtraParts2 = TotalRows - 1

    End Function

    Function InsertShortParts(AddNewRow As Boolean, strDeliveryDate As String, strDeliveryRef As String, strASNNo As String, ShortFrameControl As ScrollableControl,
                         ByVal FrameRowNumber As Long,
                         Optional ByVal TimeStartTAG As Long = 1000,
                         Optional ByVal LowerTag As Long = 1001,
                         Optional ByVal ControlsPerRow As Long = 3,
                         Optional ByVal StartFieldIndex As Long = 5,
                         Optional ByVal PartNo As String = "",
                         Optional ByVal intPartQty As Integer = 0,
                         Optional ByRef StartTABIndex As Long = 1,
                         Optional ByVal ResetFieldIndex As Boolean = True,
                         Optional ByVal FormName As String = "GI_TIMESHEET_") As Long
        Dim ShortID As Long = 0
        Dim UpperTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim StartBtnTAG As Long = 0
        Dim HighestShortTag As Long = 0
        Dim TotalControlsInFrame As Long = 0
        Dim tempTotals As clsTotals
        Dim ControlDBTable As String
        Dim dtDeliveryDate As DateTime
        Dim ShortTotalRows As Long
        Dim searchtext = strDeliveryRef
        Dim SearchField As String = "DeliveryReference"
        Dim FieldType As String = "STRING"
        Dim ReturnField As String = "ID"
        Dim ReturnValue As String = ""
        Dim SearchCriteria As String = ""
        Dim SortFields As String = ""
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String = ""
        Dim strShortTotalRows As String
        Dim lngShortTotalRows As Long
        Dim ShortLoadedOK As Boolean = False
        Dim ShortSavedOK As Boolean = False
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim TotalRowsFromQuery As Long
        Dim NewFrameRowNumber As Long
        Dim strShortPartNo As String
        Dim strShortQTY As String
        Dim ShortRecID As String
        Dim ControlTotalRows As Long
        Dim dtLastSaved As DateTime
        Dim ShortFields As String
        Dim ShortValues As String
        Dim SaveCriteria As String
        Dim SaveDeliveryDate As String
        Dim strUsername As String
        Dim strEMPNO As String
        Dim strName As String
        Dim strControlKey As String
        Dim tempControl As clsControls
        Dim ControlTabIndex As Long
        Dim FieldIndex As Long
        Dim ExtractTotals As clsTotals
        Dim TotalRows As Long
        Dim TotalRowCriteria As String

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:
        'THIS IS THE PROC THAT GETS CALLED FROM THE MAIN CONTROL PANEL !!!

        InsertShortParts = 0

        LowerTag = 1001
        UpperTAG = 0
        'TotalRows = 0
        StartBtnTAG = 0
        HighestShortTag = 0
        TimeStartTAG = 1000
        ControlTabIndex = StartTABIndex
        If Len(strDeliveryDate) = 0 Then
            MsgBox("Delivery Date Not Supplied")
            Exit Function
        End If

        If ResetFieldIndex Then
            FieldIndex = 4
        Else
            FieldIndex = StartFieldIndex
        End If

        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        ControlDBTable = "tblshortsandextraparts"
        SaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")

        If Len(strDeliveryRef) > 0 Then
            searchtext = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
            SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & FrameRowNumber
            SortFields = ""
            Reversed = False
            ExcludeFields = ""
            AllValues = Nothing
            AllFields = Nothing
        Else
            SaveDeliveryDate = CDate("1970-01-01 00:00:00")
        End If

        ShortRecID = ""
        TotalRowCriteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
        ShortTotalRows = Get_TotalRows("tblShortsAndExtraParts", strDeliveryRef, TotalRowCriteria)

        If dic_Totals(strDeliveryDate & "_" & strDeliveryRef) IsNot Nothing Then
            ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If ExtractTotals IsNot Nothing Then
                If AddNewRow Then
                    TotalRows = ShortTotalRows + 1
                Else
                    TotalRows = ShortTotalRows
                End If
                ExtractTotals.Total_ShortParts = TotalRows
                HighestShortTag = ExtractTotals.HighestShortTAGID
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = ExtractTotals
            Else
                'LOADED FROM DB:
                TotalRows = 1 'reset - default to 1.
                HighestShortTag = LowerTag
            End If
        End If

        NewFrameRowNumber = TotalRows

        strControlKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(1)
        tempControl = dic_Controls(strControlKey)
        strUsername = ""
        strName = ""
        strEMPNO = ""
        If tempControl IsNot Nothing Then
            strUsername = tempControl.ControlUpdatedByUsername
            strName = tempControl.ControlUpdatedByName
            strEMPNO = tempControl.ControlUpdatedByEmpNo
        End If

        ShortLoadedOK = Find_myQuery(frmMainGIForm.myConnString, ControlDBTable, SearchField, searchtext, FieldType, ReturnField, ShortRecID,
                                         AllValues, AllFields, SearchCriteria, SortFields, Reversed, ErrMessage, "=", TotalRowsFromQuery)

        strShortTotalRows = ""
        strShortPartNo = ""
        strShortQTY = ""
        If ShortLoadedOK Then
            'A SHORT PART record exists - get total rows and get the PART NO and QTY:
            strShortTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
            strShortPartNo = GetMYValuebyFieldname(AllValues, AllFields, "PartNo")
            strShortQTY = GetMYValuebyFieldname(AllValues, AllFields, "Qty")
            HighestShortTag = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")

            If IsNumeric(strShortTotalRows) Then
                lngShortTotalRows = CLng(strShortTotalRows) 'TotalRows from DB
            Else
                lngShortTotalRows = 0
            End If

        End If


        'when totalrows = 0 then shortid will have to be 1 as it is the row number.
        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTag
            ShortID = 1
            HighestShortTag = 1001
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTag
            ShortID = FrameRowNumber
        End If
        TotalControlsInFrame = Math.Abs(TotalRows * ControlsPerRow)
        ControlTotalRows = TotalRows

        AddNewShorts(TotalRows, ShortFrameControl, ShortID, HighestShortTag, StartBtnTAG,
                            strDeliveryDate, strDeliveryRef, strASNNo, LowerTag, UpperTAG, TimeStartTAG, ShortAndExtraFieldnames,
                        TotalRows, FieldIndex, False, ErrMessage, 0, PartNo, intPartQty, StartTABIndex)
        tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        If tempTotals IsNot Nothing Then
            tempTotals.HighestShortTAGID = HighestShortTag
            tempTotals.HighestShortTABIndex = StartTABIndex
        End If
        'frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef) = StartBtnTAG
        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalShorts", CStr(TotalRows - 1))
        If Len(CStr(FrameRowNumber)) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtShortRow:" & CStr(ShortID - 1), CStr(FrameRowNumber))
        End If
        If Len(PartNo) > 0 Then
            strShortPartNo = PartNo
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtShortPartNo:" & CStr(ShortID - 1), CStr(strShortPartNo))
        ElseIf Len(strShortPartNo) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtShortPartNo:" & CStr(ShortID - 1), CStr(strShortPartNo))
        Else
            '
        End If

        If Len(CStr(intPartQty)) Then
            strShortQTY = CStr(intPartQty)
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtShortQty:" & CStr(ShortID - 1), CStr(strShortQTY))
        ElseIf Len(strShortQTY) > 0 Then
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtShortQty:" & CStr(ShortID - 1), CStr(strShortQTY))
        Else
            '
        End If

        'TOTALS is nothing ?????

        If Len(strDeliveryRef) > 0 And AddNewRow Then

            dtLastSaved = Now()
            ShortFields = "DeliveryDate,DeliveryReference,FrameRowNumber,PartNo,Qty,ShortOrExtra,Start_TAGID,END_TAGID,TotalRows"
            ShortFields = ShortFields & ",UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"
            'ShortFields = "Total_Rows,Start_TAGID,END_TAGID,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved"

            'Getting error - cannot convert STRING to type LONG ???????????

            ShortValues = CDate(SaveDeliveryDate) & "," & strDeliveryRef & "," & NewFrameRowNumber
            ShortValues = ShortValues & "," & PartNo & "," & intPartQty & "," & "SHORT" & "," & LowerTag & "," & UpperTAG & "," & ControlTotalRows
            ShortValues = ShortValues & "," & strUsername & "," & strEMPNO & "," & strName
            ShortValues = ShortValues & "," & dtLastSaved

            If Len(ShortRecID) > 0 Then
                SaveCriteria = "ID = " & ShortRecID
            Else
                SaveCriteria = ""
            End If

            If Len(ErrMessage) > 0 Then
                MsgBox("Error;" & ErrMessage)
                Exit Function
            End If

            'ShortSavedOK = InsertUpdateMyRecord(ShortLoadedOK, frmMainGIForm.myConnString, ControlDBTable, ShortFields, ShortValues, ErrMessage, SaveCriteria, ExcludeFields)
            ShortSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, ShortLoadedOK, "", ControlDBTable, ShortFields, ShortValues,
                            SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            If Not ShortSavedOK Then
                frmMainGIForm.logger.LogError("GI_ERRORS_1_1.log", Application.StartupPath, ErrMessage, "InsertShortParts() SAVING: ", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
            Else
                'MsgBox("OK INSERTED SHORT record")
                frmMainGIForm.txtMessages.Text = "OK INSERTED SHORT record"
            End If

            If ShortSavedOK Then
                ShortFields = "TotalRows"
                ShortValues = ControlTotalRows
                SaveCriteria = "DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34) & " AND ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
                ShortSavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "", ControlDBTable, ShortFields, ShortValues,
                            SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            End If

        End If

        InsertShortParts = TotalRows - 1

    End Function


    Sub AddNewExtra(ByRef TotalOps As Long, FrameScrollControl As ScrollableControl, ByRef OpID As Long, ByRef TagID As Long, ByRef btnTAGID As Long,
                         strDeliveryDate As String, strDeliveryRef As String, ASN As String, LowerTAG As Long, UpperTAG As Long, ByVal TimeTAGStart As Long,
                         Fieldnames As String, TotalRows As Long, Optional StartFieldIndex As Long = 3, Optional DontIncrementOP As Boolean = False,
                         Optional ByRef ErrMessage As String = "", Optional ByRef NewIndex As Long = 0,
                         Optional ByVal PartNo As String = "", Optional ByVal PartQty As String = "", Optional ByRef StartTABIndex As Long = 1,
                         Optional ByVal TotalOpTEXTControlsPerRow As Long = 3)
        Dim RowGap As Long
        Dim TopPos As Integer
        Dim StartTopPos As Integer
        Dim ScrollBarHeight As Integer
        Dim ComboArray() As String
        Dim TimeTAGID As Long
        Dim ControlText As String
        Dim ControlType As String
        Dim ControlTAG As String
        Dim ControlDate As Date
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlDeliveryDate As Date
        Dim ControlDeliveryRef As String
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlRowNumber As Long
        Dim ControlTotalRows As Long
        Dim MakeVisible As Boolean
        Dim ControlBackColor As Integer
        Dim ControlForeColor As Integer
        Dim ControlLeftMargin As Boolean
        Dim ControlFieldname As String
        Dim FieldnameArr() As String
        Dim LoadFieldsOK As Boolean
        Dim ControlFieldsTable As String
        Dim SearchCriteria As String
        Dim ControlDBTable As String
        Dim ControlFontName As String
        Dim ControlFontSize As Integer
        Dim ControlStyle As String
        Dim FrameRowNumberField As String = ""
        Dim IDX As Long
        Dim RowIDX As Long
        Dim Firstname As String
        Dim Lastname As String
        Dim FullName As String
        Dim strActivity As String
        Dim ControlTABIndex As Long

        '************************* check FrameScrollControl - wrong Frame passed.

        ReDim FieldnameArr(1)
        ControlTABIndex = StartTABIndex
        ControlDBTable = "tblshortsandextraparts"
        If Len(Fieldnames) > 0 Then
            'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        Else
            ReDim FieldnameArr(20)
            Fieldnames = GetMyFields(ControlDBTable, frmMainGIForm.myConnString, ErrMessage)
            Beep()
            FieldnameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        End If
        FrameRowNumberField = "FrameRowNumber"
        StartTopPos = 2
        RowGap = 28
        If OpID < 2 Then
            TopPos = StartTopPos
        Else
            TopPos = ((OpID - 1) * RowGap) + StartTopPos
        End If

        TagID = Get_StartTAG(TimeTAGStart, OpID, TotalOpTEXTControlsPerRow, 1)

        ControlFontName = "Cambria"

        If IsDate(strDeliveryDate) Then
            ControlDeliveryDate = CDate(strDeliveryDate)
        Else
            ControlDeliveryDate = CDate("01/01/1970")
        End If
        ControlDeliveryRef = strDeliveryRef
        strDeliveryDate = ControlDeliveryDate.ToString("dd/MM/yyyy")

        MakeVisible = True
        ControlType = "TEXTBOX"
        ControlText = CStr(OpID)
        ControlLeft = 6
        ControlTop = TopPos
        ControlHeight = 23
        ControlWidth = 35
        ControlFontSize = 11
        ControlStyle = "BOLD"
        ControlTAG = CStr(TagID)
        ControlDate = Now()
        ControlDeliveryRef = strDeliveryRef
        ControlASN = ASN
        ControlOBJCount = OpID
        ControlStartTAG = CStr(LowerTAG)
        ControlEndTAG = CStr(UpperTAG)
        ControlTotalRows = TotalRows ' need to know lowerTAG and number of fields in frame_Operatives.
        'BackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlForeColor = RGB(255, 245, 60) 'yellow text
        ControlLeftMargin = False
        ControlFieldname = FrameRowNumberField
        ControlRowNumber = OpID
        ControlTotalRows = TotalOps
        ComboArray = Nothing

        Call AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtExtraRow:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", True)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        ControlLeft = 48
        ControlTop = TopPos
        ControlWidth = 120
        ControlHeight = 23
        ControlTAG = CStr(TagID)
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing

        ControlType = "TEXTBOX"
        ControlText = PartNo
        ControlStyle = "BOLD"
        ControlFontName = "Cambria"
        ControlFontSize = 11
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE, needs lt grey
        ControlFieldname = FieldnameArr(StartFieldIndex + 1)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtExtraPartNo:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", False)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        'ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
        ComboArray = Nothing
        strActivity = ""

        ControlType = "TEXTBOX"
        ControlText = PartQty
        ControlTAG = CStr(TagID)
        ControlStyle = "BOLD"
        ControlLeft = 174
        ControlWidth = 50
        ControlTop = TopPos
        ControlHeight = 23
        ControlFontName = "CAMBRIA"
        ControlFontSize = 11
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlFieldname = FieldnameArr(StartFieldIndex + 2)

        NewIndex = AddNewControl(True, FrameScrollControl, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtExtraQty:" & CStr(OpID), ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "", False)
        TagID = TagID + 1
        ControlTABIndex = ControlTABIndex + 1

        If DontIncrementOP = False Then
            OpID = OpID + 1
        End If

        FrameScrollControl.AutoScroll = True

        TotalOps = OpID
        StartTABIndex = ControlTABIndex

    End Sub

    Sub InsertExtraParts(strDeliveryDate As String, strDeliveryRef As String, strASNNo As String, OpFrameControl As ScrollableControl,
                         Optional ByRef TotalRows As Long = 0,
                         Optional ByVal TimeStartTAG As Long = 2000, Optional ByVal LowerTag As Long = 2001,
                         Optional ByVal ControlsPerRow As Long = 3, Optional ByVal StartFieldIndex As Long = 5,
                         Optional ByVal PartNo As String = "", Optional ByVal PartQty As String = "")
        Dim ExtraID As Long = 0
        Dim UpperTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim StartBtnTAG As Long = 0
        Dim HighestExtraTag As Long = 0
        Dim TotalControlsInFrame As Long = 0

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:

        LowerTag = 2001
        UpperTAG = 0
        TotalRows = 0
        StartBtnTAG = 0
        HighestExtraTag = 0



        If Len(ErrMessage) > 0 Then
            MsgBox("Error while Getting FieldNames")
            Exit Sub
        End If



    End Sub

    Sub DeleteControls(strDeliveryDate As String, strDeliveryRef As String, FrameScrollControl As ScrollableControl, TotaltxtControlsPerRow As Long,
                        Optional ByRef HighestTag As Long = 0, Optional ByVal TotalRows As Long = 0, Optional LowestTag As Long = 0, Optional ByRef ErrMEssage As String = "")
        Dim IDX As Long
        Dim ctrls() As Control
        Dim ctrlNames() As String
        Dim varKey As String = ""
        Dim varItem As New clsControls
        Dim ctrlIDX As Integer

        If HighestTag = 0 Then
            If TotalRows > 0 Then
                If LowestTag > 0 Then
                    HighestTag = (TotalRows * TotaltxtControlsPerRow) + (LowestTag - 1)
                Else
                    ErrMEssage = "LowerTag must be passed"
                    Exit Sub
                End If
            End If
        End If
        'IT SEEMS only one row has been recorded in the script dictionary - Dic_Controls ?????????
        ctrlIDX = 1
        ReDim ctrlNames(1)
        ReDim ctrls(1)
        'Remove the highest TAg downto highest tag - TotaltxtControlsPerRow
        If HighestTag > 0 Then
            'Put together the Key for the highest tag inside a loop. Gather together the control names into an array of the last row of controls.
            For IDX = HighestTag To HighestTag - TotaltxtControlsPerRow Step -1
                varKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(IDX)
                If dic_Controls.Exists(varKey) Then
                    varItem = dic_Controls.Item(varKey)
                    ctrlNames(ctrlIDX) = varItem.ControlName
                    ReDim Preserve ctrlNames(UBound(ctrlNames) + 1)
                    ctrlIDX = ctrlIDX + 1
                    'OK but we may not need the control name - yes we do - used as the key to remove the item.
                    'do we need the array of actual controls - as they can be assigned from the TheControl property of the dic_Controls() collection.
                    FrameScrollControl.Controls.Remove(varItem.TheControl)
                    dic_Controls.Remove(varKey)
                End If
            Next
            HighestTag = ((TotalRows - 1) * TotaltxtControlsPerRow) + (LowestTag - 1)
        End If

        'Retrieve the highest tag and then the control name of the highest tag downto control name of highest tag - Total_Controls_Per_Row_With_Tag

    End Sub

    Sub DeleteOpsRow(strDeliveryDate As String, strDeliveryRef As String,
                        FrameScrollControl As ScrollableControl, ByVal RowNumToDelete As Long, TotaltxtControlsPerRow As Long, Optional LowerTAG As Long = 43)
        Dim OpCtrl As Control
        Dim TagNumber As String
        Dim TagNo As Long
        Dim Answer As Integer
        Dim DeleteRecordAnswer As Integer
        Dim DeleteRow As Boolean = True
        Dim DeleteControls() As Control
        Dim strDeleteControls() As String
        Dim varKey As String = ""
        Dim ctrlIDX As Long = 0
        Dim TotalControls As Long = 0
        Dim ControlCount As Long = 0
        Dim ColonPos As Integer = 0
        Dim OpID As String = ""
        Dim SearchText As String
        Dim SearchField As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim SearchCriteria As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String
        Dim OpTotalRows As Long
        Dim LoadedOK As Boolean = False
        Dim AllFields() As String
        Dim AllValues() As Object
        Dim OpRecID As String
        Dim TableName As String = "tblOperatives"
        Dim ErrMessage As String
        Dim MarkForDeletion As Boolean = False
        Dim TempControl As clsControls

        ReDim DeleteControls(1)
        ReDim strDeleteControls(1)
        'DeleteControls = FrameScrollControl.Controls.Find("comOperativeName" & CStr(RowNumToDelete), True)
        'DeleteControls = FrameScrollControl.Controls.Find("comOperativeActivity" & CStr(RowNumToDelete), True)
        'For Each OpCtrl In DeleteControls
        ' If UCase(OpCtrl.Name) = UCase("comOperativeName" & CStr(RowNumToDelete)) Then
        'If Len(OpCtrl.Name) > 0 Then
        'Answer = MsgBox("Control contains a name - Are You Sure ?", vbYesNo, "WARNING - Contains Name !")
        'If Answer = vbYes Then
        'DeleteRow = True
        'Else
        'DeleteRow = False
        'End If
        'End If
        'End If
        'Next
        If Len(strDeliveryDate) = 0 Then
            MsgBox("No Delivery Date")
            Exit Sub
        End If
        If Len(strDeliveryRef) = 0 Then
            MsgBox("No Delivery Ref")
            Exit Sub
        End If

        SearchText = strDeliveryRef
        SearchField = "DeliveryReference"
        FieldType = "STRING"
        ReturnField = "ID"
        ReturnValue = ""
        SearchCriteria = "FrameRowNumber = " & RowNumToDelete
        SortFields = ""
        Reversed = False
        ExcludeFields = ""
        OpTotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
        LoadedOK = Find_myQuery(frmMainGIForm.myConnString, TableName, SearchField, SearchText, FieldType, ReturnField,
                                    OpRecID, AllValues, AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        For Each OpCtrl In FrameScrollControl.Controls
            ControlCount = ControlCount + 1
            ColonPos = InStrRev(OpCtrl.Name, ":", -1, CompareMethod.Text)
            OpID = Mid(OpCtrl.Name, ColonPos + 1)

            If IsNumeric(OpID) And CLng(OpID) = RowNumToDelete Then
                'If InStr(OpCtrl.Name, CStr(RowNumToDelete)) > 0 Then 'This now seems to work - removes all the controls in the row ???
                TagNumber = OpCtrl.Tag
                varKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(TagNumber)
                TempControl = dic_Controls(varKey)
                strDeleteControls(ctrlIDX) = varKey
                ReDim Preserve strDeleteControls(UBound(strDeleteControls) + 1)
                DeleteControls(ctrlIDX) = OpCtrl
                ReDim Preserve DeleteControls(UBound(DeleteControls) + 1)
                ctrlIDX = ctrlIDX + 1
                If UCase(OpCtrl.Name) = UCase("comOperativeName:" & CStr(RowNumToDelete)) Then
                    If Len(OpCtrl.Name) > 0 Then
                        If Not UCase(OpCtrl.Text) = "SELECT EMPLOYEE" Then 'check database to see if exists already in table ?
                            'Check if record exists in database table: if so get its Record ID at same time:
                            'SearchText = CDate(strDeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")

                            If LoadedOK Then
                                'This record exists in the database !
                                If UCase(frmMainGIForm.myAccessRights) = "ADMIN" Or UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
                                    'OK ISSUE WARNING and confirmation:
                                    DeleteRecordAnswer = MsgBox("WARNING - Corresponding Database record EXISTS - Are You Sure ?", vbYesNo, "WARNING - Contains Name !")
                                    MarkForDeletion = False
                                    If DeleteRecordAnswer = vbYes Then
                                        'DELETE the record:
                                        DeleteRow = True
                                        'Need primary key of DB record to delete:
                                        'Total_Operatives in the tblLabourHours and tblDeliveryInfo needs reducing by 1
                                        Call DeleteMyRecord(TableName, frmMainGIForm.myConnString, "ID = " & OpRecID, ErrMessage)
                                    Else
                                        DeleteRow = False
                                    End If
                                Else
                                    'Insufficient Access Rights to delete:
                                    DeleteRecordAnswer = MsgBox("Control contains a name - You dont have permission to Delete - mark record for deletion ?", vbYesNo, "WARNING - Contains Name !")
                                    MarkForDeletion = True
                                End If
                            Else
                                'No corresponding record exists - ok to delete row of controls:
                                Answer = MsgBox("Control contains a name - Are You Sure ?", vbYesNo, "WARNING - Contains Name !")
                                If Answer = vbYes Then
                                    DeleteRow = True
                                Else
                                    DeleteRow = False
                                End If
                            End If
                        Else
                            'Control contains Select Employee: so ok to remove corresponding record if exists
                            If LoadedOK Then
                                DeleteRow = True
                                'Need primary key of DB record to delete:
                                'Total_Operatives in the tblLabourHours and tblDeliveryInfo needs reducing by 1
                                Call DeleteMyRecord(TableName, frmMainGIForm.myConnString, "ID = " & OpRecID, ErrMessage)

                            End If
                        End If
                    End If
                End If
            End If
        Next
        'Need to somehow gather the buttons also - eg btn10 and btn11 on the first row.
        TotalControls = UBound(DeleteControls)
        If DeleteRow Then
            'Delete all the controls in the array:
            For ctrlIDX = TotalControls To 0 Step -1
                OpCtrl = DeleteControls(ctrlIDX)
                If OpCtrl IsNot Nothing Then
                    FrameScrollControl.Controls.Remove(OpCtrl)
                    If dic_Controls.exists(strDeleteControls(ctrlIDX)) Then
                        If strDeleteControls(ctrlIDX) IsNot Nothing Then
                            dic_Controls.remove(strDeleteControls(ctrlIDX))
                        End If
                    End If
                End If
            Next
        End If

    End Sub


    Sub DeletePartsRow(strDeliveryDate As String, strDeliveryRef As String, strCheckFieldName As String,
                        FrameScrollControl As ScrollableControl, ByVal RowNumToDelete As Long, TotaltxtControlsPerRow As Long)
        Dim OpCtrl As Control
        Dim TagNumber As String
        Dim Answer As Integer
        Dim DeleteRow As Boolean = True
        Dim DeleteControls() As Control
        Dim strDeleteControls() As String
        Dim varKey As String = ""
        Dim ctrlIDX As Long = 0
        Dim TotalControls As Long = 0
        Dim ControlCount As Long = 0
        Dim ColonPos As Integer = 0
        Dim OpID As String = ""
        Dim Message As String = ""

        Try
            ReDim DeleteControls(1)
            ReDim strDeleteControls(1)
            'DeleteControls = FrameScrollControl.Controls.Find("comOperativeName" & CStr(RowNumToDelete), True)
            'DeleteControls = FrameScrollControl.Controls.Find("comOperativeActivity" & CStr(RowNumToDelete), True)
            'For Each OpCtrl In DeleteControls
            ' If UCase(OpCtrl.Name) = UCase("comOperativeName" & CStr(RowNumToDelete)) Then
            'If Len(OpCtrl.Name) > 0 Then
            'Answer = MsgBox("Control contains a name - Are You Sure ?", vbYesNo, "WARNING - Contains Name !")
            'If Answer = vbYes Then
            'DeleteRow = True
            'Else
            'DeleteRow = False
            'End If
            'End If
            'End If
            'Next
            '********************************************** check - giving error *********************************************** 14/08/2018
            For Each OpCtrl In FrameScrollControl.Controls
                ControlCount = ControlCount + 1
                ColonPos = InStrRev(OpCtrl.Name, ":", -1, CompareMethod.Text)
                OpID = Mid(OpCtrl.Name, ColonPos + 1)
                If IsNumeric(OpID) And CLng(OpID) = RowNumToDelete Then
                    'If InStr(OpCtrl.Name, CStr(RowNumToDelete)) > 0 Then 'This now seems to work - removes all the controls in the row ???
                    TagNumber = OpCtrl.Tag
                    varKey = strDeliveryDate & "_" & strDeliveryRef & "_" & CStr(TagNumber)
                    strDeleteControls(ctrlIDX) = varKey
                    ReDim Preserve strDeleteControls(UBound(strDeleteControls) + 1)
                    DeleteControls(ctrlIDX) = OpCtrl
                    ReDim Preserve DeleteControls(UBound(DeleteControls) + 1)
                    ctrlIDX = ctrlIDX + 1
                    If UCase(OpCtrl.Name) = UCase(strCheckFieldName & CStr(RowNumToDelete)) Then
                        If Len(OpCtrl.Name) > 0 Then
                            If Not UCase(OpCtrl.Text) = "" Then
                                Answer = MsgBox("Control contains a Part - Are You Sure ?", vbYesNo, "WARNING - Contains Part !")
                                If Answer = vbYes Then
                                    DeleteRow = True
                                Else
                                    DeleteRow = False
                                End If
                            Else

                            End If
                        End If
                    End If
                End If
            Next
            'Need to somehow gather the buttons also - eg btn10 and btn11 on the first row.
            TotalControls = UBound(DeleteControls)
            If DeleteRow Then
                'Delete all the controls in the array:
                For ctrlIDX = TotalControls To 0 Step -1
                    OpCtrl = DeleteControls(ctrlIDX)
                    If OpCtrl IsNot Nothing Then
                        FrameScrollControl.Controls.Remove(OpCtrl)
                        dic_Controls.remove(strDeleteControls(ctrlIDX))
                    End If
                Next
            End If

        Catch ex As Exception
            Message = "Exception Error In Test Asset-Register DB Login: " & ex.Message
            frmMainGIForm.logger.LogError("GI_Error_v1_3.log", Application.StartupPath, Message, "TestLogin()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("GI_Message_v1_1.log", Application.StartupPath, Message, "Something wrong with Deleting Parts" & CStr(DeleteRow))
        End Try

    End Sub


    Function SearchFrame(FrameScrollControl As ScrollableControl, ControlName As String) As Control
        Dim FrameCtrl As Control
        Dim ControlCount As Long = 0

        SearchFrame = Nothing
        For Each FrameCtrl In FrameScrollControl.Controls 'not passing anything = nothing. 0 controls passed.
            ControlCount = ControlCount + 1
            If UCase(FrameCtrl.Name) = UCase(ControlName) Then
                Return FrameCtrl
            End If

        Next
    End Function

    Function Get_TotalHours_From_DB(strDeliveryDate As String, strDeliveryRef As String,
                                    Optional FormName As String = "GI_TIMESHEET_") As Double
        Dim DBTable As String = "tblLabourHours"
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim SearchCriteria As String = ""
        Dim SortFields As String = "ID"
        Dim Reversed As Boolean = False
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllValues() As Object = Nothing
        Dim AllFields() As String = Nothing
        Dim ErrMessage As String = ""
        Dim LoadedOK As Boolean = False
        Dim FieldType As String = "STRING"
        Dim dblTotalHours As Double = 0.0F
        Dim strTotalHours As String = ""
        Dim FoundCtrl As Control = Nothing
        Dim ChildCtrl As Control = Nothing
        Dim TagNumber As Long
        Dim TotalHoursSPAN As TimeSpan
        Dim strFinalTotalHours As String
        Dim Message As String

        Try
            Get_TotalHours_From_DB = 0.0F
            'Total_Labour_Hours
            ReturnField = "ID"
            SearchField = "DeliveryReference"
            SearchText = strDeliveryRef

            LoadedOK = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            strFinalTotalHours = ""

            If LoadedOK Then
                strTotalHours = GetMYValuebyFieldname(AllValues, AllFields, "Total_Labour_Hours")
                If IsNumeric(strTotalHours) Then
                    dblTotalHours = CDbl(strTotalHours)
                End If
                TotalHoursSPAN = TimeSpan.FromHours(dblTotalHours)
                If TotalHoursSPAN.Days > 0 Then
                    'strFinalTotalHours = CStr(TotalHoursSPAN.Days) & "d" 'NEEDS TO BE CONVERTED TO HOURS !!!
                End If
                'strFinalTotalHours = strFinalTotalHours & " " & CStr(TotalHoursSPAN.Hours) & "h " & CStr(TotalHoursSPAN.Minutes) & "m " & CStr(TotalHoursSPAN.Seconds) & "s "
                strFinalTotalHours = CStr(TotalHoursSPAN.Hours) & "h " & CStr(TotalHoursSPAN.Minutes) & "m " & CStr(TotalHoursSPAN.Seconds) & "s "
                'Find correct field to put value into now.
                'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                TagNumber = 24
                FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(TagNumber), ChildCtrl)
                If FoundCtrl IsNot Nothing Then
                    'put value into control:
                    frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strFinalTotalHours)
                End If
            End If

        Catch ex As Exception
            Message = "Exception Error In Get_TotalHours_From_DB: " & ex.Message
            frmMainGIForm.logger.LogError("GI_Error_v1_3.log", Application.StartupPath, Message, "Get_TotalHours_From_DB()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("GI_Message_v1_1.log", Application.StartupPath, Message, "Something wrong with Get_TotalHours_From_DB")
        End Try

        Get_TotalHours_From_DB = dblTotalHours

    End Function

    Sub PopulateControls(ByRef TheDeliveryDate As String, ByRef TheDeliveryRef As String,
                         Optional ByRef ASNNo As String = "",
                         Optional FormName As String = "",
                         Optional ByRef ViewGrid As DataGridView = Nothing)
        Dim DBTables() As String
        Dim LowerTAG() As Long
        Dim UpperTAG() As Long
        Dim TAbleIDX As Long
        Dim LoadedOK() As Boolean
        Dim FieldIDX As Long
        Dim StartIDX() As Long
        Dim EndIDX() As Long
        Dim FrameRows() As Long
        Dim TotalControlsInFrame() As Long
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim SearchCriteria As String = ""
        Dim SortFields As String = ""
        Dim Reversed As Boolean = False
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllValues() As Object = Nothing
        Dim AllFields() As String = Nothing
        Dim ErrMessage As String = ""
        Dim dtDeliveryDate As DateTime
        Dim strDeliveryDate As String
        Dim FieldType As String = "STRING"
        Dim strSupplier As String
        Dim strDueTime As String
        Dim strSHIFT As String
        Dim strOrigin As String
        Dim strResult As String = ""
        Dim Fieldname As String = ""
        Dim TagNumber As Long = 0
        Dim NewTAGNumber As Long = 0
        Dim TimeTagNumber As Long = 0
        Dim TagIndex As Long = 0
        Dim FoundCtrl As Control = Nothing
        Dim ChildCtrl As Control = Nothing
        Dim FieldNameArr() As String
        Dim Fieldnames() As String
        Dim FrameRowNumberField As String
        Dim ControlDBTable As String
        Dim ControlFontName As String
        Dim ControlDeliveryDate As DateTime
        Dim strDeliveryRef As String
        Dim ControlDeliveryRef As String
        Dim MakeVisible As Boolean = True
        Dim ControlType As String = "TEXTBOX"
        Dim ControlText As String = ""
        Dim ControlLeft As Integer = 7
        Dim ControlTop As Integer = 1
        Dim ControlHeight As Integer = 23
        Dim ControlWidth As Integer = 35
        Dim ControlFontSize As Integer = 11
        Dim ControlFontStyle As String
        Dim ControlTAG As String
        Dim ControlDate As String = Now()
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlTotalRows As Long ' need to know lowerTAG and number of fields in frame_Operatives.
        Dim OpStartTAG As Long = 43
        Dim OpFinishTAG As Long = 48
        Dim OpTotalRows As Long = 0
        Dim OpStartTime As String = "00:00:00"
        Dim OpEndTime As String = "00:00:00"
        Dim OpCriteria As String = ""
        Dim OpGridFields As String
        Dim ShortTotalRows As Long = 0
        Dim ExtraTotalRows As Long = 0
        Dim ShortStartTAG As Long = 0
        Dim ExtraStartTAG As Long = 0
        Dim ShortOrExtra As String = ""
        Dim ControlBACKCOLOUR As Integer = RGB(0, 112, 192) 'BLUE, RGB(240, 248, 255) 'ALICEBLUE
        Dim ControlForeColor As Integer = RGB(255, 245, 60) 'yellow text
        Dim ControlLeftMargin As Boolean = False
        Dim ControlFieldname As String = ""
        Dim ControlRowNumber As Long
        Dim ComboArray() As String = Nothing
        Dim ControlName As String = ""
        Dim varKeyAsn As String = ""
        Dim varKeyRef As String = ""
        Dim Frame_Extra_Parts As ScrollableControl
        Dim Frame_FLMDetails As ScrollableControl
        Dim Frame_InboundSchedule As ScrollableControl
        Dim Frame_OperationalInput As ScrollableControl
        Dim Frame_Operatives As ScrollableControl
        Dim Frame_OpsShortsAndExtras As ScrollableControl
        Dim Frame_Short_Parts As ScrollableControl
        Dim Frame_SupplierCompliance As ScrollableControl
        Dim ControlFrameName As String = ""
        Dim NewIndex As Long = 0
        Dim ListArray() As Object = Nothing
        Dim ControlTextAlign As String = ""
        Dim NewTextBox As New TextBox
        Dim LastRow As Long
        Dim dtStartTime As DateTime
        Dim dtEndTime As DateTime
        Dim FLMName As String = ""
        Dim OpName As String = ""
        Dim OpActivity As String = ""
        Dim OpComments As String = ""
        Dim ShortPartNo As String = ""
        Dim ShortQty As String = ""
        Dim intShortQty As Long
        Dim ExtraPartNo As String = ""
        Dim intExtraQty As Long
        Dim ExtraQty As String = ""
        Dim SavedOK() As Boolean
        Dim LabourFields As String
        Dim LabourFieldValues As String
        Dim SaveCriteria As String = ""
        Dim ExcludeFields As String = ""
        Dim SCFieldsArr() As String = Nothing
        Dim tempTotals As New clsTotals
        Dim ExtractTotals As New clsTotals
        Dim strSaveDeliveryDate As String
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim ControlTABIndex As Long
        Dim StartControlTABIndex As Long
        Dim tempControl As clsControls
        Dim UpdatedByName As String
        Dim UpdatedByUsername As String
        Dim UpdatedByEmpNo As String
        Dim ControlPrimaryKey As Long
        Dim SCPrimaryKey As Long
        Dim OpPrimaryKey As Long
        Dim ShortPrimaryKey As Long
        Dim ExtraPrimaryKey As Long
        Dim ReturnFromSql As Object
        Dim strSQL As String
        Dim OpLoadedOK As Boolean
        Dim OpControl As Control
        Dim ChangeProperty As String
        Dim testCollection As Object
        Dim RecordLocked As Boolean = False
        Dim DisableFLMStartButton As Boolean = False
        Dim DisableFLMFinishButton As Boolean = False
        Dim DisableOpStartButton As Boolean = False
        Dim DisableOpFinishButton As Boolean = False
        Dim DisableOpStartEntry As Boolean = True
        Dim DisableOpFinishEntry As Boolean = True
        Dim dblTotalHours As Double = 0.0F
        Dim LockUsername As String
        Dim LockFullName As String
        Dim LockEmpNo As String
        Dim LockSTATUS As String
        Dim CompleteSTATUS As String
        Dim CTRLIDX As String
        Dim NewFormTitle As String
        Dim Message As String
        Dim PopulateGridMessage As String
        Dim NumRows As Long
        Dim GridSQL As String
        Dim OpFieldnames As String
        Dim OpSortFields As String

        'Static CountPopulate As Integer
        CountPopulate = CountPopulate + 1

        If Len(FormName) = 0 Then
            FormName = CPFormName
        End If

        Try

            CTRLIDX = FormID
            'MsgBox("CONTROL PANEL IDX= " & CTRLIDX)
            Frame_Extra_Parts = GetFrameControl(CPFormName & FormID, "Frame_Extra_Parts")
            Frame_FLMDetails = GetFrameControl(FormName & FormID, "Frame_FLMDetails")
            Frame_InboundSchedule = GetFrameControl(FormName & FormID, "Frame_InboundSchedule")
            Frame_OperationalInput = GetFrameControl(FormName & FormID, "Frame_OperationalInput")
            Frame_Operatives = GetFrameControl(FormName & FormID, "Frame_Operatives")
            Frame_OpsShortsAndExtras = GetFrameControl(FormName & FormID, "Frame_OpsShortsAndExtras")
            Frame_Short_Parts = GetFrameControl(FormName & FormID, "Frame_Short_Parts")
            Frame_SupplierCompliance = GetFrameControl(FormName & FormID, "Frame_SupplierCompliance")

            Frame_Operatives.Controls.Clear()
            Frame_Short_Parts.Controls.Clear()
            Frame_Extra_Parts.Controls.Clear()
            Frame_FLMDetails.Controls.Clear()


            ReDim DBTables(6)
            ReDim LowerTAG(6)
            ReDim UpperTAG(6)
            ReDim LoadedOK(6)
            ReDim StartIDX(6)
            ReDim EndIDX(6)
            ReDim FrameRows(6)
            ReDim TotalControlsInFrame(6)
            ReDim Fieldnames(6)
            ReDim SavedOK(6)

            '**************************************************************************************************************************************
            'HOW DO WE TEST IF ANOTHER USER IS ALREADY EDITING THIS REFERENCE ?????????????????????
            '**************************************************************************************************************************************
            ' - Need to put some kind of LOCK in place here. record username / EmpNo until finished editing - ie SAVED
            '- maybe a LOCK table in the database ?
            ' - tblLOCKUsers
            'ID, DeliveryDate, DeliveryReferenceBeingEdited, Username, UserEmpNo,Current DateTime, LOCKED = YES / NO , LastEdited
            '- number of records in this table will depend on number of users currently logged on.
            'Check the LOCK table as soon as user has selected a reference. HALT further processing until table has been checked and is CLEAR to EDIT.
            'Purge table the following day. (shows total References that have been edited / worked on.

            'TRY
            TheDeliveryDate = ""
            If Len(ASNNo) > 0 Then
                'lookup in DB - tblDeliveryInfo the corresponding Delivery Reference:
                SearchText = ASNNo
                ControlASN = ASNNo
                SearchField = "ASN_Number"
                FieldType = "STRING"
                ReturnField = "ID"
                ReturnValue = ""
                SortFields = "DeliveryDate"
                Reversed = True
                DBTables(1) = "tblDeliveryInfo"
                'comASNNo
                'frmMainGIForm.InsertValueIntoForm(ControlPanelFormName & FormID, "comASNNo", ControlASN)
                'SearchCriteria = "ASN_Number = " & Chr(34) & ASNNo & Chr(34)
                LoadedOK(1) = Find_myQuery(frmMainGIForm.myConnString, DBTables(1), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                If LoadedOK(1) Then
                    strDeliveryRef = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryReference")
                    strDeliveryDate = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryDate")
                    strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
                    TheDeliveryDate = strDeliveryDate
                    TheDeliveryRef = strDeliveryRef
                Else
                    MsgBox("Could not find corresponding Delivery Reference")
                    Exit Sub
                End If
            Else
                strDeliveryRef = ""
            End If

            If Len(strDeliveryRef) = 0 Then
                strDeliveryRef = TheDeliveryRef
            Else
                'Check if REFERENCE is Locked:
                'Has it got LOCKED against the reference in tblDeliveryInfo ???


                'LOCK THIS REFERENCE:
                'SET LOCkeD against This Reference in tblDeliveryInfo - use UPDATE method


                'UNLOCK reference after the SAVE button has been clicked ?
                ' - No - need a flag to see IF it has been saved -  if not warn user before switching to another Reference.
                'Need to check for LOCKED against reference - If there IS and there is NO reference in the comReference field then 
                ' - UNLOCK the reference.

            End If


            If Len(TheDeliveryDate) = 0 Then
                dtDeliveryDate = CDate("1970-01-01 00:00:00")
                strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")

                If Len(strDeliveryRef) > 0 Then
                    SearchText = strDeliveryRef
                    ControlASN = ASNNo
                    SearchField = "DeliveryReference"
                    FieldType = "STRING"
                    ReturnField = "ID"
                    ReturnValue = ""
                    SortFields = "DeliveryDate"
                    Reversed = True
                    DBTables(1) = "tblDeliveryInfo"
                    'comASNNo
                    'frmMainGIForm.InsertValueIntoForm(ControlPanelFormName & FormID, "comASNNo", ControlASN)
                    'SearchCriteria = "ASN_Number = " & Chr(34) & ASNNo & Chr(34)
                    LoadedOK(1) = Find_myQuery(frmMainGIForm.myConnString, DBTables(1), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                           AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                    If LoadedOK(1) Then
                        strDeliveryDate = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryDate")
                        strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
                        TheDeliveryDate = strDeliveryDate
                        'Check if record locked BY ANOTHER USER:
                        LockSTATUS = ""
                        LockUsername = ""
                        LockEmpNo = ""
                        LockFullName = ""
                        If Get_Reference_Lock_Info(False, strDeliveryDate, strDeliveryRef, CompleteSTATUS, LockSTATUS, LockUsername, LockEmpNo,
                                                   LockFullName) Then
                            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "lblTitle_OperationalInput", "Operational Input, CLERK: " & LockUsername)
                            If UCase(frmMainGIForm.myEmpNo) = UCase(LockEmpNo) Then
                                'OK to EDIT - belongs to THIS USER LOGGED IN:
                            ElseIf Len(LockUsername) = 0 Then
                                'Either the username did not save correctly - or its ok as not been edited yet.

                            Else
                                'Someone ELSE IS CURRENTLY EDITING THIS REFERENCE:
                                'could be SAME USER with the other login eg dgoss2 or dgoss3 etc.
                                'Also check if the username is still in the tblSessions or tblUsersonline table:
                                '   This checks if they are still online - only if the LOGIN DATE is LESS THAN 12 HOURS (past SHIFT TIME).
                                If UCase(frmMainGIForm.myAccessRights) = "ADMIN" Or UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
                                    'OK Permission to see ALL references - whether in progress or not. Use at own risk.
                                Else
                                    If UCase(CompleteSTATUS) = "IN PROGRESS" Then
                                        frmMainGIForm.txtMessages.Text = "Being Edited By: " & LockUsername & " EmpNO: " & LockEmpNo
                                        MsgBox("Already LOCKED by : " & LockUsername & " EmpNO: " & LockEmpNo)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtSelectDeliveryDate", strDeliveryDate)
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtDeliveryDate", strDeliveryDate)

                    Else
                        MsgBox("Could not find corresponding Delivery Date")
                        Exit Sub
                    End If


                End If

            Else
                dtDeliveryDate = CDate(TheDeliveryDate)

                strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
            End If

            'frmMainGIForm.dic_Totals = CreateObject("Scripting.Dictionary") 'NOT part of ANYMORE in frmMainGIForm - 28/08/2018
            'frmMainGIForm.dic_Totals.comparemode = vbTextCompare

            'rmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) = 1

            tempTotals.Total_Operatives = 0
            tempTotals.Total_ShortParts = 0
            tempTotals.Total_ExtraParts = 0
            tempTotals.HighestOpTAGID = 43
            tempTotals.HighestOpBtnTAGID = 10
            tempTotals.HighestShortTAGID = 1001
            tempTotals.HighestExtraTAGID = 2001
            tempTotals.HighestOpTabIndex = 1
            tempTotals.HighestShortTABIndex = 1
            tempTotals.HighestExtraTABIndex = 1
            tempTotals.TotalOpHours = 0
            tempTotals.TotalFLMHours = 0

            'If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
            'End If

            DBTables(1) = "tblDeliveryInfo"
            DBTables(2) = "tblSupplierCompliance"
            DBTables(3) = "tblLabourHours"
            DBTables(4) = "tblOperatives"
            DBTables(5) = "tblShortsAndExtraParts"
            DBTables(6) = "tblShortsAndExtraParts"

            TotalControlsInFrame(1) = 1
            TotalControlsInFrame(2) = 1
            TotalControlsInFrame(3) = 1
            TotalControlsInFrame(4) = 6
            TotalControlsInFrame(5) = 3
            TotalControlsInFrame(6) = 3
            LowerTAG(1) = 1
            UpperTAG(1) = 30
            LowerTAG(2) = 801
            UpperTAG(2) = 807
            LowerTAG(3) = 40
            UpperTAG(3) = 42
            LowerTAG(4) = 43
            UpperTAG(4) = 48
            LowerTAG(5) = 1001
            UpperTAG(5) = 1003
            LowerTAG(6) = 2001
            UpperTAG(6) = 2003
            If Len(strDeliveryDate) = 0 Then
                MsgBox("No Delivery Date found")
                Exit Sub
            End If
            SearchCriteria = ""
            ReturnField = "ID"
            dtDeliveryDate = CDate(strDeliveryDate)
            strSaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")
            'strDeliveryRef = TheDeliveryRef
            ControlDeliveryRef = TheDeliveryRef
            SearchText = strSaveDeliveryDate
            'SearchCriteria = "([DeliveryDate] = #" & strDeliveryDate & "#) And ([DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34) & ")"
            If Len(TheDeliveryRef) > 0 Then
                'SearchCriteria = "DeliveryReference = " & Chr(34) & TheDeliveryRef & Chr(34)
            End If

            SortFields = ""

            'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
            'FieldNameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 1, False, False, False, "_", False)
            ReDim FieldNameArr(20)
            'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)

            Fieldnames(1) = GetMyFields(DBTables(1), frmMainGIForm.myConnString, ErrMessage)
            Fieldnames(2) = GetMyFields(DBTables(2), frmMainGIForm.myConnString, ErrMessage)
            Fieldnames(3) = GetMyFields(DBTables(3), frmMainGIForm.myConnString, ErrMessage)
            Fieldnames(4) = GetMyFields(DBTables(4), frmMainGIForm.myConnString, ErrMessage)
            Fieldnames(5) = GetMyFields(DBTables(5), frmMainGIForm.myConnString, ErrMessage)
            Fieldnames(6) = GetMyFields(DBTables(6), frmMainGIForm.myConnString, ErrMessage)
            FieldNameArr = DanG_DB_Tools.strToStringArray(Fieldnames(1), ",", 1, False, False, False, "_", False)

            'ID is at pos=1
            FrameRowNumberField = "FrameRowNumber"

            ControlDBTable = DBTables(1)
            ControlFontName = "Cambria"

            If IsDate(strDeliveryDate) Then
                ControlDeliveryDate = CDate(strDeliveryDate)
            Else
                'MsgBox ("Need to pass proper delivery date")
                'Exit Sub
                ControlDeliveryDate = CDate("01/01/1970")
            End If
            ControlDeliveryRef = strDeliveryRef

            MakeVisible = True
            ControlType = "TEXTBOX"
            'ControlText = CStr(OpID)
            ControlLeft = 7
            ControlTop = 1
            ControlHeight = 23
            ControlWidth = 35
            ControlFontSize = 11
            ControlFontStyle = "BOLD"
            'ControlTAG = CStr(TagID)
            ControlDate = Now()
            ControlASN = ""
            ControlOBJCount = 1
            ControlStartTAG = ""
            ControlEndTAG = ""
            ReturnValue = ""
            ControlTotalRows = 0 ' need to know lowerTAG and number of fields in frame_Operatives.
            'BackColor = RGB(240, 248, 255) 'ALICEBLUE
            'ControlBackColor = RGB(0, 112, 192) 'BLUE
            'ControlForeColor = RGB(255, 245, 60) 'yellow text
            'ControlLeftMargin = False
            'ControlFieldname = FrameRowNumberField
            'ControlRowNumber = OpID
            'ControlTotalRows = TotalOps
            'ComboArray = Nothing

            'OperativeFrame = GetFrameControl(ControlPanelFormName & FormID, "Frame_Operatives")
            'TotalFrameRows = GetTotalFrameRows(OperativeFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(4), VTTAG)
            'Beep()

            Clear_Entry_Controls(FormName & FormID, 1, 40)
            Clear_Entry_Controls(FormName & FormID, 801, 813)
            SearchText = strDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            Reversed = True
            SortFields = "DeliveryDate"
            LoadedOK(1) = Find_myQuery(frmMainGIForm.myConnString, DBTables(1), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(1) Then
                ControlDeliveryRef = strDeliveryRef
                strSupplier = AllValues(3)
                ControlASN = AllValues(4)
                If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
                    ControlPrimaryKey = CLng(ReturnValue)
                End If
                If Len(ControlASN) > 0 Then
                    varKeyAsn = strDeliveryDate & "_" & ControlASN & "_" & "1"
                End If
                varKeyRef = strDeliveryDate & "_" & ControlDeliveryRef & "_" & "1"
                'Call UpdateCollection(dic_Controls, "", AllValues(1), AllValues(1), AllValues(5), AllValues(2), "1", "VALUE")
                'Call UpdateCollection(dic_Controls, "", Frame_InboundSchedule.Name, AllValues(1), AllValues(5), AllValues(2), "4", "FRAMENAME")
                strDeliveryDate = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryDate")
                strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
                If Len(strDeliveryRef) > 0 Then
                    NewFormTitle = "Goods In Control Panel_" & "REF:" & strDeliveryRef
                ElseIf Len(ASNNo) > 0 Then
                    NewFormTitle = "Goods In Control Panel_" & "ASN:" & ASNNo
                Else
                    NewFormTitle = "Goods In Control Panel_" & Now().ToString("dd/MM/yyyy HH:mm:ss")
                End If
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtDeliveryDate", strDeliveryDate)
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtSelectDeliveryDate", strDeliveryDate)
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtDeliveryRef", AllValues(2), Nothing, NewFormTitle)
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtSupplier", strSupplier)
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtASNNum", ControlASN)
                'SET MAIN FORM TITLE:
                frmMainGIForm.Text = frmMainGIForm.myUsername & ", REF: " & strDeliveryRef

                FieldIDX = 1
                TagNumber = LowerTAG(1)
                Do While FieldIDX < 32
                    ControlFieldname = AllFields(FieldIDX)
                    strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                    'Find correct field to put value into now.
                    'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                    FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(TagNumber), ChildCtrl)
                    If FoundCtrl IsNot Nothing Then
                        'put value into control:
                        If IsDate(strResult) And FieldIDX = 1 Then
                            strResult = CDate(strResult).ToString("dd/MM/yyyy")
                        End If
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strResult)
                        'Insert into Dic_Controls:
                        varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber
                        ControlName = FoundCtrl.Name
                        ControlTAG = CStr(TagNumber)
                        ControlTABIndex = FoundCtrl.TabIndex
                        ControlText = strResult
                        ControlType = "TEXTBOX" 'Assumption that all tag controls are just textboxes - not dropdowns either ?
                        ControlLeft = FoundCtrl.Left
                        ControlTop = FoundCtrl.Top
                        ControlWidth = FoundCtrl.Width
                        ControlHeight = FoundCtrl.Height
                        ControlOBJCount = TagNumber
                        ControlStartTAG = "1"
                        ControlEndTAG = "27"
                        ControlRowNumber = 1
                        ControlTotalRows = 1
                        ListArray = Nothing
                        ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                        ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                        ControlFontName = FoundCtrl.Font.Name
                        ControlFontSize = FoundCtrl.Font.Size
                        If FoundCtrl.Font.Bold Then
                            ControlFontStyle = "BOLD"
                        Else
                            ControlFontStyle = "NORMAL"
                        End If
                        If FoundCtrl.Font.Italic Then
                            ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                        End If
                        If FoundCtrl.Font.Underline Then
                            ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                        End If
                        If FoundCtrl.Font.Strikeout Then
                            ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                        End If
                        If TypeOf (FoundCtrl) Is TextBox Then
                            NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                            If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                ControlTextAlign = "LEFT"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                ControlTextAlign = "CENTER"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                ControlTextAlign = "RIGHT"
                            Else
                                ControlTextAlign = ""
                            End If
                        End If
                        NewIndex = AddNewControl(False, Frame_InboundSchedule, DBTables(1), ControlFieldname, "ID", FoundCtrl, ControlName,
                                             ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                                             ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                             frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                                             ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                             True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                                             ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False,
                                             CDate("01/01/1970"), CDate("01/01/1970"), ControlPrimaryKey)
                    End If
                    FieldIDX = FieldIDX + 1
                    TagNumber = TagNumber + 1
                Loop
                'strOrigin = GetMYValuebyFieldname(AllValues, AllFields, "Origin")
            End If

            Call Get_TotalHours_From_DB(strDeliveryDate, strDeliveryRef)
            'Problem - may not have all the values loaded into the fields correctly here ???

            FieldIDX = 3
            SCFieldsArr = strToStringArray(SCFieldnames, ",", 0, False, False, False, "_", False, 34, 39)
            varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & "1"
            tempControl = dic_Controls(varKeyRef)
            If tempControl Is Nothing Then Exit Sub
            UpdatedByEmpNo = tempControl.ControlUpdatedByEmpNo
            UpdatedByName = tempControl.ControlUpdatedByName
            UpdatedByUsername = tempControl.ControlUpdatedByUsername
            LoadedOK(2) = Find_myQuery(frmMainGIForm.myConnString, DBTables(2), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            'If LoadedOK(2) Then
            'tblsupplierCompliance:
            If Not LoadedOK(2) Then
                dtLastSaved = Now()
                strLastSaved = dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss")
                SaveCriteria = ""
                LabourFields = "DeliveryDate,DeliveryReference,UpdatedByEmpNo,UpdatedByUsername,UpdatedByName,LastSaved"
                'LabourFieldValues = Chr(39) & strSaveDeliveryDate & Chr(39) & "," & Chr(34) & strDeliveryRef & Chr(34) & "," & Chr(39) & strLastSaved & Chr(39)
                LabourFieldValues = CDate(strDeliveryDate) & "," & strDeliveryRef
                LabourFieldValues = LabourFieldValues & "," & UpdatedByEmpNo & "," & UpdatedByUsername & "," & UpdatedByName
                LabourFieldValues = LabourFieldValues & "," & dtLastSaved
                ExcludeFields = ""
                'SavedOK(2) = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, DBTables(2), LabourFields, LabourFieldValues, ErrMessage, SaveCriteria, ExcludeFields)
                SavedOK(2) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, "", DBTables(2), LabourFields, LabourFieldValues, SaveCriteria, ExcludeFields, ErrMessage)
            End If
            Clear_Entry_Controls(FormName & FormID, 801, 812)
            TagNumber = LowerTAG(2)
            LoadedOK(2) = Find_myQuery(frmMainGIForm.myConnString, DBTables(2), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
                SCPrimaryKey = CLng(ReturnValue)
            End If
            Do While FieldIDX < UBound(SCFieldsArr) - 1
                ControlFieldname = SCFieldsArr(FieldIDX)
                If ControlFieldname Is Nothing Then
                    Continue Do
                End If
                If LoadedOK(2) Then
                    strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                Else
                    strResult = ""
                End If

                'Find correct field to put value into now.
                'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(TagNumber), ChildCtrl)
                If FoundCtrl IsNot Nothing Then
                    'put value into control:
                    frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strResult)
                    'Insert into Dic_Controls:
                    varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber
                    ControlName = FoundCtrl.Name
                    ControlTAG = CStr(TagNumber)
                    ControlTABIndex = FoundCtrl.TabIndex
                    ControlText = strResult
                    ControlType = "TEXTBOX"
                    ControlLeft = FoundCtrl.Left
                    ControlTop = FoundCtrl.Top
                    ControlWidth = FoundCtrl.Width
                    ControlHeight = FoundCtrl.Height
                    ControlOBJCount = TagNumber
                    ControlStartTAG = CStr(LowerTAG(2))
                    ControlEndTAG = CStr(UpperTAG(2))
                    ControlRowNumber = 1
                    ControlTotalRows = 1
                    ListArray = Nothing
                    ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                    ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                    ControlFontName = FoundCtrl.Font.Name
                    ControlFontSize = FoundCtrl.Font.Size
                    If FoundCtrl.Font.Bold Then
                        ControlFontStyle = "BOLD"
                    Else
                        ControlFontStyle = "NORMAL"
                    End If
                    If FoundCtrl.Font.Italic Then
                        ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                    End If
                    If FoundCtrl.Font.Underline Then
                        ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                    End If
                    If FoundCtrl.Font.Strikeout Then
                        ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                    End If
                    If TypeOf (FoundCtrl) Is TextBox Then
                        NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                        If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                            ControlTextAlign = "LEFT"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                            ControlTextAlign = "CENTER"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                            ControlTextAlign = "RIGHT"
                        Else
                            ControlTextAlign = ""
                        End If
                    End If
                    NewIndex = AddNewControl(False, Frame_SupplierCompliance, DBTables(2), ControlFieldname, "ID", FoundCtrl, ControlName,
                                                 ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                                                 ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                                                 ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                                 True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                                                 ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False,
                                                    CDate("1970-01-01"), CDate("1970-01-01"), SCPrimaryKey)
                End If 'FoundCtrl
                FieldIDX = FieldIDX + 1
                TagNumber = TagNumber + 1
            Loop

            '******************************************************************** CHECK FLM DETAILS and LOAD:

            AllValues = Nothing
            AllFields = Nothing
            SortFields = ""
            Reversed = False
            ErrMessage = ""
            SearchCriteria = ""
            SearchField = "DELIVERYREFERENCE"
            SearchText = strDeliveryRef
            FieldType = "String"
            ReturnField = "ID"
            ReturnValue = ""

            OpTotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
            ReturnValue = ""
            LoadedOK(3) = Find_myQuery(frmMainGIForm.myConnString, DBTables(3), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(3) Then
                'tblLAbourHours: - extract FLM info:
                FLMName = GetMYValuebyFieldname(AllValues, AllFields, "FLM_Name")
                dtStartTime = CDate(GetMYValuebyFieldname(AllValues, AllFields, "FLM_StartTime"))
                dtEndTime = CDate(GetMYValuebyFieldname(AllValues, AllFields, "FLM_FinishTime"))
                If dtStartTime > CDate("1970-01-01") Then
                    DisableFLMStartButton = True
                Else
                    DisableFLMStartButton = False
                End If
                If dtEndTime > CDate("1970-01-01") Then
                    DisableFLMStartButton = True
                Else
                    DisableFLMStartButton = False
                End If
                FieldIDX = 3
                TagNumber = LowerTAG(3)
                NewTAGNumber = TagNumber
                ControlType = "COMBOBOX"
                ControlTABIndex = 33 'START TAB INDEX - need to remove the employee textbox in following proc:
                Call InsertFLMButtons(LoadedOK(3), Frame_FLMDetails, strDeliveryDate, strDeliveryRef, ControlASN, "",
                                      FLMName, dtStartTime, dtEndTime, ControlTABIndex, DisableFLMStartButton, DisableFLMFinishButton)
                Do While FieldIDX < 7
                    NewTAGNumber = TagNumber
                    ControlFieldname = AllFields(FieldIDX)
                    strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                    If UCase(ControlFieldname) = UCase("FLM_Name") Then
                        FLMName = strResult
                    End If
                    If UCase(ControlFieldname) = UCase("FLM_StartTime") Then
                        dtStartTime = CDate(strResult)
                        strResult = dtStartTime.ToString("HH:mm:ss")

                    End If
                    If UCase(ControlFieldname) = UCase("FLM_FinishTime") Then
                        dtEndTime = CDate(strResult)
                        strResult = dtEndTime.ToString("HH:mm:ss")

                    End If
                    If UCase(ControlFieldname) = UCase("Labour_Comments") Then
                        OpComments = strResult
                    End If
                    'Find correct field to put value into now.
                    If FieldIDX > 3 And FieldIDX < 7 Then
                        NewTAGNumber = TagNumber + 400 'DANGEROUS - any more controls in this frame - use a sep variable.
                        ControlType = "TEXTBOX"
                    End If
                    'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                    FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(NewTAGNumber), ChildCtrl)
                    If FoundCtrl IsNot Nothing Then
                        'put value into control:
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strResult)
                        'Insert into Dic_Controls:
                        varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & NewTAGNumber 'WRONG DATE FORMAT - needs / not -
                        ControlName = FoundCtrl.Name
                        ControlTAG = CStr(NewTAGNumber)
                        ControlText = strResult

                        ControlLeft = FoundCtrl.Left
                        ControlTop = FoundCtrl.Top
                        ControlWidth = FoundCtrl.Width
                        ControlHeight = FoundCtrl.Height
                        ControlOBJCount = NewTAGNumber
                        'LOAD FROM DATABASE TABLE tblLAbourHours:
                        If GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID") > 0 Then
                            OpStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID")
                        Else
                            OpStartTAG = LowerTAG(4)
                        End If
                        If GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID") > 0 Then
                            OpFinishTAG = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")
                        Else
                            OpFinishTAG = UpperTAG(4)
                        End If

                        tempTotals.OpStartTAG = OpStartTAG
                        tempTotals.OpFinishTAG = OpFinishTAG
                        tempTotals.Total_Operatives = OpTotalRows

                        dtDeliveryDate = CDate(strDeliveryDate)
                        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
                        'The following adds an incorect property item to the dic_controls - due to the DeliveryDate being wrong format:
                        ' - Amended. check again.
                        dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

                        ControlStartTAG = CStr(OpStartTAG)
                        ControlEndTAG = CStr(OpFinishTAG)
                        ControlRowNumber = 0

                        'OpTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
                        If OpTotalRows > 0 Then
                            tempTotals.Total_Operatives = OpTotalRows
                            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
                        End If
                        ListArray = Nothing
                        ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                        ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                        ControlFontName = FoundCtrl.Font.Name
                        ControlFontSize = FoundCtrl.Font.Size
                        If FoundCtrl.Font.Bold Then
                            ControlFontStyle = "BOLD"
                        Else
                            ControlFontStyle = "NORMAL"
                        End If
                        If FoundCtrl.Font.Italic Then
                            ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                        End If
                        If FoundCtrl.Font.Underline Then
                            ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                        End If
                        If FoundCtrl.Font.Strikeout Then
                            ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                        End If
                        If TypeOf (FoundCtrl) Is TextBox Then
                            NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                            If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                ControlTextAlign = "LEFT"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                ControlTextAlign = "CENTER"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                ControlTextAlign = "RIGHT"
                            Else
                                ControlTextAlign = ""
                            End If
                        End If
                        If IsNumeric(ReturnValue) And Len(ReturnValue) > 0 Then
                            ControlPrimaryKey = CLng(ReturnValue)
                        End If
                        'set tempControl to clsControls then can use tempControl.ControlTag = "48", tempControl.ControlFontStyle = "BOLD" etc. and just pass tempControl
                        NewIndex = AddNewControl(False, Frame_FLMDetails, DBTables(3), ControlFieldname, "ID", FoundCtrl, ControlName,
                        ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                        ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                                True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                        ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False, dtStartTime, dtEndTime, ControlPrimaryKey)
                    End If
                    FieldIDX = FieldIDX + 1
                    TagNumber = TagNumber + 1
                Loop
            Else
                'NO FLM record so INSERT NEW FLM Details Record into tblLabourHours:
                ControlTABIndex = 33
                Call InsertFLMButtons(False, Frame_FLMDetails, strDeliveryDate, strDeliveryRef, ControlASN, "", FLMName, dtStartTime, dtEndTime, ControlTABIndex)

            End If

            '****************************************************************************************************************==================================
            'LOAD tblOperatives:
            '****************************************************************************************************************==================================

            AllValues = Nothing
            AllFields = Nothing
            SortFields = ""
            Reversed = False
            ErrMessage = ""
            SearchCriteria = ""
            SearchField = "DELIVERYREFERENCE"
            SearchText = strDeliveryRef
            FieldType = "String"
            ReturnField = "ID"
            ReturnValue = ""
            FieldIDX = 4
            PopulateGridMessage = ""

            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtOpName", ReturnValue, "", "Op_Name", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtOpActivity", ReturnValue, "", "Activity_Name", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "btnOp_StartTime", ReturnValue, "", "Op_StartTime", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtOp_StartTime", ReturnValue, "00:00", "Op_StartTime", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "btnOp_FinishTime", ReturnValue, "", "Op_FinishTime", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtOp_FinishTime", ReturnValue, "00:00", "Op_FinishTime", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtTotalTime", ReturnValue, "00:00", "OpComments", ControlASN)
            Call Populate_New_Controls(FormName, strDeliveryDate, strDeliveryRef, Frame_OpsShortsAndExtras, DBTables(4), "txtOp_StartTime2", ReturnValue, "00:00", "Op_StartTime", ControlASN)

            'CALL PopulateGrid() procedure - dgvOperatives with appropriate SQL script to select certain fields etc.
            'Do we still need TAG 43 to TAG 48 in clsControls / Dic_Controls() ???

            If Not IsNothing(ViewGrid) Then
                'OpFieldnames = "ID,Op_Name as 'Name',Activity_Name as 'Activity',Op_StartTime as 'Start Time',Op_FinishTime as 'Fin Time',OpComments as 'Comment'"
                OpFieldnames = "ID,Op_Name,Activity_Name,Op_StartTime,Op_FinishTime,OpComments"
                OpGridFields = "ID,Op_Name AS 'NAME',Activity_Name AS 'Activity',Op_StartTime AS 'Start',Op_FinishTime AS 'FINISH',OpComments AS 'TOTAL'"
                OpSortFields = "FrameRowNumber"
                OpCriteria = " WHERE DeliveryReference = " & Chr(39) & SearchText & Chr(39)
                'GridSQL = "SELECT " & OpFieldnames & " FROM " & DBTables(4) & OpCriteria & " ORDER BY " & OpSortFields
                GridSQL = "SELECT  " & OpGridFields & " FROM " & DBTables(4) & OpCriteria & " ORDER BY " & OpSortFields

                'Call PopulateMyDataSource(ViewGrid.DataSource, frmMainGIForm.myConnString, GridSQL, NumRows, PopulateGridMessage)
                Call PopulateMyDataSource(ViewGrid.DataSource, frmMainGIForm.myConnString, GridSQL, NumRows, PopulateGridMessage, ViewGrid, "10,20,200,30,30,30,30")
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalHours", CStr(NumRows))
                If Len(PopulateGridMessage) > 0 Then
                    'MsgBox("Populate GRID ERROR: " & PopulateGridMessage)
                    frmMainGIForm.txtMessages.Text = "NO OPERATIVES ENTERED"
                End If
            End If

            'Insert Short Parts Rows :
            SortFields = ""
            Reversed = False
            ErrMessage = ""
            ShortOrExtra = "SHORT"
            SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
            LoadedOK(5) = Find_myQuery(frmMainGIForm.myConnString, DBTables(5), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(5) Then
                If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
                    ShortPrimaryKey = CLng(ReturnValue)
                End If
                ControlRowNumber = 1
                ShortTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
                ControlStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID")
                ControlEndTAG = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")

                tempTotals.Total_ShortParts = ShortTotalRows
                tempTotals.HighestShortTAGID = ControlEndTAG
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
                tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                If Not tempTotals Is Nothing Then
                    StartControlTABIndex = tempTotals.HighestShortTABIndex
                Else
                    StartControlTABIndex = 1001
                End If
                ControlTABIndex = StartControlTABIndex

                If IsNumeric(ControlStartTAG) Then
                    TagNumber = CLng(ControlStartTAG)
                Else
                    TagNumber = 1001
                End If
                ControlStartTAG = CStr(TagNumber)
                ControlTotalRows = ShortTotalRows
                Do While ControlRowNumber <= ShortTotalRows
                    SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
                    SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & ControlRowNumber
                    LoadedOK(5) = Find_myQuery(frmMainGIForm.myConnString, DBTables(5), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                           AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                    If LoadedOK(5) Then
                        'tblshortsandextraparts: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                        FieldIDX = 4
                        ShortPartNo = GetMYValuebyFieldname(AllValues, AllFields, "PartNo")
                        ShortQty = GetMYValuebyFieldname(AllValues, AllFields, "Qty")
                        If IsNumeric(ShortQty) Then
                            intShortQty = CLng(ShortQty)
                        Else
                            intShortQty = 0
                        End If
                        InsertShortParts(False, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Short_Parts, ControlRowNumber, 1000,
                                         ControlStartTAG, 3, FieldIDX, ShortPartNo, intShortQty, ControlTABIndex)
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalShorts", CStr(ControlRowNumber))
                        'what about using TotalShorts ???
                    Else
                        'Cannot Find record:

                    End If 'If the current numbered row exists
                    ControlRowNumber = ControlRowNumber + 1
                Loop
                tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                tempTotals.HighestShortTABIndex = ControlTABIndex
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
            End If 'IF any rows exist in tblShortsAndExtraParts WHERE ShortORExtra = Short

            'INSERT EXTRA PARTS ROWS:
            SortFields = ""
            Reversed = False
            ErrMessage = ""
            ShortOrExtra = "EXTRA"
            SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
            LoadedOK(6) = Find_myQuery(frmMainGIForm.myConnString, DBTables(6), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(6) Then
                If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
                    ExtraPrimaryKey = CLng(ReturnValue)
                End If
                ControlRowNumber = 1
                ExtraTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
                ControlStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID")
                ControlEndTAG = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")

                tempTotals.Total_ExtraParts = ExtraTotalRows
                tempTotals.HighestExtraTAGID = ControlEndTAG
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

                tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                If Not tempTotals Is Nothing Then
                    StartControlTABIndex = tempTotals.HighestExtraTABIndex
                Else
                    StartControlTABIndex = 2001
                End If
                ControlTABIndex = StartControlTABIndex

                If IsNumeric(ControlStartTAG) Then
                    TagNumber = CLng(ControlStartTAG)
                Else
                    TagNumber = 2001
                End If
                ControlStartTAG = CStr(TagNumber)
                ControlTotalRows = ExtraTotalRows
                Do While ControlRowNumber <= ExtraTotalRows
                    SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
                    SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & ControlRowNumber
                    LoadedOK(6) = Find_myQuery(frmMainGIForm.myConnString, DBTables(6), SearchField, SearchText, FieldType, ReturnField, ReturnValue,
                                               AllValues, AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                    If LoadedOK(6) Then
                        'tblshortsandextraparts: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                        FieldIDX = 4
                        ExtraPartNo = GetMYValuebyFieldname(AllValues, AllFields, "PartNo")
                        ExtraQty = GetMYValuebyFieldname(AllValues, AllFields, "Qty")
                        If IsNumeric(ExtraQty) Then
                            intExtraQty = CLng(ExtraQty)
                        Else
                            intExtraQty = 0
                        End If

                        InsertExtraParts2(False, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Extra_Parts, ControlRowNumber, 2000,
                                          ControlStartTAG, 3, FieldIDX, ExtraPartNo, intExtraQty, ControlTABIndex,)
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalExtras", CStr(ControlRowNumber))

                    Else
                        'Not found record:

                    End If 'If the current numbered row exists
                    ControlRowNumber = ControlRowNumber + 1
                Loop
                tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                If Not tempTotals Is Nothing Then
                    tempTotals.HighestExtraTABIndex = ControlTABIndex
                Else
                    tempTotals.HighestExtraTABIndex = 2001
                End If
                dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
            End If 'IF any rows exist in tblShortsAndExtraParts WHERE ShortORExtra = Extra

            Call Check_For_Completion(strDeliveryDate, strDeliveryRef, False)
            frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtASNNum", ControlASN)

        Catch ex As Exception
            Message = "Exception Error In PopulateControls() " & ex.Message
            frmMainGIForm.txtMessages.Text = Message
            frmMainGIForm.logger.LogError("GI_Error_" & frmMainGIForm.myVersion & ".log", Application.StartupPath, Message, "PopulateControls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("GI_Message_" & frmMainGIForm.myVersion & ".log", Application.StartupPath, Message, "Something wrong with PopulateControls")
        End Try


    End Sub

    Sub Populate_New_Controls(FormName As String, strDeliveryDate As String, strDeliveryRef As String, Op_Frame As ScrollableControl, DBTable As String,
                              Optional ByVal controlName As String = "",
                              Optional ByVal strResult As String = "",
                              Optional ByVal strControlTAG As String = "",
                              Optional ByVal ControlFieldname As String = "",
                              Optional ByVal ControlASN As String = "",
                              Optional ByVal comboList() As String = Nothing)
        Dim FoundCtrl As Control
        Dim TagNumber As Long
        Dim ChildCtrl As Control
        Dim varKeyRef As String
        Dim ControlNames() As String
        Dim ControlTabIndex As Integer
        Dim ControlText As String
        Dim ControlType As String
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlOBJCount As Object
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlRowNumber As Long
        Dim ControlTotalRows As Long
        Dim ControlDeliveryRef As String
        Dim LowerTAG As Long
        Dim UpperTAG As Long
        Dim ListArray() As String
        Dim ControlBACKCOLOUR As Integer
        Dim ControlForeColor As Integer
        Dim ControlFontName As String
        Dim NewIndex As Long
        Dim ControlFontSize As Integer
        Dim ControlFontStyle As String
        Dim NewTextBox As TextBox
        Dim ControlTextAlign As String
        Dim ControlDate As DateTime
        Dim ControlDeliveryDate As DateTime
        Dim SCPrimaryKey As Long
        Dim Message As String

        Try
            ControlNames = {""}
            'controlName is favoured over ControlTAG first:
            FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, controlName, strControlTAG, ChildCtrl)
            If FoundCtrl IsNot Nothing Then
                'put value into control:
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strResult)
                'Insert into Dic_Controls:

                controlName = FoundCtrl.Name
                strControlTAG = FoundCtrl.Tag
                varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & strControlTAG
                ControlTabIndex = FoundCtrl.TabIndex
                ControlText = strResult
                ControlType = FoundCtrl.GetType.ToString
                'ControlType = "TEXTBOX"
                ControlLeft = FoundCtrl.Left
                ControlTop = FoundCtrl.Top
                ControlWidth = FoundCtrl.Width
                ControlHeight = FoundCtrl.Height
                If IsNumeric(strControlTAG) Then
                    ControlOBJCount = strControlTAG
                Else
                    ControlOBJCount = 0
                End If
                ControlStartTAG = CStr(LowerTAG)
                ControlEndTAG = CStr(UpperTAG)
                ControlRowNumber = 1
                ControlTotalRows = 1
                ListArray = Nothing
                If comboList IsNot Nothing Then
                    ListArray = comboList
                End If
                ControlDate = CDate(strDeliveryDate)
                ControlDeliveryDate = CDate(strDeliveryDate)
                ControlDeliveryRef = strDeliveryRef
                ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                ControlFontName = FoundCtrl.Font.Name
                ControlFontSize = FoundCtrl.Font.Size
                If FoundCtrl.Font.Bold Then
                    ControlFontStyle = "BOLD"
                Else
                    ControlFontStyle = "NORMAL"
                End If
                If FoundCtrl.Font.Italic Then
                    ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                End If
                If FoundCtrl.Font.Underline Then
                    ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                End If
                If FoundCtrl.Font.Strikeout Then
                    ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                End If
                ControlTextAlign = "LEFT"
                If TypeOf (FoundCtrl) Is TextBox Then
                    NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                    If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                        ControlTextAlign = "LEFT"
                    ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                        ControlTextAlign = "CENTER"
                    ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                        ControlTextAlign = "RIGHT"
                    Else
                        ControlTextAlign = ""
                    End If
                End If
                NewIndex = AddNewControl(False, Op_Frame, DBTable, ControlFieldname, "ID", FoundCtrl, controlName,
                                                 ControlText, ControlType, strControlTAG, ControlTabIndex, ControlDate, ControlLeft, ControlTop,
                                                 ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                                                 ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                                 True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                                                 ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False,
                                                    CDate("1970-01-01"), CDate("1970-01-01"), SCPrimaryKey)
            End If 'FoundCtrl
        Catch ex As Exception
            Message = "Exception Error In PopulateControls() " & ex.Message
            frmMainGIForm.txtMessages.Text = Message
            frmMainGIForm.logger.LogError("GI_Error_v1_3.log", Application.StartupPath, Message, "Populate_New_Controls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("GI_Message_v1_3.log", Application.StartupPath, Message, "Something wrong with Populate_New_Controls")
        End Try
    End Sub

    Sub Populate_Dynamic_Operative_Controls(ByVal LowerTag() As Long, ByVal strDeliveryDate As String, ByVal strDeliveryRef As String, ByVal DBTables() As String,
                                            ByVal SearchField As String, ByVal SearchText As String, ByVal FieldType As String, ByVal ReturnField As String,
                                            ByVal ReturnValue As String, ByVal SortFields As String, ByVal Reversed As Boolean, ByVal FieldIDX As Long,
                                            ByVal ControlASN As String, ByVal Frame_Operatives As ScrollableControl, ByVal TotalControlsInFrame() As Long,
                                            ByVal FormName As String, ByVal ControlDate As DateTime, ByVal ControlDeliveryDate As DateTime, ByVal ControlDeliveryRef As String,
                                            ByRef OpTotalRows As Long, ByRef ExtractTotals As clsTotals, ByRef LoadedOK() As Long, ByRef OpLoadedOK As Boolean)
        Dim ControlRowNumber As Long
        Dim OpStartTAG As Long
        Dim OpFinishTAG As Long
        Dim TagNumber As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlTotalRows As Long
        Dim StartControlTABIndex As Long
        Dim ControlTABIndex As Long
        Dim SearchCriteria As String
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim ErrMessage As String
        Dim OpPrimaryKey As Long
        Dim TimeTagNumber As Long
        Dim ControlFieldname As String
        Dim OpName As String
        Dim OpActivity As String
        Dim OpComments As String
        Dim strResult As String
        Dim DisableOpStartButton As Boolean
        Dim DisableOpFinishButton As Boolean
        Dim dtStartTime As DateTime
        Dim dtEndTime As DateTime
        Dim DisableOpStartEntry As Boolean
        Dim DisableOpFinishEntry As Boolean
        Dim ControlTAG As String
        Dim ControlType As String
        Dim OpStartTime As String
        Dim OpEndTime As String
        Dim FoundCtrl As Control
        Dim ChildCtrl As Control
        Dim ControlText As String
        Dim varKeyRef As String
        Dim ControlName As String
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlOBJCount As Long
        Dim ListArray As Object
        Dim ControlBACKCOLOUR As Integer
        Dim ControlForeColor As Integer
        Dim ControlFontName As String
        Dim ControlFontSize As Single
        Dim ControlFontStyle As String
        Dim NewTextBox As TextBox
        Dim ControlTextAlign As String
        Dim NewIndex As Long
        Dim OpControl As Control
        Dim ChangeProperty As String
        Dim testCollection As Object

        ControlRowNumber = 1
        If OpStartTAG = 0 Then
            OpStartTAG = LowerTag(4)
        End If
        TagNumber = OpStartTAG
        ControlStartTAG = CStr(TagNumber)
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        ControlTotalRows = OpTotalRows
        StartControlTABIndex = ExtractTotals.HighestOpTabIndex
        ControlTABIndex = StartControlTABIndex

        'Need to populate ControlPrimaryKey for EACH Operative row - in each txtOpRow: :

        Do While ControlRowNumber <= OpTotalRows
            SearchCriteria = "FrameRowNumber = " & ControlRowNumber

            LoadedOK(4) = Find_myQuery(frmMainGIForm.myConnString, DBTables(4), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(4) Then
                'tblOperatives: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
                    OpPrimaryKey = CLng(ReturnValue)
                End If
                TimeTagNumber = 0
                ControlFieldname = AllFields(FieldIDX)
                OpName = GetMYValuebyFieldname(AllValues, AllFields, "Op_Name")
                OpActivity = GetMYValuebyFieldname(AllValues, AllFields, "Activity_Name")
                OpComments = GetMYValuebyFieldname(AllValues, AllFields, "OpComments")
                strResult = GetMYValuebyFieldname(AllValues, AllFields, "Op_StartTime")
                ErrMessage = "1) CONTROL FIELDNAME = " & ControlFieldname
                frmMainGIForm.logger.LogError("OpProperties", Application.StartupPath, ErrMessage, "Populate_Dynamic_Operative_Controls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED IN:" & frmMainGIForm.myUsername)
                ErrMessage = "2) CONTROL VALUE = " & strResult
                frmMainGIForm.logger.LogError("OpProperties", Application.StartupPath, ErrMessage, "Populate_Dynamic_Operative_Controls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED IN:" & frmMainGIForm.myUsername)
                DisableOpStartButton = False
                DisableOpFinishButton = False
                If IsDate(strResult) Then
                    dtStartTime = CDate(strResult)
                    If dtStartTime > CDate("1970-01-01") Then
                        DisableOpStartButton = True
                        If UCase(frmMainGIForm.myAccessRights) = "ADMIN" Or
                                UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
                            DisableOpStartEntry = False
                        End If
                    End If
                Else
                    dtStartTime = CDate("1970-01-01 01:00:00")
                End If
                strResult = GetMYValuebyFieldname(AllValues, AllFields, "Op_FinishTime")
                If IsDate(strResult) Then
                    dtEndTime = CDate(strResult)
                    If dtEndTime > CDate("1970-01-01") Then
                        DisableOpFinishButton = True
                        If UCase(frmMainGIForm.myAccessRights) = "ADMIN" Or
                                UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
                            DisableOpFinishEntry = False
                        End If
                    End If
                Else
                    dtEndTime = CDate("1970-01-01 01:00:00")
                End If
                'LastRow is the ONLY place used here - but the operative record seems to load in still ok ????
                'it does trigger off a few odd values - like totalrows = -1 etc.
                'LastRow ??????? should be TotalRows here :

                'IF OpNAME or OpActivity etc is BLANK - dont SAVE IT !

                InsertOperatives(False, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Operatives, ControlRowNumber, 400, OpStartTAG, TotalControlsInFrame(4),
                                 FieldIDX, OpName, OpActivity, dtStartTime, dtEndTime, OpComments, ControlTABIndex, True,
                                 DisableOpStartButton, DisableOpFinishButton, DisableOpStartEntry, DisableOpFinishEntry)

                Do While FieldIDX <= 9
                    'TagNumber = TagNumber + 1
                    TimeTagNumber = 0
                    ControlTAG = CStr(TagNumber)
                    ControlFieldname = AllFields(FieldIDX)
                    strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                    If UCase(ControlFieldname) = UCase("Op_StartTime") Then
                        TimeTagNumber = TagNumber + 400
                        ControlTAG = CStr(TimeTagNumber)
                        ControlType = "TEXTBOX"
                        dtStartTime = CDate(strResult)
                        strResult = dtStartTime.ToString("HH:mm:ss")
                        OpStartTime = strResult
                    End If
                    If UCase(ControlFieldname) = UCase("Op_FinishTime") Then
                        TimeTagNumber = TagNumber + 400
                        ControlTAG = CStr(TimeTagNumber)
                        ControlType = "TEXTBOX"
                        dtEndTime = CDate(strResult)
                        strResult = dtEndTime.ToString("HH:mm:ss")
                        OpEndTime = strResult
                    End If
                    If TimeTagNumber > 0 Then
                        FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(TimeTagNumber), ChildCtrl)
                    Else
                        FoundCtrl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(TagNumber), ChildCtrl)
                    End If
                    If FoundCtrl IsNot Nothing Then
                        frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundCtrl.Name, strResult)
                        strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                        'Find correct field to put value into now.
                        'ControlTAG = cstr(TagNumber)

                        If UCase(ControlFieldname) = UCase("FrameRowNumber") Then
                            ControlType = "TEXTBOX"
                        End If
                        If UCase(ControlFieldname) = UCase("Op_Name") Then
                            ControlType = "COMBOBOX"
                            If Len(strResult) > 0 Then
                                OpName = strResult
                            End If
                        End If
                        If UCase(ControlFieldname) = UCase("Activity_Name") Then
                            ControlType = "COMBOBOX"
                            If Len(strResult) > 0 Then
                                OpActivity = strResult

                            End If
                        End If
                        If UCase(ControlFieldname) = UCase("OpComments") Then
                            ControlType = "TEXTBOX"
                        End If

                        ControlText = strResult
                        'put value into control:

                        'Insert into Dic_Controls:

                        varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber 'not correct if Time is the field  - not adding 400.
                        ControlName = FoundCtrl.Name
                        'ControlTAG = CStr(TagNumber)
                        ControlText = strResult
                        ControlLeft = FoundCtrl.Left
                        ControlTop = FoundCtrl.Top
                        ControlWidth = FoundCtrl.Width
                        ControlHeight = FoundCtrl.Height
                        ControlOBJCount = TagNumber ' will record 46 rather than 446 for the time.
                        ControlStartTAG = CStr(OpStartTAG)
                        ControlEndTAG = CStr(OpFinishTAG)
                        'ControlStartTAG = CStr(OpStartTAG)
                        'ControlEndTAG = CStr(OpFinishTAG)
                        'OpTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
                        ListArray = Nothing
                        ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                        ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                        ControlFontName = FoundCtrl.Font.Name
                        ControlFontSize = FoundCtrl.Font.Size
                        If FoundCtrl.Font.Bold Then
                            ControlFontStyle = "BOLD"
                        Else
                            ControlFontStyle = "NORMAL"
                        End If
                        If FoundCtrl.Font.Italic Then
                            ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                        End If
                        If FoundCtrl.Font.Underline Then
                            ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                        End If
                        If FoundCtrl.Font.Strikeout Then
                            ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                        End If
                        If TypeOf (FoundCtrl) Is TextBox Then
                            NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                            If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                ControlTextAlign = "LEFT"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                ControlTextAlign = "CENTER"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                ControlTextAlign = "RIGHT"
                            Else
                                ControlTextAlign = ""
                            End If
                        End If

                        NewIndex = AddNewControl(False, Frame_Operatives, DBTables(4), ControlFieldname, "ID", FoundCtrl, ControlName,
                            ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                            ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                            ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                        True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                            ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False, dtStartTime, dtEndTime,
                                                 OpPrimaryKey)
                    End If 'FoundCtrl
                    FieldIDX = FieldIDX + 1
                    TagNumber = TagNumber + 1
                    ControlTABIndex = ControlTABIndex + 1

                Loop
                'TagNumber = TagNumber + 1
                'reset fieldidx back to 4 for NEXT iteration - due to StartFieldIDX becomming 10 and therefore issuing WRONG FIELDNAMES !!!
                'fieldIDX = 4
                frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalOps", CStr(ControlRowNumber))
            Else
                'Create the blank Operatives here:
                'STILL WITHIN LOOP
            End If

            ControlRowNumber = ControlRowNumber + 1
        Loop
        frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtTotalOps", CStr(OpTotalRows))
        If OpTotalRows = 0 Then
            InsertOperatives(True, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Operatives, ControlRowNumber, 400, OpStartTAG, TotalControlsInFrame(4),
                                 FieldIDX, OpName, OpActivity, dtStartTime, dtEndTime, OpComments, ControlTABIndex)
            'Now Populate the ControlPrimaryKey here:
            'Load the first record just created.
            'Get the RecordID.
            'use UpdateControls().
            SearchCriteria = "FrameRowNumber = " & ControlRowNumber
            OpLoadedOK = Find_myQuery(frmMainGIForm.myConnString, DBTables(4), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If OpLoadedOK Then
                'Need TAG number or control name to insert primary key value into Dic_Controls():
                '************** OK we have the NewFrameRowNumber - just need to pick a control name to alter: *****************
                ControlName = "txtOpRow:" & CStr(ControlRowNumber)
                OpControl = frmMainGIForm.FindFrameControls(FormName & FormID, ControlName)
                If OpControl IsNot Nothing Then
                    ChangeProperty = "PRIMARYKEY"
                    testCollection = UpdateCollection(dic_Controls, Nothing, ReturnValue, CDate(strDeliveryDate), "",
                                                          strDeliveryRef, OpControl.Tag, ChangeProperty)

                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                End If
            End If
        End If
    End Sub


    Sub PopulateStaticControls(TheDeliveryDate As String, TheDeliveryRef As String, Optional ASNNo As String = "")
        Dim DBTables() As String
        Dim LowerTAG() As Long
        Dim UpperTAG() As Long
        Dim TAbleIDX As Long
        Dim LoadedOK() As Boolean
        Dim FieldIDX As Long
        Dim StartIDX() As Long
        Dim EndIDX() As Long
        Dim FrameRows() As Long
        Dim TotalControlsInFrame() As Long
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim SearchCriteria As String = ""
        Dim SortFields As String = ""
        Dim Reversed As Boolean = False
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllValues() As Object = Nothing
        Dim AllFields() As String = Nothing
        Dim ErrMessage As String = ""
        Dim dtDeliveryDate As DateTime
        Dim strDeliveryDate As String
        Dim FieldType As String = "STRING"
        Dim strSupplier As String
        Dim strDueTime As String
        Dim strSHIFT As String
        Dim strOrigin As String
        Dim strResult As String = ""
        Dim Fieldname As String = ""
        Dim TagNumber As Long = 0
        Dim TimeTagNumber As Long = 0
        Dim TagIndex As Long = 0
        Dim FoundCtrl As Control = Nothing
        Dim ChildCtrl As Control = Nothing
        Dim FieldNameArr() As String
        Dim Fieldnames() As String
        Dim FrameRowNumberField As String
        Dim ControlDBTable As String
        Dim ControlFontName As String
        Dim ControlDeliveryDate As DateTime
        Dim strDeliveryRef As String
        Dim ControlDeliveryRef As String
        Dim MakeVisible As Boolean = True
        Dim ControlType As String = "TEXTBOX"
        Dim ControlText As String = ""
        Dim ControlLeft As Integer = 7
        Dim ControlTop As Integer = 1
        Dim ControlHeight As Integer = 23
        Dim ControlWidth As Integer = 35
        Dim ControlFontSize As Integer = 11
        Dim ControlFontStyle As String
        Dim ControlTAG As String
        Dim ControlDate As String = Now()
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlTotalRows As Long ' need to know lowerTAG and number of fields in frame_Operatives.
        Dim OpStartTAG As Long = 43
        Dim OpFinishTAG As Long = 48
        Dim OpTotalRows As Long = 0
        Dim OpStartTime As String = "00:00:00"
        Dim OpEndTime As String = "00:00:00"
        Dim ShortTotalRows As Long = 0
        Dim ExtraTotalRows As Long = 0
        Dim ShortStartTAG As Long = 0
        Dim ExtraStartTAG As Long = 0
        Dim ShortOrExtra As String = ""
        Dim ControlBACKCOLOUR As Integer = RGB(0, 112, 192) 'BLUE, RGB(240, 248, 255) 'ALICEBLUE
        Dim ControlForeColor As Integer = RGB(255, 245, 60) 'yellow text
        Dim ControlLeftMargin As Boolean = False
        Dim ControlFieldname As String = ""
        Dim ControlRowNumber As Long
        Dim ComboArray() As String = Nothing
        Dim ControlName As String = ""
        Dim varKeyAsn As String = ""
        Dim varKeyRef As String = ""
        Dim Frame_Extra_Parts As ScrollableControl
        Dim Frame_FLMDetails As ScrollableControl
        Dim Frame_InboundSchedule As ScrollableControl
        Dim Frame_OperationalInput As ScrollableControl
        Dim Frame_Operatives As ScrollableControl
        Dim Frame_OpsShortsAndExtras As ScrollableControl
        Dim Frame_Short_Parts As ScrollableControl
        Dim Frame_SupplierCompliance As ScrollableControl
        Dim ControlFrameName As String = ""
        Dim NewIndex As Long = 0
        Dim ListArray() As Object = Nothing
        Dim ControlTextAlign As String = ""
        Dim NewTextBox As New TextBox
        Dim LastRow As Long
        Dim dtStartTime As DateTime
        Dim dtEndTime As DateTime
        Dim FLMName As String = ""
        Dim OpName As String = ""
        Dim OpActivity As String = ""
        Dim OpComments As String = ""
        Dim ShortPartNo As String = ""
        Dim ShortQty As String = ""
        Dim ExtraPartNo As String = ""
        Dim ExtraQty As String = ""
        Dim SavedOK() As Boolean
        Dim LabourFields As String
        Dim LabourFieldValues As String
        Dim SaveCriteria As String = ""
        Dim ExcludeFields As String = ""
        Dim SCFieldsArr() As String = Nothing
        Dim tempTotals As New clsTotals
        Dim ExtractTotals As New clsTotals
        Dim strSaveDeliveryDate As String
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim ControlTABIndex As Long
        Dim StartControlTABIndex As Long
        Dim TotalControlsInOpFrame As Long
        Dim OpGridControl As Control
        Dim ShortGridControl As Control
        Dim ExtraGridControl As Control

        Frame_Extra_Parts = GetFrameControl("GI_TIMESHEET2", "Frame_Extra_Parts")
        Frame_FLMDetails = GetFrameControl("GI_TIMESHEET2", "Frame_FLMDetails")
        Frame_InboundSchedule = GetFrameControl("GI_TIMESHEET2", "Frame_InboundSchedule")
        Frame_OperationalInput = GetFrameControl("GI_TIMESHEET2", "Frame_OperationalInput")
        Frame_Operatives = GetFrameControl("GI_TIMESHEET2", "Frame_Operatives")
        Frame_OpsShortsAndExtras = GetFrameControl("GI_TIMESHEET2", "Frame_OpsShortsAndExtras")
        Frame_Short_Parts = GetFrameControl("GI_TIMESHEET2", "Frame_Short_Parts")
        Frame_SupplierCompliance = GetFrameControl("GI_TIMESHEET2", "Frame_SupplierCompliance")

        'Frame_Operatives.Controls.Clear()
        'Frame_Short_Parts.Controls.Clear()
        'Frame_Extra_Parts.Controls.Clear()

        Frame_FLMDetails.Controls.Clear()


        ReDim DBTables(6)
        ReDim LowerTAG(6)
        ReDim UpperTAG(6)
        ReDim LoadedOK(6)
        ReDim StartIDX(6)
        ReDim EndIDX(6)
        ReDim FrameRows(6)
        ReDim TotalControlsInFrame(6)
        ReDim Fieldnames(6)
        ReDim SavedOK(6)

        If Len(TheDeliveryRef) = 0 Then
            strDeliveryRef = ""
        Else
            strDeliveryRef = TheDeliveryRef
        End If
        If Len(TheDeliveryDate) = 0 Then
            dtDeliveryDate = CDate("1970-01-01 00:00:00")
            strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        Else
            dtDeliveryDate = CDate(TheDeliveryDate)
            strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        End If

        'frmMainGIForm.dic_Totals = CreateObject("Scripting.Dictionary") 'NOT part of ANYMORE in frmMainGIForm - 28/08/2018
        'frmMainGIForm.dic_Totals.comparemode = vbTextCompare

        'rmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) = 1

        tempTotals.Total_Operatives = 0
        tempTotals.Total_ShortParts = 0
        tempTotals.Total_ExtraParts = 0
        tempTotals.HighestOpTAGID = 43
        tempTotals.HighestOpBtnTAGID = 10
        tempTotals.HighestShortTAGID = 1001
        tempTotals.HighestExtraTAGID = 2001
        tempTotals.TotalOpHours = 0
        tempTotals.TotalFLMHours = 0

        If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
        End If

        DBTables(1) = "tblDeliveryInfo"
        DBTables(2) = "tblSupplierCompliance"
        DBTables(3) = "tblLabourHours"
        DBTables(4) = "tblOperatives"
        DBTables(5) = "tblShortsAndExtraParts"
        DBTables(6) = "tblShortsAndExtraParts"

        TotalControlsInFrame(1) = 1
        TotalControlsInFrame(2) = 1
        TotalControlsInFrame(3) = 1
        TotalControlsInFrame(4) = 6
        TotalControlsInFrame(5) = 3
        TotalControlsInFrame(6) = 3
        LowerTAG(1) = 1
        UpperTAG(1) = 30
        LowerTAG(2) = 801
        UpperTAG(2) = 807
        LowerTAG(3) = 40
        UpperTAG(3) = 42
        LowerTAG(4) = 43
        UpperTAG(4) = 48
        LowerTAG(5) = 1001
        UpperTAG(5) = 1003
        LowerTAG(6) = 2001
        UpperTAG(6) = 2003
        If Len(TheDeliveryDate) = 0 Then
            MsgBox("No Delivery Date passed")
            Exit Sub
        End If
        SearchCriteria = ""
        ReturnField = "ID"
        dtDeliveryDate = CDate(TheDeliveryDate)
        strSaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd") & " 00:00:00"
        strDeliveryRef = TheDeliveryRef
        ControlDeliveryRef = TheDeliveryRef
        SearchText = strSaveDeliveryDate
        'SearchCriteria = "([DeliveryDate] = #" & strDeliveryDate & "#) And ([DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34) & ")"
        If Len(ASNNo) > 0 Then
            SearchText = ASNNo
            ControlASN = ASNNo
            SearchField = "ASN_Number"
            FieldType = "STRING"
            'SearchCriteria = "ASN_Number = " & Chr(34) & ASNNo & Chr(34)
        Else
            SearchText = TheDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            'SearchCriteria = "DeliveryReference = " & Chr(34) & TheDeliveryRef & Chr(34)
        End If

        SortFields = ""

        'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)
        'FieldNameArr = DanG_DB_Tools.strToStringArray(Fieldnames, ",", 1, False, False, False, "_", False)
        ReDim FieldNameArr(20)
        'Fieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage)

        Fieldnames(1) = GetMyFields(DBTables(1), frmMainGIForm.myConnString, ErrMessage)
        Fieldnames(2) = GetMyFields(DBTables(2), frmMainGIForm.myConnString, ErrMessage)
        Fieldnames(3) = GetMyFields(DBTables(3), frmMainGIForm.myConnString, ErrMessage)
        Fieldnames(4) = GetMyFields(DBTables(4), frmMainGIForm.myConnString, ErrMessage)
        Fieldnames(5) = GetMyFields(DBTables(5), frmMainGIForm.myConnString, ErrMessage)
        Fieldnames(6) = GetMyFields(DBTables(6), frmMainGIForm.myConnString, ErrMessage)
        FieldNameArr = DanG_DB_Tools.strToStringArray(Fieldnames(1), ",", 1, False, False, False, "_", False)
        FrameRowNumberField = "FrameRowNumber"

        ControlDBTable = DBTables(1)
        ControlFontName = "Cambria"

        If IsDate(strDeliveryDate) Then
            ControlDeliveryDate = CDate(strDeliveryDate)
        Else
            'MsgBox ("Need to pass proper delivery date")
            'Exit Sub
            ControlDeliveryDate = CDate("01/01/1970")
        End If
        ControlDeliveryRef = strDeliveryRef

        MakeVisible = True
        ControlType = "TEXTBOX"
        'ControlText = CStr(OpID)
        ControlLeft = 7
        ControlTop = 1
        ControlHeight = 23
        ControlWidth = 35
        ControlFontSize = 11
        ControlFontStyle = "BOLD"
        'ControlTAG = CStr(TagID)
        ControlDate = Now()
        ControlASN = ""
        ControlOBJCount = 1
        ControlStartTAG = ""
        ControlEndTAG = ""
        ReturnValue = ""
        ControlTotalRows = 0 ' need to know lowerTAG and number of fields in frame_Operatives.
        'BackColor = RGB(240, 248, 255) 'ALICEBLUE
        'ControlBackColor = RGB(0, 112, 192) 'BLUE
        'ControlForeColor = RGB(255, 245, 60) 'yellow text
        'ControlLeftMargin = False
        'ControlFieldname = FrameRowNumberField
        'ControlRowNumber = OpID
        'ControlTotalRows = TotalOps
        'ComboArray = Nothing

        'OperativeFrame = GetFrameControl("GI_TIMESHEET2", "Frame_Operatives")
        'TotalFrameRows = GetTotalFrameRows(OperativeFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(4), VTTAG)
        'Beep()
        MsgBox("IN PopulateStaticControls()")
        Clear_Entry_Controls("GI_TIMESHEET2", 1, 40)
        Clear_Entry_Controls("GI_TIMESHEET2", 801, 807)

        Reversed = False
        LoadedOK(1) = Find_myQuery(frmMainGIForm.myConnString, DBTables(1), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If LoadedOK(1) Then
            ControlDeliveryRef = strDeliveryRef
            strSupplier = AllValues(3)
            ControlASN = AllValues(4)
            If Len(ControlASN) > 0 Then
                varKeyAsn = strDeliveryDate & "_" & ControlASN & "_" & "1"
            End If
            varKeyRef = strDeliveryDate & "_" & ControlDeliveryRef & "_" & "1"
            'Call UpdateCollection(dic_Controls, "", AllValues(1), AllValues(1), AllValues(5), AllValues(2), "1", "VALUE")
            'Call UpdateCollection(dic_Controls, "", Frame_InboundSchedule.Name, AllValues(1), AllValues(5), AllValues(2), "4", "FRAMENAME")
            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtDeliveryDate", AllValues(1))
            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtDeliveryRef", AllValues(2))
            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtSupplier", strSupplier)
            FieldIDX = 1
            TagNumber = LowerTAG(1)
            Do While FieldIDX < UBound(AllValues) - 1
                ControlFieldname = AllFields(FieldIDX)
                strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                'Find correct field to put value into now.
                'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
                If FoundCtrl IsNot Nothing Then
                    'put value into control:
                    frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                    'Insert into Dic_Controls:
                    varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber
                    ControlName = FoundCtrl.Name
                    ControlTAG = CStr(TagNumber)
                    ControlTABIndex = FoundCtrl.TabIndex
                    ControlText = strResult
                    ControlType = "TEXTBOX"
                    ControlLeft = FoundCtrl.Left
                    ControlTop = FoundCtrl.Top
                    ControlWidth = FoundCtrl.Width
                    ControlHeight = FoundCtrl.Height
                    ControlOBJCount = TagNumber
                    ControlStartTAG = "1"
                    ControlEndTAG = "27"
                    ControlRowNumber = 1
                    ControlTotalRows = 1
                    ListArray = Nothing
                    ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                    ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                    ControlFontName = FoundCtrl.Font.Name
                    ControlFontSize = FoundCtrl.Font.Size
                    If FoundCtrl.Font.Bold Then
                        ControlFontStyle = "BOLD"
                    Else
                        ControlFontStyle = "NORMAL"
                    End If
                    If FoundCtrl.Font.Italic Then
                        ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                    End If
                    If FoundCtrl.Font.Underline Then
                        ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                    End If
                    If FoundCtrl.Font.Strikeout Then
                        ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                    End If
                    If TypeOf (FoundCtrl) Is TextBox Then
                        NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                        If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                            ControlTextAlign = "LEFT"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                            ControlTextAlign = "CENTER"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                            ControlTextAlign = "RIGHT"
                        Else
                            ControlTextAlign = ""
                        End If
                    End If
                    NewIndex = AddNewControl(False, Frame_InboundSchedule, DBTables(1), ControlFieldname, "ID", FoundCtrl, ControlName,
                                             ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                                             ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                             frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                                             ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                             True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                                             ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False)
                End If
                FieldIDX = FieldIDX + 1
                TagNumber = TagNumber + 1
            Loop
            'strOrigin = GetMYValuebyFieldname(AllValues, AllFields, "Origin")
        End If

        FieldIDX = 3
        SCFieldsArr = strToStringArray(SCFieldnames, ",", 0, False, False, False, "_", False, 34, 39)
        LoadedOK(2) = Find_myQuery(frmMainGIForm.myConnString, DBTables(2), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        'If LoadedOK(2) Then
        'tblsupplierCompliance:
        If Not LoadedOK(2) Then
            dtLastSaved = Now()
            strLastSaved = dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss")
            SaveCriteria = ""
            LabourFields = "DeliveryDate,DeliveryReference,LastSaved"
            LabourFieldValues = Chr(39) & strDeliveryDate & Chr(39) & "," & strDeliveryRef & "," & Chr(39) & strLastSaved & Chr(39)
            'LabourFieldValues = strDeliveryDate & ";" & strDeliveryRef & ";" & strLastSaved
            ExcludeFields = ""
            SavedOK(2) = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, DBTables(2), LabourFields, LabourFieldValues, ErrMessage, SaveCriteria, ExcludeFields)
            'SavedOK(2) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, "", DBTables(2), LabourFields, LabourFieldValues, SaveCriteria, ExcludeFields, ErrMessage)
        End If
        Clear_Entry_Controls("GI_TIMESHEET2", 801, 807)
        TagNumber = LowerTAG(2)
        Do While FieldIDX < UBound(SCFieldsArr) - 1
            ControlFieldname = SCFieldsArr(FieldIDX)
            If ControlFieldname Is Nothing Then
                Continue Do
            End If
            If LoadedOK(2) Then
                strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
            Else
                strResult = ""
            End If

            'Find correct field to put value into now.
            'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
            FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
            If FoundCtrl IsNot Nothing Then
                'put value into control:
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                'Insert into Dic_Controls:
                varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber
                ControlName = FoundCtrl.Name
                ControlTAG = CStr(TagNumber)
                ControlTABIndex = FoundCtrl.TabIndex
                ControlText = strResult
                ControlType = "TEXTBOX"
                ControlLeft = FoundCtrl.Left
                ControlTop = FoundCtrl.Top
                ControlWidth = FoundCtrl.Width
                ControlHeight = FoundCtrl.Height
                ControlOBJCount = TagNumber
                ControlStartTAG = CStr(LowerTAG(2))
                ControlEndTAG = CStr(UpperTAG(2))
                ControlRowNumber = 1
                ControlTotalRows = 1
                ListArray = Nothing
                ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                ControlFontName = FoundCtrl.Font.Name
                ControlFontSize = FoundCtrl.Font.Size
                If FoundCtrl.Font.Bold Then
                    ControlFontStyle = "BOLD"
                Else
                    ControlFontStyle = "NORMAL"
                End If
                If FoundCtrl.Font.Italic Then
                    ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                End If
                If FoundCtrl.Font.Underline Then
                    ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                End If
                If FoundCtrl.Font.Strikeout Then
                    ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                End If
                If TypeOf (FoundCtrl) Is TextBox Then
                    NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                    If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                        ControlTextAlign = "LEFT"
                    ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                        ControlTextAlign = "CENTER"
                    ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                        ControlTextAlign = "RIGHT"
                    Else
                        ControlTextAlign = ""
                    End If
                End If
                NewIndex = AddNewControl(False, Frame_SupplierCompliance, DBTables(2), ControlFieldname, "ID", FoundCtrl, ControlName,
                                             ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                                             ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                             frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                                             ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                             True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                                             ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False)
            End If 'FoundCtrl
            FieldIDX = FieldIDX + 1
            TagNumber = TagNumber + 1
        Loop

        'CHECK FLM DETAILS and LOAD:

        AllValues = Nothing
        AllFields = Nothing
        SortFields = ""
        Reversed = False
        ErrMessage = ""
        SearchCriteria = ""
        SearchField = "DELIVERYREFERENCE"
        SearchText = strDeliveryRef
        FieldType = "String"
        ReturnField = "ID"
        ReturnValue = ""

        LoadedOK(3) = Find_myQuery(frmMainGIForm.myConnString, DBTables(3), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If LoadedOK(3) Then
            'tblLAbourHours: - extract FLM info:
            FieldIDX = 3
            TagNumber = LowerTAG(3)
            ControlType = "COMBOBOX"
            ControlTABIndex = 33 'START TAB INDEX
            Call InsertFLMButtons(False, Frame_FLMDetails, strDeliveryDate, strDeliveryRef, ControlASN, "", FLMName, dtStartTime, dtEndTime, ControlTABIndex)
            Do While FieldIDX < 6
                ControlFieldname = AllFields(FieldIDX)
                strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                If UCase(ControlFieldname) = UCase("FLM_Name") Then
                    FLMName = strResult
                End If
                If UCase(ControlFieldname) = UCase("FLM_StartTime") Then
                    dtStartTime = CDate(strResult)
                    strResult = dtStartTime.ToString("HH:mm:ss")

                End If
                If UCase(ControlFieldname) = UCase("FLM_FinishTime") Then
                    dtEndTime = CDate(strResult)
                    strResult = dtEndTime.ToString("HH:mm:ss")

                End If
                'Find correct field to put value into now.
                If FieldIDX = 4 Then
                    TagNumber = TagNumber + 400 'DANGEROUS - any more controls in this frame - use a sep variable.
                    ControlType = "TEXTBOX"
                End If
                'FoundCtrl = frmMainGIForm.FindControl_Recursive(frmMainGIForm.)
                FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
                If FoundCtrl IsNot Nothing Then
                    'put value into control:
                    frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                    'Insert into Dic_Controls:
                    varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber 'WRONG DATE FORMAT - needs / not -
                    ControlName = FoundCtrl.Name
                    ControlTAG = CStr(TagNumber)
                    ControlText = strResult

                    ControlLeft = FoundCtrl.Left
                    ControlTop = FoundCtrl.Top
                    ControlWidth = FoundCtrl.Width
                    ControlHeight = FoundCtrl.Height
                    ControlOBJCount = TagNumber
                    'LOAD FROM DATABASE TABLE tblLAbourHours:
                    If GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID") > 0 Then
                        OpStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID")
                    Else
                        OpStartTAG = LowerTAG(4)
                    End If
                    If GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID") > 0 Then
                        OpFinishTAG = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")
                    Else
                        OpFinishTAG = UpperTAG(4)
                    End If

                    tempTotals.OpStartTAG = OpStartTAG
                    tempTotals.OpFinishTAG = OpFinishTAG

                    dtDeliveryDate = CDate(strDeliveryDate)
                    strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
                    'The following adds an incorect property item to the dic_controls - due to the DeliveryDate being wrong format:
                    ' - Amended. check again.
                    dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

                    ControlStartTAG = CStr(OpStartTAG)
                    ControlEndTAG = CStr(OpFinishTAG)
                    ControlRowNumber = 0

                    TotalControlsInOpFrame = Frame_Operatives.Controls.Count 'RETURNS 1 as its the Data GRid View only in there !!!!
                    OpTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
                    If OpTotalRows > 0 Then
                        tempTotals.Total_Operatives = OpTotalRows
                        dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
                    End If
                    ListArray = Nothing
                    ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                    ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                    ControlFontName = FoundCtrl.Font.Name
                    ControlFontSize = FoundCtrl.Font.Size
                    If FoundCtrl.Font.Bold Then
                        ControlFontStyle = "BOLD"
                    Else
                        ControlFontStyle = "NORMAL"
                    End If
                    If FoundCtrl.Font.Italic Then
                        ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                    End If
                    If FoundCtrl.Font.Underline Then
                        ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                    End If
                    If FoundCtrl.Font.Strikeout Then
                        ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                    End If
                    If TypeOf (FoundCtrl) Is TextBox Then
                        NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                        If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                            ControlTextAlign = "LEFT"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                            ControlTextAlign = "CENTER"
                        ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                            ControlTextAlign = "RIGHT"
                        Else
                            ControlTextAlign = ""
                        End If
                    End If
                    'set tempControl to clsControls then can use tempControl.ControlTag = "48", tempControl.ControlFontStyle = "BOLD" etc. and just pass tempControl
                    NewIndex = AddNewControl(False, Frame_FLMDetails, DBTables(3), ControlFieldname, "ID", FoundCtrl, ControlName,
                    ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                    ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                             frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                    ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                            True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                    ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False, dtStartTime, dtEndTime)
                End If
                FieldIDX = FieldIDX + 1
                TagNumber = TagNumber + 1
            Loop
        Else
            'NO FLM record so INSERT NEW FLM Details Record into tblLabourHours:
            ControlTABIndex = 33
            Call InsertFLMButtons(False, Frame_FLMDetails, strDeliveryDate, strDeliveryRef, ControlASN, "", FLMName, dtStartTime, dtEndTime, ControlTABIndex)

        End If

        '*************************************
        'LOAD tblOperatives:
        '*************************************

        AllValues = Nothing
        AllFields = Nothing
        SortFields = ""
        Reversed = False
        ErrMessage = ""
        SearchCriteria = ""
        SearchField = "DELIVERYREFERENCE"
        SearchText = strDeliveryRef
        FieldType = "String"
        ReturnField = "ID"
        ReturnValue = ""


        ControlRowNumber = 1
        If OpStartTAG = 0 Then
            OpStartTAG = LowerTAG(4)
        End If
        TagNumber = OpStartTAG
        ControlStartTAG = CStr(TagNumber)
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        ControlTotalRows = OpTotalRows
        StartControlTABIndex = ExtractTotals.HighestOpTabIndex
        ControlTABIndex = StartControlTABIndex

        Do While ControlRowNumber <= OpTotalRows
            SearchCriteria = "FrameRowNumber = " & ControlRowNumber
            LoadedOK(4) = Find_myQuery(frmMainGIForm.myConnString, DBTables(4), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            If LoadedOK(4) Then
                'tblOperatives: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                FieldIDX = 4
                TimeTagNumber = 0
                ControlFieldname = AllFields(FieldIDX)
                OpName = GetMYValuebyFieldname(AllValues, AllFields, "Op_Name")
                OpActivity = GetMYValuebyFieldname(AllValues, AllFields, "Activity_Name")
                OpComments = GetMYValuebyFieldname(AllValues, AllFields, "OpComments")
                strResult = GetMYValuebyFieldname(AllValues, AllFields, "Op_StartTime")
                If IsDate(strResult) Then
                    dtStartTime = CDate(strResult)
                Else
                    dtStartTime = CDate("1970-01-01 01:00:00")
                End If
                strResult = GetMYValuebyFieldname(AllValues, AllFields, "Op_FinishTime")
                If IsDate(strResult) Then
                    dtEndTime = CDate(strResult)
                Else
                    dtEndTime = CDate("1970-01-01 01:00:00")
                End If
                'LastRow is the ONLY place used here - but the operative record seems to load in still ok ????
                'it does trigger off a few odd values - like totalrows = -1 etc.
                'LastRow ??????? should be TotalRows here :

                'gridname = Frame_Operatives.dgvOperatives.name 'No does not work - need a control var to refer to directly.
                InsertOperatives(False, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Operatives, ControlRowNumber, 400, OpStartTAG, TotalControlsInFrame(4),
                                 FieldIDX, OpName, OpActivity, dtStartTime, dtEndTime, OpComments, ControlTABIndex)

                Do While FieldIDX <= 9
                    'TagNumber = TagNumber + 1
                    TimeTagNumber = 0
                    ControlTAG = CStr(TagNumber)
                    ControlFieldname = AllFields(FieldIDX)
                    strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                    If UCase(ControlFieldname) = UCase("Op_StartTime") Then
                        TimeTagNumber = TagNumber + 400
                        ControlTAG = CStr(TimeTagNumber)
                        ControlType = "TEXTBOX"
                        dtStartTime = CDate(strResult)
                        strResult = dtStartTime.ToString("HH:mm:ss")
                        OpStartTime = strResult
                    End If
                    If UCase(ControlFieldname) = UCase("Op_FinishTime") Then
                        TimeTagNumber = TagNumber + 400
                        ControlTAG = CStr(TimeTagNumber)
                        ControlType = "TEXTBOX"
                        dtEndTime = CDate(strResult)
                        strResult = dtEndTime.ToString("HH:mm:ss")
                        OpEndTime = strResult
                    End If
                    If TimeTagNumber > 0 Then
                        FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TimeTagNumber), ChildCtrl)
                    Else
                        FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
                    End If
                    If FoundCtrl IsNot Nothing Then
                        frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                        strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                        'Find correct field to put value into now.
                        'ControlTAG = cstr(TagNumber)
                        If UCase(ControlFieldname) = UCase("FrameRowNumber") Then
                            ControlType = "TEXTBOX"
                        End If
                        If UCase(ControlFieldname) = UCase("Op_Name") Then
                            ControlType = "COMBOBOX"
                            OpName = strResult
                        End If
                        If UCase(ControlFieldname) = UCase("Activity_Name") Then
                            ControlType = "COMBOBOX"

                        End If
                        If UCase(ControlFieldname) = UCase("OpComments") Then
                            ControlType = "TEXTBOX"
                        End If

                        ControlText = strResult
                        'put value into control:

                        'Insert into Dic_Controls:

                        varKeyRef = strDeliveryDate & "_" & strDeliveryRef & "_" & TagNumber 'not correct if Time is the field  - not adding 400.
                        ControlName = FoundCtrl.Name
                        'ControlTAG = CStr(TagNumber)
                        ControlText = strResult
                        ControlLeft = FoundCtrl.Left
                        ControlTop = FoundCtrl.Top
                        ControlWidth = FoundCtrl.Width
                        ControlHeight = FoundCtrl.Height
                        ControlOBJCount = TagNumber ' will record 46 rather than 446 for the time.
                        ControlStartTAG = CStr(OpStartTAG)
                        ControlEndTAG = CStr(OpFinishTAG)
                        'ControlStartTAG = CStr(OpStartTAG)
                        'ControlEndTAG = CStr(OpFinishTAG)
                        'OpTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
                        ListArray = Nothing
                        ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                        ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                        ControlFontName = FoundCtrl.Font.Name
                        ControlFontSize = FoundCtrl.Font.Size
                        If FoundCtrl.Font.Bold Then
                            ControlFontStyle = "BOLD"
                        Else
                            ControlFontStyle = "NORMAL"
                        End If
                        If FoundCtrl.Font.Italic Then
                            ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                        End If
                        If FoundCtrl.Font.Underline Then
                            ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                        End If
                        If FoundCtrl.Font.Strikeout Then
                            ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                        End If
                        If TypeOf (FoundCtrl) Is TextBox Then
                            NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                            If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                ControlTextAlign = "LEFT"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                ControlTextAlign = "CENTER"
                            ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                ControlTextAlign = "RIGHT"
                            Else
                                ControlTextAlign = ""
                            End If
                        End If

                        NewIndex = AddNewControl(False, Frame_Operatives, DBTables(4), ControlFieldname, "ID", FoundCtrl, ControlName,
                            ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                            ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                 frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                            ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                        True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                            ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False, dtStartTime, dtEndTime)
                    End If 'FoundCtrl
                    FieldIDX = FieldIDX + 1
                    TagNumber = TagNumber + 1
                    ControlTABIndex = ControlTABIndex + 1

                Loop
                'TagNumber = TagNumber + 1
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtTotalOps", CStr(ControlRowNumber))
            Else
                'Create the blank Operatives here:
                'STILL WITHIN LOOP
            End If

            ControlRowNumber = ControlRowNumber + 1
        Loop
        If OpTotalRows = 0 Then
            InsertOperatives(True, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Operatives, ControlRowNumber, 400, OpStartTAG, TotalControlsInFrame(4),
                                 FieldIDX, OpName, OpActivity, dtStartTime, dtEndTime, OpComments, ControlTABIndex)
        End If


        SortFields = ""
        Reversed = False
        ErrMessage = ""
        ShortOrExtra = "SHORT"
        SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
        LoadedOK(5) = Find_myQuery(frmMainGIForm.myConnString, DBTables(5), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If LoadedOK(5) Then
            ControlRowNumber = 1
            ShortTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
            ControlStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "Start_TAGID")
            ControlEndTAG = GetMYValuebyFieldname(AllValues, AllFields, "End_TAGID")

            tempTotals.Total_ShortParts = ShortTotalRows
            tempTotals.HighestShortTAGID = ControlEndTAG
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
            tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If Not tempTotals Is Nothing Then
                StartControlTABIndex = tempTotals.HighestShortTABIndex
            Else
                StartControlTABIndex = 1001
            End If
            ControlTABIndex = StartControlTABIndex

            If IsNumeric(ControlStartTAG) Then
                TagNumber = CLng(ControlStartTAG)
            Else
                TagNumber = 1001
            End If
            ControlStartTAG = CStr(TagNumber)
            ControlTotalRows = ShortTotalRows
            Do While ControlRowNumber <= ShortTotalRows
                SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
                SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & ControlRowNumber
                LoadedOK(5) = Find_myQuery(frmMainGIForm.myConnString, DBTables(5), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                If LoadedOK(5) Then
                    'tblshortsandextraparts: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                    FieldIDX = 4
                    ShortPartNo = GetMYValuebyFieldname(AllValues, AllFields, "PartNo")
                    ShortQty = GetMYValuebyFieldname(AllValues, AllFields, "Qty")
                    InsertShortParts(False, strDeliveryDate, strDeliveryRef, ControlASN, Frame_Short_Parts, LastRow, 1000, ControlStartTAG, 3, FieldIDX,
                                     ShortPartNo, ShortQty, ControlTABIndex)
                    Do While FieldIDX <= 6
                        ControlFieldname = AllFields(FieldIDX)
                        strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                        'Find correct field to put value into now.
                        If UCase(ControlFieldname) = UCase("PartNo") Then
                            ControlType = "TEXTBOX"
                            ShortPartNo = strResult
                        End If
                        If UCase(ControlFieldname) = UCase("Qty") Then
                            ControlType = "TEXTBOX"
                            ShortQty = strResult
                        End If
                        FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
                        If FoundCtrl IsNot Nothing Then
                            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                            ControlText = strResult
                            ControlName = FoundCtrl.Name
                            ControlTAG = CStr(TagNumber)
                            ControlOBJCount = TagNumber
                            ControlTABIndex = FoundCtrl.TabIndex
                            ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                            ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                            ControlFontName = FoundCtrl.Font.Name
                            ControlFontSize = FoundCtrl.Font.Size
                            If FoundCtrl.Font.Bold Then
                                ControlFontStyle = "BOLD"
                            Else
                                ControlFontStyle = "NORMAL"
                            End If
                            If FoundCtrl.Font.Italic Then
                                ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                            End If
                            If FoundCtrl.Font.Underline Then
                                ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                            End If
                            If FoundCtrl.Font.Strikeout Then
                                ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                            End If
                            If TypeOf (FoundCtrl) Is TextBox Then
                                NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                                If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                    ControlTextAlign = "LEFT"
                                ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                    ControlTextAlign = "CENTER"
                                ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                    ControlTextAlign = "RIGHT"
                                Else
                                    ControlTextAlign = ""
                                End If
                            End If
                            NewIndex = AddNewControl(False, Frame_Short_Parts, DBTables(5), ControlFieldname, "ID", FoundCtrl, ControlName,
                            ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                            ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                     frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                            ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                        True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                            ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False)
                        End If
                        FieldIDX = FieldIDX + 1
                        TagNumber = TagNumber + 1
                    Loop
                    frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtTotalShorts", CStr(ControlRowNumber))
                    'what about using TotalShorts ???
                Else
                    'Cannot Find record:

                End If 'If the current numbered row exists
                ControlRowNumber = ControlRowNumber + 1
            Loop
            tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            tempTotals.HighestShortTABIndex = ControlTABIndex
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
        End If 'IF any rows exist in tblShortsAndExtraParts WHERE ShortORExtra = Short

        ShortOrExtra = "Extra"
        SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
        LoadedOK(6) = Find_myQuery(frmMainGIForm.myConnString, DBTables(6), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If LoadedOK(6) Then
            ControlRowNumber = 1
            ExtraTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "TotalRows")
            ControlStartTAG = GetMYValuebyFieldname(AllValues, AllFields, "StartTAGID")
            ControlEndTAG = GetMYValuebyFieldname(AllValues, AllFields, "EndTAGID")

            tempTotals.Total_ExtraParts = ExtraTotalRows
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

            tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If Not tempTotals Is Nothing Then
                StartControlTABIndex = tempTotals.HighestExtraTABIndex
            Else
                StartControlTABIndex = 2001
            End If
            ControlTABIndex = StartControlTABIndex

            If IsNumeric(ControlStartTAG) Then
                TagNumber = CLng(ControlStartTAG)
            Else
                TagNumber = 2001
            End If
            ControlStartTAG = CStr(TagNumber)
            ControlTotalRows = ExtraTotalRows
            Do While ControlRowNumber <= ExtraTotalRows
                SearchCriteria = "ShortOrExtra = " & Chr(34) & ShortOrExtra & Chr(34)
                SearchCriteria = SearchCriteria & " AND FrameRowNumber = " & ControlRowNumber
                LoadedOK(6) = Find_myQuery(frmMainGIForm.myConnString, DBTables(6), SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                If LoadedOK(6) Then
                    'tblshortsandextraparts: - extract Operative info and using some of the previous info from tblLaourHours - totalrows and starttag and endtag.
                    FieldIDX = 4
                    ExtraPartNo = ""
                    ExtraQty = ""
                    InsertExtraParts(strDeliveryDate, strDeliveryRef, ControlASN, Frame_Extra_Parts, LastRow, 2000, ControlStartTAG, 3, FieldIDX,
                                     ExtraPartNo, ExtraQty)
                    Do While FieldIDX <= 6
                        ControlFieldname = AllFields(FieldIDX)
                        strResult = GetMYValuebyFieldname(AllValues, AllFields, ControlFieldname)
                        'Find correct field to put value into now.
                        If UCase(ControlFieldname) = UCase("PartNo") Then
                            ControlType = "TEXTBOX"
                            ExtraPartNo = strResult
                        End If
                        If UCase(ControlFieldname) = UCase("Qty") Then
                            ControlType = "TEXTBOX"
                            ExtraQty = strResult
                        End If
                        FoundCtrl = frmMainGIForm.FindControls("GI_TIMESHEET2", "", CStr(TagNumber), ChildCtrl)
                        If FoundCtrl IsNot Nothing Then
                            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", FoundCtrl.Name, strResult)
                            ControlText = strResult
                            ControlName = FoundCtrl.Name
                            ControlTAG = CStr(TagNumber)
                            ControlOBJCount = TagNumber
                            ControlTABIndex = FoundCtrl.TabIndex
                            ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                            ControlForeColor = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.ForeColor)
                            ControlFontName = FoundCtrl.Font.Name
                            ControlFontSize = FoundCtrl.Font.Size
                            If FoundCtrl.Font.Bold Then
                                ControlFontStyle = "BOLD"
                            Else
                                ControlFontStyle = "NORMAL"
                            End If
                            If FoundCtrl.Font.Italic Then
                                ControlFontStyle = ControlFontStyle & "," & "ITALIC"
                            End If
                            If FoundCtrl.Font.Underline Then
                                ControlFontStyle = ControlFontStyle & "," & "UNDERLINE"
                            End If
                            If FoundCtrl.Font.Strikeout Then
                                ControlFontStyle = ControlFontStyle & "," & "STRIKEOUT"
                            End If
                            If TypeOf (FoundCtrl) Is TextBox Then
                                NewTextBox = DirectCast(FoundCtrl, System.Windows.Forms.TextBox)
                                If NewTextBox.TextAlign = HorizontalAlignment.Left Then
                                    ControlTextAlign = "LEFT"
                                ElseIf NewTextBox.TextAlign = HorizontalAlignment.Center Then
                                    ControlTextAlign = "CENTER"
                                ElseIf NewTextBox.TextAlign = HorizontalAlignment.Right Then
                                    ControlTextAlign = "RIGHT"
                                Else
                                    ControlTextAlign = ""
                                End If
                            End If
                            NewIndex = AddNewControl(False, Frame_Short_Parts, DBTables(6), ControlFieldname, "ID", FoundCtrl, ControlName,
                            ControlText, ControlType, ControlTAG, ControlTABIndex, ControlDate, ControlLeft, ControlTop,
                            ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                                                     frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
                            ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows,
                                        True, ListArray, ControlBACKCOLOUR, False, ControlFontName, ControlFontSize,
                            ControlFontStyle, ControlForeColor, ControlTextAlign, "MIDDLE CENTER", False, False)
                        End If
                        FieldIDX = FieldIDX + 1
                        TagNumber = TagNumber + 1
                    Loop
                    frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET2", "txtTotalExtras", CStr(ControlRowNumber))

                Else
                    'Not found record:

                End If 'If the current numbered row exists
                ControlRowNumber = ControlRowNumber + 1
            Loop
        End If 'IF any rows exist in tblShortsAndExtraParts WHERE ShortORExtra = Extra
    End Sub

    Sub ResetControls(Optional LowerRange As Long = 0, Optional UpperRange As Long = 0, Optional Confirmation As Boolean = True,
                      Optional FrameName As String = "", Optional ClearFrameControls As Boolean = False, Optional FormName As String = "GI_TIMESHEET_")
        Dim TagNumber As Long
        Dim IDX As Long
        Dim FoundControl As Control
        Dim ctrl As Control
        Dim ctrlFrame As ScrollableControl
        Dim Answer As Integer

        If Len(FrameName) > 0 Then
            ctrlFrame = GetFrameControl(FormName & FormID, FrameName)

            If ClearFrameControls Then
                For Each ctrl In ctrlFrame.Controls
                    If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
                        ctrl.Text = ""
                    End If
                Next
            Else
                ctrlFrame.Controls.Clear()
            End If
        End If

        If Confirmation = False Then
            Answer = vbYes
        Else
            Answer = MsgBox("Are you sure ?", vbYesNoCancel, "CLEAR ALL ENTRY BOXES ???")
        End If

        If Answer = vbYes Then
            For IDX = LowerRange To UpperRange
                TagNumber = IDX
                FoundControl = frmMainGIForm.FindControls(FormName & FormID, "", CStr(IDX))
                If FoundControl IsNot Nothing Then
                    frmMainGIForm.InsertValueIntoForm(FormName & FormID, FoundControl.Name, "")
                End If
            Next
        End If

    End Sub

    Function GetFrameControl(Formname As String, FrameName As String) As ScrollableControl
        Dim FrameControl As New ScrollableControl
        Dim FindFrameControl As New Control
        Dim frm As Form
        Dim MEssage As String
        Dim frmCount As Integer

        frmCount = 0
        Try
            FrameControl = Nothing
            'MsgBox("FORM NAME= " & Formname & ", CP FormName=" & CPFormName & ", Form ID=" & FormID)
            For Each frm In Application.OpenForms
                frmCount = frmCount + 1
                If UCase(frm.Name) = UCase(Formname) Then
                    FindFrameControl = frmMainGIForm.FindControls(Formname, FrameName, "")
                    'FindFrameControl = frmMainGIForm.FindFrameControls(Formname, FrameName)
                    If FindFrameControl IsNot Nothing Then
                        FrameControl = CType(FindFrameControl, System.Windows.Forms.ScrollableControl)
                        GetFrameControl = FrameControl
                    Else
                        'MsgBox("Problem with GetFRAMEControl (Forms Open=" & CStr(frmCount) & ") - not a valid form name passed: FormID= " & FormID)
                        frmMainGIForm.txtMessages.Text = "Problem with GetFRAMEControl (Forms Open=" & CStr(frmCount) & ") - not a valid form name passed: FormID= " & FormID
                    End If
                End If
            Next

        Catch ex As Exception
            Message = "Exception Error In Test Asset-Register DB Login: " & ex.Message
            frmMainGIForm.logger.LogError("GI_Error_v1_2.log", Application.StartupPath, MEssage, "GetFrameControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("GI_Message_v1_2.log", Application.StartupPath, MEssage, "Something wrong with Frame / FormName Passed?" & CStr(Formname) & " " & CStr(FrameName))
        End Try

    End Function

    Function SaveAllControls(ByVal DeliveryDate As String, ByVal DeliveryRef As String, Optional FormName As String = "GI_TIMESHEET_") As Boolean
        Dim VarControl As System.String
        Dim ctrlProperty As clsControls
        Dim OperativeFrame As ScrollableControl = Nothing
        Dim ShortsFrame As ScrollableControl = Nothing
        Dim ExtrasFrame As ScrollableControl = Nothing
        Dim SCFrame As ScrollableControl = Nothing
        Dim TagName As String
        Dim ControlType As String
        Dim ControlName As String
        Dim ControlLowerTag As Long
        Dim ControlUpperTag As Long
        Dim ControlFieldname As String
        Dim ControlDBTable As String
        Dim ControlValue As String
        Dim PropertyName As String
        Dim ControlTAG As String
        Dim ControlID As String
        Dim FrameRow As Long
        Dim TotalFrameRows As Long
        Dim TotalFrames As Long
        Dim TotalControlsPerRow As Long
        Dim VTTAG As Long
        Dim OpLowerTag As Long
        Dim RowNumber As Long
        Dim RowIDX As Long
        Dim FieldIDX As Long
        Dim SaveOK() As Boolean
        Dim IsUpdate() As Boolean
        Dim EncaseFields As Boolean
        Dim RecID() As String
        Dim Fieldnames() As String
        Dim FieldValues() As String
        Dim UpdateCriteria() As String
        Dim ExcludeFields() As String
        Dim ErrMessages() As String
        Dim ValueDelim As String
        Dim ControlDeliveryDate As Date
        Dim ControlDeliveryRef As String
        Dim DBTables() As String
        Dim FoundRecord() As Boolean
        Dim ControlFieldsTable As String
        Dim JustTAG As Boolean
        Dim SearchCriteria As String
        Dim allLookupFields As String
        Dim dict_ReturnLookupFields As New Scripting.Dictionary
        Dim dic_SaveUs As New Scripting.Dictionary
        Dim LookupFields As String
        Dim LookupValues As String
        Dim FieldsTableLoadedOK As Boolean
        Dim Messages As String
        Dim dtLastSavedDate As DateTime
        'Dim FieldOpControls() As String
        'Dim FieldShortControls() As String
        'Dim FieldExtraControls() As String
        'Dim ValueOpControls() As String
        'Dim ValueShortControls() As String
        'Dim ValueExtraControls() As String
        Dim CurrentRow As Long
        Dim TotalRows As Long
        Dim OpTotalRows As Long = 0
        Dim TotalProperty As clsTotals
        Dim ValueProperty As clsControls
        Dim FLMStartTime As String
        Dim FLMEndTime As String
        Dim OpStartTime As String
        Dim OpEndTime As String
        Dim varKey As String
        Dim dynamFieldsArr() As String
        Dim dynamValuesArr() As String
        Dim ShortOrExtra As String
        Dim ValueIDX As Long
        Dim TotalVAlues As Long
        Dim TAbleIDX As Long
        Dim FrameValue As String
        Dim SearchField As String = ""
        Dim SearchValue As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim SplitRowArr() As String
        Dim OperativeName As String
        Dim Activity As String
        Dim strPartNo As String
        Dim strQty As String
        Dim ExtendCriteria As String
        Dim AllFieldnames() As String = Nothing
        Dim AllFieldValues() As Object = Nothing
        Dim ControlsPerRow() As Long
        Dim ValueArray() As String
        Dim ControlTotalRows As Long
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim strUsername As String
        Dim strEmpNo As String
        Dim strName As String
        Dim dtSaveDeliveryDate As DateTime
        Dim strSaveDeliveryDate As String
        Dim dtDeliveryDate As DateTime
        Dim LockSavedOK As Boolean = False
        Dim StatusSavedOK As Boolean = False
        Dim LockError As String
        Dim StatusError As String
        Dim FLMCompleteMessage As String = ""
        Dim OpCompleteMessage As String = ""
        Dim OpComments As String = ""
        Dim TotalShorts As Long = 0
        Dim TotalExtras As Long = 0

        ReDim DBTables(6)
        ReDim Fieldnames(6)
        ReDim FieldValues(6)
        ReDim SaveOK(6)
        ReDim ErrMessages(6)
        ReDim UpdateCriteria(6)
        ReDim IsUpdate(6)
        ReDim FoundRecord(6)
        ReDim RecID(6)
        ReDim ControlsPerRow(6)
        ReDim ExcludeFields(6)
        ReDim ValueArray(50)

        'ReDim FieldOpControls(20, 1)
        'ReDim FieldShortControls(20, 1)
        'ReDim FieldExtraControls(20, 1)
        'ReDim ValueOpControls(20, 1)
        'ReDim ValueShortControls(20, 1)
        'ReDim ValueExtraControls(20, 1)

        If Len(DeliveryDate) = 0 Then
            'MsgBox("No Delivery Date Passed to SAVE procedure")
            frmMainGIForm.txtMessages.Text = "No Delivery Date Passed to SAVE procedure"
            Exit Function
        End If
        If Len(DeliveryRef) = 0 Then
            'MsgBox("No Delivery Reference Passed to SAVE procedure")
            frmMainGIForm.txtMessages.Text = "No Delivery Reference Passed to SAVE procedure"
            Exit Function
        End If

        SaveAllControls = False
        dic_SaveUs.CompareMode = vbTextCompare

        DeliveryDate = Replace(DeliveryDate, "-", "/")
        strSaveDeliveryDate = CDate(DeliveryDate).ToString("yyyy-MM-dd HH:mm:ss")

        DBTables(1) = "tblDeliveryInfo"
        DBTables(2) = "tblSupplierCompliance"
        DBTables(3) = "tblLabourHours"
        DBTables(4) = "tblOperatives"
        DBTables(5) = "tblShortsAndExtraParts"
        DBTables(6) = "tblShortsAndExtraParts"

        ControlsPerRow(1) = 1
        ControlsPerRow(2) = 1
        ControlsPerRow(3) = 1
        ControlsPerRow(4) = 6
        ControlsPerRow(5) = 3
        ControlsPerRow(6) = 3


        ControlFieldsTable = "tblFieldsAndTAGS"

        Fieldnames(1) = ""
        Fieldnames(2) = ""
        Fieldnames(3) = ""
        Fieldnames(4) = ""
        Fieldnames(5) = ""
        Fieldnames(6) = ""
        FieldValues(1) = ""
        FieldValues(2) = ""
        FieldValues(3) = ""
        FieldValues(4) = ""
        FieldValues(5) = ""
        FieldValues(6) = ""

        Messages = ""

        frmMainGIForm.txtMessages.Text = "OK SAVING ... PLEASE WAIT"

        EncaseFields = False
        For TAbleIDX = 1 To 6
            ErrMessages(TAbleIDX) = ""
            ValueDelim = ","
            ExcludeFields(1) = "DeliveryComments,TAGID"
            SearchCriteria = ""
            SearchField = "DeliveryReference"
            SearchValue = DeliveryRef
            ReturnField = "ID"
            'ONLY WORKS FOR SINGLE RECORDS:
            'test for IsUpdate: Search Database first for Delivery Date and Reference:
            FoundRecord(TAbleIDX) = Find_myQuery(frmMainGIForm.myConnString, DBTables(TAbleIDX), SearchField, SearchValue, "STRING",
                                                          ReturnField, RecID(TAbleIDX), AllFieldValues, AllFieldnames, SearchCriteria)
            If FoundRecord(TAbleIDX) Then
                IsUpdate(TAbleIDX) = True
                UpdateCriteria(TAbleIDX) = "ID = " & RecID(TAbleIDX)
            Else
                IsUpdate(TAbleIDX) = False
                UpdateCriteria(TAbleIDX) = ""
            End If
        Next

        ctrlProperty = New clsControls
        TotalProperty = New clsTotals
        'Need to find the Delivery Reference selected first:
        For Each VarControl In dic_Controls
            If VarControl Is Nothing Then
                Continue For
            End If

            'VarControl = DeliveryDate & "_" & DeliveryRef & "_" & TagName
            ctrlProperty = dic_Controls(VarControl)

            If ctrlProperty Is Nothing Then
                Continue For
            End If
            ControlDeliveryDate = ctrlProperty.ControlDeliveryDate
            ControlDeliveryRef = ctrlProperty.ControlDeliveryRef
            'ctrlProperty.ControlLastSaved = Now() 'does not seem to work here.


            If Not UCase(CStr(ControlDeliveryDate)) = UCase(DeliveryDate) Then
                Continue For
            End If

            TotalProperty = dic_Totals(ControlDeliveryDate.ToString("dd/MM/yyyy") & "_" & ControlDeliveryRef)
            OpTotalRows = TotalProperty.Total_Operatives
            TotalShorts = TotalProperty.Total_ShortParts
            TotalExtras = TotalProperty.Total_ExtraParts

            LabourComments = ""

            If UCase(ControlDeliveryDate.ToString("yyyy-MM-dd")) = UCase(CDate(DeliveryDate).ToString("yyyy-MM-dd")) Then
                If UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "COMBOBOX" Or UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "TEXTBOX" Then
                    ControlFieldname = ctrlProperty.ControlFieldName

                    LabourComments = ctrlProperty.ControlLabour_Comments

                    ctrlProperty.ControlLastSaved = Now()
                    dtLastSavedDate = ctrlProperty.ControlLastSaved
                    TagName = ctrlProperty.ControlTAG
                    ControlID = ctrlProperty.ControlID 'DeliveryDate and DeliveryRef and TAG NUMBER - for actual ctrlCollection KEY.
                    ControlName = ctrlProperty.ControlName
                    ControlLowerTag = ctrlProperty.ControlStartTAG
                    ControlUpperTag = ctrlProperty.ControlEndTAG
                    ControlDBTable = ctrlProperty.ControlDBTable

                    If IsDate(ctrlProperty.ControlValue) Then
                        ControlValue = CStr(ctrlProperty.ControlValue)
                    Else
                        ControlValue = ctrlProperty.ControlValue
                    End If


                    'TEST HERE IF ANY BAD INPUT FROM USER

                    If Len(ControlFieldname) = 0 Then
                        'GET it from the Lookup table ! - will need some form of ID though - pass the Control Name:
                        'Messages = Messages & vbCrLf & " No Fieldname for : " & ControlName
                        Continue For
                    End If

                    'OK WARNING - THE COLLECTION MAY CONTAIN OTHER KEYS NOT ASSOCIATED WITH THE CURRENT DELIVERY DATE OR REFERENCE.
                    strUsername = ""
                    strEmpNo = ""
                    strName = ""
                    If UCase(ControlDBTable) = UCase(DBTables(1)) Then
                        'THIS WILL BE AN UPDATE:
                        strUsername = ctrlProperty.ControlUpdatedByUsername
                        strEmpNo = ctrlProperty.ControlUpdatedByEmpNo 'NOT BEING SAVED in tblDeliveryInfo here.
                        strName = ctrlProperty.ControlUpdatedByName
                        If UCase(ControlFieldname) = UCase("Last_Saved") Then
                            ControlValue = CStr(ctrlProperty.ControlLastSaved)
                        End If
                        If Len(Fieldnames(1)) = 0 Then
                            Fieldnames(1) = "Total_Operatives,TotalShortRows,TotalExtraRows,UpdatedByUsername,UpdatedbyEmpNo,UpdatedByName,DeliveryDate"
                            FieldValues(1) = OpTotalRows & "," & TotalShorts & "," & TotalExtras & "," & strUsername & "," & strEmpNo & "," & strName & "," & ControlValue
                        Else
                            If UCase(ControlFieldname) = UCase("UpdatedByEmpNo") Then
                                'Dont do anything
                            Else
                                Fieldnames(1) = Fieldnames(1) & "," & ControlFieldname
                                FieldValues(1) = FieldValues(1) & "," & ControlValue
                            End If

                        End If
                        'If Len(FieldValues(1)) = 0 Then
                        'FieldValues(1) = Chr(34) & strUsername & Chr(34) & "," & Str(Chr(34) & strEmpNo & Chr(34) & "," & Chr(34) & strName & Chr(34)
                        'FieldValues(1) = strUsername & "," & strEmpNo & "," & strName & "," & ControlValue
                        'FieldValues(1) = FieldValues(1) & "," & ControlValue
                        'Else
                        'FieldValues(1) = FieldValues(1) & "," & ControlValue
                        'End If
                    End If

                    'SAVE tblSupplierCompliance:
                    If UCase(ControlDBTable) = UCase(DBTables(2)) Then
                        'STATIC DATA - NOW ROWS. WILL BE INSERT TO START WITH - THEN USER MAY CHANGE THE NOs TO YES AND REMOVE THE COMMENTS.
                        'Save To SUPPLIER COMPLIANCE TABLE
                        dtLastSaved = Now()
                        strLastSaved = dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss")
                        dtDeliveryDate = ControlDeliveryDate
                        If Len(Fieldnames(2)) = 0 Then
                            Fieldnames(2) = "DeliveryDate,DeliveryReference,UpdatedByUsername,UpdatedByEmpNo,UpdatedByName,LastSaved" & "," & ControlFieldname
                        Else
                            Fieldnames(2) = Fieldnames(2) & "," & ControlFieldname
                        End If
                        If Len(FieldValues(2)) = 0 Then
                            FieldValues(2) = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss") & "," & ControlDeliveryRef
                            FieldValues(2) = FieldValues(2) & "," & ctrlProperty.ControlUpdatedByUsername & "," & ctrlProperty.ControlUpdatedByEmpNo & "," & ctrlProperty.ControlUpdatedByName
                            FieldValues(2) = FieldValues(2) & "," & strLastSaved & "," & ControlValue
                        Else
                            FieldValues(2) = FieldValues(2) & "," & ControlValue
                        End If

                    End If

                    'SAVE tblLabourHours:
                    If UCase(ControlDBTable) = UCase(DBTables(3)) Then
                        'STATIC DATA - INSERT first and then maybe UPDATE later is needed.
                        'CHECK if RECORD EXISTS FIRST.
                        'How / WHERE do we get TOTAL_ROWS from ????
                        'use tblDeliveryInfo ? or calculate total number of rows in the tblOperatives table ?
                        'OR use clsTotals.TotalOperatives like it was designed for ???
                        'Seems to be adding 1 to the TotalRows - NOT NEEDED !
                        'tblLabourhours
                        dtLastSaved = Now()
                        strLastSaved = dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss")
                        TotalControlsPerRow = ControlsPerRow(4) ' for tblOperatives:
                        ValueProperty = dic_Controls(DeliveryDate & "_" & DeliveryRef & "_" & "40")
                        If ValueProperty IsNot Nothing Then
                            ControlLowerTag = ValueProperty.ControlStartTAG
                            ControlUpperTag = ValueProperty.ControlEndTAG
                            VTTAG = 400
                            OperativeFrame = GetFrameControl(FormName & FormID, "Frame_Operatives")
                            'does this work ? - yes if last param=true
                            TotalFrameRows = GetTotalFrameRows(OperativeFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(4), VTTAG, True)
                            TotalRows = TotalProperty.Total_Operatives
                            If TagName = "40" Then
                                'FLMName but insert the total rows into the correct field also:
                                ValueProperty = dic_Controls(CStr(DeliveryDate) & "_" & DeliveryRef & "_" & TagName)
                                ControlValue = ValueProperty.ControlValue
                            End If
                            If TagName = "441" Then
                                ValueProperty = dic_Controls(CStr(DeliveryDate) & "_" & DeliveryRef & "_" & TagName)
                                FLMStartTime = CStr(ValueProperty.ControlFLMStartDateTime)
                                ControlValue = FLMStartTime
                            End If
                            If TagName = "442" Then
                                ValueProperty = dic_Controls(CStr(DeliveryDate) & "_" & DeliveryRef & "_" & TagName)
                                FLMEndTime = CStr(ValueProperty.ControlFLMEndDateTime)
                                ControlValue = FLMEndTime
                            End If
                            If TagName = "443" Then
                                ValueProperty = dic_Controls(CStr(DeliveryDate) & "_" & DeliveryRef & "_" & TagName)
                                OpComments = CStr(ValueProperty.ControlValue)
                                ControlValue = OpComments
                            End If
                            If Len(Fieldnames(3)) = 0 Then
                                Fieldnames(3) = "Total_Rows,UpdatedByUsername,UpdatedbyEmpNo,UpdatedByName,LastSaved" & "," & ControlFieldname
                            Else
                                Fieldnames(3) = Fieldnames(3) & "," & ControlFieldname
                            End If
                            If Len(FieldValues(3)) = 0 Then
                                FieldValues(3) = CStr(TotalRows) & "," & strUsername & "," & strEmpNo & "," & strName & "," & strLastSaved & "," & ControlValue
                            Else
                                FieldValues(3) = FieldValues(3) & "," & ControlValue
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If Len(Messages) > 0 Then
            MsgBox("msg: " & Messages)
        End If
        'SaveOK(1) = InsertUpdateMyRecord(IsUpdate(1), frmMainGIForm.myConnString, DBTables(1), Fieldnames(1), FieldValues(1), ErrMessages(1),
        'UpdateCriteria(1), ExcludeFields(1), 0, True, True, True, False)
        SaveOK(1) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, IsUpdate(1), "", DBTables(1), Fieldnames(1), FieldValues(1),
            UpdateCriteria(1), ExcludeFields(1), ErrMessages(1), False, ",")
        'SaveOK(2) = InsertUpdateMyRecord(IsUpdate(2), frmMainGIForm.myConnString, DBTables(2), Fieldnames(2), FieldValues(2), ErrMessages(2),
        'UpdateCriteria(2), ExcludeFields(2), 0, True, True, True, False)
        SaveOK(2) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, IsUpdate(2), "", DBTables(2), Fieldnames(2), FieldValues(2),
            UpdateCriteria(2), ExcludeFields(2), ErrMessages(2), False, ",")

        'SaveOK(3) = InsertUpdateMyRecord(IsUpdate(3), frmMainGIForm.myConnString, DBTables(3), Fieldnames(3), FieldValues(3), ErrMessages(3),
        'UpdateCriteria(3), ExcludeFields(3), 0, True, True, True, False)
        SaveOK(3) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, IsUpdate(3), "", DBTables(3), Fieldnames(3), FieldValues(3),
            UpdateCriteria(3), ExcludeFields(3), ErrMessages(3), False, ",")


        'SaveOK(1) = InsertUpdateRecords_Using_Parameters(IsUpdate(1), RecID(1), AccessDBpath, DBTables(1), Fieldnames(1), FieldValues(1), UpdateCriteria(1),
        'ExcludeFields, ErrMessages(1), EncaseFields, ValueDelim)
        'SaveOK(2) = InsertUpdateRecords_Using_Parameters(IsUpdate(2), RecID(2), AccessDBpath, DBTables(2), Fieldnames(2), FieldValues(2), UpdateCriteria(2),
        'ExcludeFields, ErrMessages(2), EncaseFields, ValueDelim)
        'SaveOK(3) = InsertUpdateRecords_Using_Parameters(IsUpdate(3), RecID(3), AccessDBpath, DBTables(3), Fieldnames(3), FieldValues(3), UpdateCriteria(3),
        'ExcludeFields, ErrMessages(3), EncaseFields, ValueDelim)
        'SaveOK(4) = InsertUpdateRecords_Using_Parameters(IsUpdate(4), RecID(4), AccessDBpath, DBTables(4), Fieldnames(4), FieldValues(4), UpdateCriteria(4), _
        'ExcludeFields, ErrMessages(4), EncaseFields, ValueDelim)
        'SaveOK(5) = InsertUpdateRecords_Using_Parameters(IsUpdate(5), RecID(5), AccessDBpath, DBTables(5), Fieldnames(5), FieldValues(5), UpdateCriteria(5), _
        'ExcludeFields, ErrMessages(5), EncaseFields, ValueDelim)

        SaveOK(4) = Save_Dynamic_Controls2(DeliveryDate, DeliveryRef, RecID)

        If SaveOK(1) = True And SaveOK(2) = True And SaveOK(3) = True Then
            SaveAllControls = True
            If SaveOK(4) Then
                dic_AnyChanges(FormID) = "NO"
                LockSavedOK = Change_LockSTATUS(DeliveryRef, "LOCKED", LockError)
            End If
            If SaveOK(4) Then
                FLMCompleteMessage = ""
                OpCompleteMessage = ""
                Call Check_For_Completion(DeliveryDate, DeliveryRef, True)

            End If
        Else
            SaveAllControls = False
        End If


    End Function

    Function Get_RecordIDs(strDeliveryDate As String, strDeliveryRef As String, LowerTagNumber As Long, TableName As String, strFrameControl As String,
                           VTTAG As Long, TotalControlsPerRow As Long, ByRef Dic_RecIDs As Object, ByRef IsUpdate() As Boolean,
                           Optional Dic_UpdateCriteria As Object = Nothing,
                           Optional ByVal ExtraCriteria As String = "",
                           Optional ByVal FormName As String = "GI_TIMESHEET_") As Long
        Dim ctrlProperty As clsControls
        Dim RowIDX As Long
        Dim FrameRowNum As Long
        Dim TotalOpFrameRows As Long
        Dim FrameControl As ScrollableControl
        Dim ControlLowerTAG As Long
        Dim ControlUpperTAG As Long
        Dim TotalFrameRows As Long
        Dim EncaseFields As Boolean = False
        Dim ErrMessages As String
        Dim ValueDelim As String = ","
        Dim ExcludeFields As String
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim FoundRecord As Boolean = False
        Dim RecID As String
        Dim AllFieldValues() As Object
        Dim AllFieldnames() As String
        Dim UpdateIDX As Long
        Dim SortField As String
        Dim Reversed As Boolean = False
        Dim SearchCriteria As String

        ReDim AllFieldnames(1)
        ReDim AllFieldValues(1)
        ReDim IsUpdate(1)
        Dic_RecIDs = CreateObject("Scripting.Dictionary")
        Dic_RecIDs.comparemode = vbTextCompare
        Dic_RecIDs.removeall

        Get_RecordIDs = 0

        ctrlProperty = dic_Controls(strDeliveryDate & "_" & strDeliveryRef & "_" & LowerTagNumber)
        'SearchSHORTCriteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
        'SearchEXTRACriteria = "ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
        If ctrlProperty Is Nothing Then
            'No Rows Exist - the control has not been created in Dic_Controls()
            Exit Function
        Else
            ControlLowerTAG = ctrlProperty.ControlStartTAG
            ControlUpperTAG = ctrlProperty.ControlEndTAG
            FrameControl = GetFrameControl(FormName & FormID, strFrameControl)
            'TotalFrameRows = GetTotalFrameRows(FrameControl, ControlLowerTAG, ControlUpperTAG, TotalControlsPerRow, VTTAG, True)
            TotalFrameRows = Get_TotalRows(TableName, strDeliveryRef, ExtraCriteria)
            RowIDX = 1
            UpdateIDX = 0
            Do While RowIDX <= TotalFrameRows
                EncaseFields = False
                ErrMessages = ""
                ExcludeFields = ""
                SearchCriteria = "FrameRowNumber = " & RowIDX
                If Len(ExtraCriteria) > 0 Then
                    SearchCriteria = SearchCriteria & " AND " & ExtraCriteria
                End If
                SearchField = "DeliveryReference"
                SearchValue = strDeliveryRef
                ReturnField = "ID"
                RecID = ""
                SortField = ""
                FoundRecord = Find_myQuery(frmMainGIForm.myConnString, TableName, SearchField, SearchValue, "STRING",
                                            ReturnField, RecID, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)
                If FoundRecord Then
                    IsUpdate(UpdateIDX) = True
                    Dic_UpdateCriteria(TableName & "_" & CStr(RowIDX)) = "ID = " & RecID
                    Dic_RecIDs(TableName & "_" & CStr(RowIDX)) = RecID
                Else
                    IsUpdate(UpdateIDX) = False
                    Dic_UpdateCriteria(TableName & "_" & CStr(RowIDX)) = ""
                    Dic_RecIDs(TableName & "_" & CStr(RowIDX)) = 0
                End If
                ReDim Preserve IsUpdate(UBound(IsUpdate) + 1)
                UpdateIDX = UpdateIDX + 1
                RowIDX = RowIDX + 1
            Loop

        End If

        Get_RecordIDs = TotalFrameRows

    End Function

    Function Save_TotalHours(strDeliveryDAte As String, strDeliveryRef As String, ByRef TimesheetHrs As Object,
                             Optional FormName As String = "GI_TIMESHEET_") As Boolean
        Dim dblTotalOpHours As Double
        Dim dblFLMHours As Double
        Dim Totals As clsTotals
        Dim TotalRows As Integer
        Dim FLMFrame As ScrollableControl
        Dim OpFrame As ScrollableControl
        Dim strTotalOpHours As String
        Dim InfoUpdateCriteria As String
        Dim LabourHoursUpdateCriteria As String
        Dim OpUpdateCriteria As String
        Dim DeliveryInfoRecID As Long
        Dim LabourHrsRecID As Long
        Dim OpRecID As Long
        Dim InfoSaveOK As Boolean
        Dim OpSaveOK As Boolean
        Dim LabHoursSaveOK As Boolean
        Dim FoundInfoRecord As Boolean
        Dim FoundLabourHrsRecord As Boolean
        Dim TableName As String
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim AllFieldValues() As Object
        Dim AllFieldnames() As String
        Dim SearchCriteria As String = ""
        Dim SortField As String = ""
        Dim Reversed As Boolean = False
        Dim ErrMessages As String = ""
        Dim LabFieldNames As String = ""
        Dim LabFieldValues As String = ""
        Dim InfoFieldNames As String = ""
        Dim InfoFieldValues As String = ""
        Dim ExcludeFields As String = ""
        Dim NeedsUpdate As Boolean = False
        Dim LabourHoursTable As String = "tblLabourHours"
        Dim InfoTable As String = "tblDeliveryInfo"

        Save_TotalHours = False
        strDeliveryDAte = CDate(strDeliveryDAte).ToString("dd/MM/yyyy")
        dblFLMHours = 0.0F
        dblTotalOpHours = 0.0F

        FLMFrame = GetFrameControl(FormName & FormID, "Frame_FLMDetails")
        OpFrame = GetFrameControl(FormName & FormID, "Frame_Operatives")

        dblFLMHours = Module_TIMESHEET.Get_TotalHours(FLMFrame, strDeliveryDAte, strDeliveryRef, TimesheetHrs)
        dblFLMHours = Decimal.Round(dblFLMHours, 2, MidpointRounding.AwayFromZero)
        dblTotalOpHours = Module_TIMESHEET.Get_TotalHours(OpFrame, strDeliveryDAte, strDeliveryRef, TimesheetHrs)
        dblTotalOpHours = Decimal.Round(dblTotalOpHours, 2, MidpointRounding.AwayFromZero)

        Totals = dic_Totals(strDeliveryDAte & "_" & strDeliveryRef)
        If IsNothing(Totals) Then Exit Function
        TotalRows = Totals.Total_Operatives
        If UBound(TimesheetHrs) >= TotalRows + 1 Then
            strTotalOpHours = CStr(TimesheetHrs(TotalRows + 1))
        End If
        'Need record ID for DeliveryInfo table from Current:
        TableName = "tblDeliveryInfo"
        SearchField = "DeliveryReference"
        SearchValue = strDeliveryRef
        ReturnField = "ID"
        SortField = "DeliveryDate"
        Reversed = True

        FoundInfoRecord = Find_myQuery(frmMainGIForm.myConnString, InfoTable, SearchField, SearchValue, "STRING",
                                            ReturnField, DeliveryInfoRecID, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)

        TableName = "tblLabourHours"
        SearchField = "DeliveryReference"
        SearchValue = strDeliveryRef
        ReturnField = "ID"

        FoundLabourHrsRecord = Find_myQuery(frmMainGIForm.myConnString, LabourHoursTable, SearchField, SearchValue, "STRING",
                                            ReturnField, LabourHrsRecID, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)
        InfoUpdateCriteria = "ID = " & DeliveryInfoRecID
        LabourHoursUpdateCriteria = "ID =" & LabourHrsRecID
        InfoFieldNames = "Total_Hours"
        InfoFieldValues = dblTotalOpHours
        If FoundInfoRecord Then
            NeedsUpdate = True
            InfoSaveOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate, "", InfoTable, InfoFieldNames, InfoFieldValues,
                    InfoUpdateCriteria, ExcludeFields, ErrMessages, False, ",")
        Else
            NeedsUpdate = False
        End If
        'tblLabourHours DB TABLE:
        LabFieldNames = "Total_FLM_Hours,Total_Labour_Hours"
        LabFieldValues = dblFLMHours & "," & dblTotalOpHours
        If FoundLabourHrsRecord = True Then
            LabHoursSaveOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate, "", LabourHoursTable, LabFieldNames, LabFieldValues,
                    LabourHoursUpdateCriteria, ExcludeFields, ErrMessages, False, ",")
        Else
            MsgBox("Save Total Hours: Could not find record")
            Exit Function
        End If
        If LabHoursSaveOK = False Then
            MsgBox("Save Total Hours: Did not save")
            Exit Function
        End If
        Save_TotalHours = LabHoursSaveOK

    End Function

    Function Save_Dynamic_Controls2(DeliveryDAte As String, DeliveryRef As String, RecordID As Object,
                                    Optional ByVal FormName As String = "GI_TIMESHEET_") As Boolean
        Dim OperativeFrame As ScrollableControl
        Dim ShortsFrame As ScrollableControl
        Dim ExtraFrame As ScrollableControl
        Dim ctrlProperty As clsControls
        Dim ValueProperty As clsControls
        Dim Totals As clsTotals
        Dim ControlDeliveryDate As Date
        Dim ControlDeliveryRef As String
        Dim DBTAbles() As String
        Dim ControlsPerRow() As Long
        Dim ControlFieldname As String
        Dim TagName As String
        Dim ControlID As String
        Dim ControlName As String
        Dim ControlLowerTag As String
        Dim ControlUpperTag As String
        Dim ControlDBTable As String
        Dim ControlValue As String
        Dim TAbleIDX As Long
        Dim Dic_SaveUs As Object
        Dim Messages As String = ""
        Dim TotalControlsPerRow As Long
        Dim VTTAG As Long
        Dim RowNumber As Long
        Dim TotalOPFrameRows As Long
        Dim TotalShortFrameRows As Long
        Dim TotalExtraFrameRows As Long
        Dim FrameValue As String
        Dim ShortOrExtra As String
        Dim SaveOK() As Boolean
        Dim ErrMessages() As String
        Dim UpdateCriteria() As String
        Dim IsUpdate() As Boolean
        Dim FoundRecord() As Boolean
        Dim RecID() As String
        Dim Dic_RecIDs As Object
        Dim Dic_Criteria As Object
        Dim ExcludeFields() As String
        Dim ValueArray() As String
        Dim RowIDX As Long
        Dim OperativeName As String
        Dim SplitRowArr() As String
        Dim Activity As String
        Dim ExtendCriteria As String
        Dim EncaseFields As Boolean
        Dim ValueDelim As String
        Dim SearchCriteria As String
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim AllFieldValues() As Object = Nothing
        Dim AllFieldnames() As String = Nothing
        Dim Fieldnames() As String
        Dim FieldValues() As String
        Dim FieldIDX As Long
        Dim strPartNo As String
        Dim strQty As String
        Dim Fieldnames2(,) As String
        Dim FieldValues2(,) As String
        Dim CountIterations As Long
        Dim ErrMessage As String
        Dim TimeWaste As Long
        Dim RowNum As Long
        Dim NeedsUpdate() As Boolean
        Dim UpdatedByUsername As String
        Dim UpdatedByEmpNo As String
        Dim UpdatedByName As String
        Dim dtLastSaved As DateTime
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim SearchSHORTCriteria As String
        Dim SearchEXTRACriteria As String
        Dim DelimChar As String = "|"
        Dim TimeSheetHours() As Object

        Save_Dynamic_Controls2 = False
        Dic_SaveUs = CreateObject("Scripting.Dictionary")
        Dic_SaveUs.CompareMode = vbTextCompare
        Dic_SaveUs.removeall
        Dic_RecIDs = CreateObject("Scripting.Dictionary")
        Dic_RecIDs.comparemode = vbTextCompare
        Dic_RecIDs.removeall
        Dic_Criteria = CreateObject("Scripting.Dictionary")
        Dic_Criteria.comparemode = vbTextCompare
        Dic_Criteria.removeall

        ReDim DBTAbles(6)
        ReDim ControlsPerRow(6)
        DeliveryDAte = Replace(DeliveryDAte, "-", "/")

        strDeliveryDate = CDate(DeliveryDAte).ToString("dd/MM/yyyy")
        strDeliveryRef = DeliveryRef

        ReDim SaveOK(6)
        ReDim ErrMessages(6)
        ReDim UpdateCriteria(6)
        ReDim IsUpdate(6)
        ReDim FoundRecord(6)
        ReDim RecID(6)
        ReDim ControlsPerRow(6)
        ReDim ExcludeFields(6)
        ReDim ValueArray(50)
        ReDim Fieldnames(6)
        ReDim FieldValues(6)
        ReDim Fieldnames2(6, 20) '20 fields.
        ReDim FieldValues2(6, 20)
        ReDim NeedsUpdate(1)

        RecID(1) = RecordID(1)
        RecID(2) = RecordID(2)
        RecID(3) = RecordID(3)
        RecID(4) = RecordID(4)
        RecID(5) = RecordID(5)

        DBTAbles(1) = "tblDeliveryInfo"
        DBTAbles(2) = "tblSupplierCompliance"
        DBTAbles(3) = "tblLabourHours"
        DBTAbles(4) = "tblOperatives"
        DBTAbles(5) = "tblShortsAndExtraParts"
        DBTAbles(6) = "tblShortsAndExtraParts"

        ControlsPerRow(1) = 1
        ControlsPerRow(2) = 1
        ControlsPerRow(3) = 1
        ControlsPerRow(4) = 6
        ControlsPerRow(5) = 3
        ControlsPerRow(6) = 3

        OperativeFrame = GetFrameControl(FormName & FormID, "Frame_Operatives")
        CountIterations = 0

        'READ IN FROM tblLabourHours to get total rows:
        ' - Then loop for each dynamic table - tblOperatives ; tblShortsAndExtraRows - Short and Extra criteria
        '   to gather each RECORD ID (if any) to UPDATE to.


        'frmMainGIForm.InsertValueIntoForm(FormName & FormID, "txtMEssages", "Error during save: key not recognised: " & DeliveryDAte & "_" & DeliveryRef & "_" & "43")
        'Exit Function

        '*******************************************************************************************************************************
        'Seems NO CRITERIA SPECIFIED for the Delivery Reference ???????????
        '*******************************************************************************************************************************

        For Each VarControl In dic_Controls
            If VarControl Is Nothing Then
                Continue For
            End If

            'VarControl = DeliveryDate & "_" & DeliveryRef & "_" & TagName
            ctrlProperty = dic_Controls(VarControl)

            If ctrlProperty Is Nothing Then
                'Default DATE:
                'ControlDeliveryDate = CDate("1970-01-01")
                Continue For
                'ControlDeliveryRef = ctrlProperty.ControlDeliveryRef
            Else
                ControlDeliveryDate = ctrlProperty.ControlDeliveryDate
                ControlDeliveryRef = ctrlProperty.ControlDeliveryRef

            End If

            ctrlProperty.ControlLastSaved = Now()
            ControlFieldname = ctrlProperty.ControlFieldName

            If Not UCase(CStr(ControlDeliveryDate)) = DeliveryDAte Then
                Continue For
            End If

            If UCase(CStr(ControlDeliveryDate)) = UCase(DeliveryDAte) Then
                If UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "COMBOBOX" Or
                        UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "TEXTBOX" Then

                    ControlFieldname = ctrlProperty.ControlFieldName

                    TagName = ctrlProperty.ControlTAG
                    If IsNumeric(TagName) Then
                        If CLng(TagName) < 33 Then
                            Continue For
                        End If
                    End If
                    ControlID = ctrlProperty.ControlID 'DeliveryDate and DeliveryRef and TAG NUMBER - for actual ctrlCollection KEY.
                    ControlName = ctrlProperty.ControlName
                    ControlLowerTag = ctrlProperty.ControlStartTAG
                    ControlUpperTag = ctrlProperty.ControlEndTAG
                    ControlDBTable = ctrlProperty.ControlDBTable
                    UpdatedByEmpNo = ctrlProperty.ControlUpdatedByEmpNo
                    UpdatedByUsername = ctrlProperty.ControlUpdatedByUsername
                    UpdatedByName = ctrlProperty.ControlUpdatedByName
                    dtLastSaved = Now()

                    ControlValue = ctrlProperty.ControlValue
                    'TEST HERE IF ANY BAD INPUT FROM USER

                    If Len(ControlFieldname) = 0 Then
                        'GET it from the Lookup table ! - will need some form of ID though - pass the Control Name:
                        Messages = Messages & vbCrLf & " No Fieldname for : " & ControlName
                        Continue For
                    End If

                    If UCase(ControlDBTable) = UCase(DBTAbles(4)) Then
                        'NEED TO LOOP THROUGH EACH ROW IN THE FRAME_OPERATIVES CONTROLS HERE AND SAVE.
                        'THIS WILL MORE LIKELY BE AN INSERT - BUT CAN ALSO BE UPDATE - IF USER IS CHANGING ANY INFO WITHIN THE ROWS.
                        '*** tblOperatives ***
                        TAbleIDX = 4
                        VTTAG = 400

                        RowNumber = Get_NumericPartOfString(ControlName)
                        'ctrlProperty.ControlRowNumber = RowNumber
                        'ControlLowerTag = set at start of loop for the current control ...

                        TotalOPFrameRows = GetTotalFrameRows(OperativeFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(4), VTTAG, True)
                        'ok so we have total frames and the highest tag now.
                        'ctrlProperty.ControlStartTAG = ControlLowerTag
                        'ctrlProperty.ControlEndTAG = ControlUpperTag

                        'Turn all the values into fields and values to save:

                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) = CStr(RecID(3))
                            Fieldnames2(TAbleIDX, 1) = "LINKID"
                            FieldValues2(TAbleIDX, 1) = CStr(RecID(3))

                        Else
                            'ControlValue = DeliveryRef
                        End If

                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) = DeliveryDAte
                            Fieldnames2(TAbleIDX, 2) = "DELIVERYDATE"
                            FieldValues2(TAbleIDX, 2) = DeliveryDAte
                        Else
                            'ControlValue = DeliveryDate
                        End If

                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) = DeliveryRef
                            Fieldnames2(TAbleIDX, 3) = "DELIVERYREFERENCE"
                            FieldValues2(TAbleIDX, 3) = DeliveryRef
                        Else
                            'ControlValue = DeliveryRef
                        End If

                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) = CStr(RowNumber)
                            Fieldnames2(TAbleIDX, 4) = "FRAMEROWNUMBER"
                            FieldValues2(TAbleIDX, 4) = CStr(RowNumber)
                        Else
                            'ControlValue = DeliveryDate
                        End If

                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) = CStr(TotalOPFrameRows)

                        Else
                            'ControlValue = DeliveryDate
                        End If

                        If UCase(ControlFieldname) = "OP_NAME" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_NAME|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_NAME|" & CStr(RowNumber)) = ControlValue
                                Fieldnames2(TAbleIDX, 5) = ControlFieldname
                                FieldValues2(TAbleIDX, 5) = ControlValue
                            Else
                                'repeated.
                            End If
                            'FieldOpControls(4, RowIDX) = UCase(ControlFieldname)
                            'ValueOpControls(4, RowIDX) = UCase(ControlValue)
                        End If
                        If UCase(ControlFieldname) = "ACTIVITY_NAME" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "ACTIVITY_NAME|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "ACTIVITY_NAME|" & CStr(RowNumber)) = ControlValue
                                Fieldnames2(TAbleIDX, 6) = ControlFieldname
                                FieldValues2(TAbleIDX, 6) = ControlValue
                            Else
                                'repeated:
                            End If
                        End If
                        If UCase(ControlFieldname) = "OP_STARTTIME" Then
                            ValueProperty = dic_Controls(DeliveryDAte & "_" & DeliveryRef & "_" & TagName)
                            If ValueProperty IsNot Nothing Then
                                ControlValue = ValueProperty.ControlOpStartDateTime
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_STARTTIME|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_STARTTIME|" & CStr(RowNumber)) = ControlValue
                                    Fieldnames2(TAbleIDX, 7) = ControlFieldname
                                    FieldValues2(TAbleIDX, 7) = ControlValue
                                Else
                                    'Repeated.
                                End If
                            Else
                                'MsgBox("Could not find KEY for Start Time")
                                ErrMessage = "Could not find KEY for Start Time"
                                frmMainGIForm.txtMessages.Text = ErrMessage
                                frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, ErrMessage,
                                    "Save_Dynamic_Controls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                            End If
                        End If
                        If UCase(ControlFieldname) = "OP_FINISHTIME" Then
                            ValueProperty = dic_Controls(DeliveryDAte & "_" & DeliveryRef & "_" & TagName)
                            If ValueProperty IsNot Nothing Then
                                ControlValue = ValueProperty.ControlOpEndDateTime
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_FINISHTIME|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OP_FINISHTIME|" & CStr(RowNumber)) = ControlValue
                                    Fieldnames2(TAbleIDX, 8) = ControlFieldname
                                    FieldValues2(TAbleIDX, 8) = ControlValue
                                Else
                                    'Repeated.
                                End If
                            Else
                                'MsgBox("Could not find key for Finish Time")
                                ErrMessage = "Could not find KEY for Finish Time"
                                frmMainGIForm.txtMessages.Text = ErrMessage
                                frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, ErrMessage, "Save_Dynamic_Controls()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                            End If
                        End If
                        If UCase(ControlFieldname) = "OPCOMMENTS" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OPCOMMENTS|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "OPCOMMENTS|" & CStr(RowNumber)) = ControlValue
                                Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                FieldValues2(TAbleIDX, 9) = ControlValue
                            Else
                                'repeated:
                            End If
                        End If
                        'If UCase(ControlFieldname) = "UPDATEDBYUSERNAME" Then
                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) = UpdatedByUsername
                            Fieldnames2(TAbleIDX, 9) = ControlFieldname
                            FieldValues2(TAbleIDX, 9) = ControlValue
                        Else
                            'repeated:
                        End If
                        'End If
                        'If UCase(ControlFieldname) = "UPDATEDBYEMPNO" Then
                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) = UpdatedByEmpNo
                            Fieldnames2(TAbleIDX, 9) = ControlFieldname
                            FieldValues2(TAbleIDX, 9) = ControlValue
                        Else
                            'repeated:
                        End If
                        'End If
                        'If UCase(ControlFieldname) = "UPDATEDBYNAME" Then
                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) Then
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) = UpdatedByName
                            Fieldnames2(TAbleIDX, 9) = ControlFieldname
                            FieldValues2(TAbleIDX, 9) = ControlValue
                        Else
                            'repeated:
                        End If
                        'End If
                        'If UCase(ControlFieldname) = "LASTSAVED" Then
                        If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) Then
                            dtLastSaved = Now()
                            Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) = dtLastSaved
                            Fieldnames2(TAbleIDX, 9) = ControlFieldname
                            FieldValues2(TAbleIDX, 9) = ControlValue
                        Else
                            'repeated:
                        End If
                        'End If
                    End If

                    If UCase(ControlDBTable) = UCase(DBTAbles(5)) Then
                        'AGAIN NEED TO LOOP THROUGH EACH ROW OF CONTROLS IN FRAME_SHORTSANDEXTRAS AND SAVE EACH ROW:
                        TAbleIDX = 5
                        TotalControlsPerRow = 3
                        '*** tblShortAndExtraParts ***
                        RowNumber = Get_NumericPartOfString(ControlName)
                        ShortsFrame = GetFrameControl(FormName & FormID, "Frame_Short_Parts")
                        ExtraFrame = GetFrameControl(FormName & FormID, "Frame_Extra_Parts")
                        'ValueProperty.ControlRowNumber = RowNumber

                        If IsNumeric(TagName) Then
                            If CLng(TagName) > 1000 And CLng(TagName) < 2000 Then
                                ShortOrExtra = "SHORT"
                                VTTAG = 1000
                                TAbleIDX = 5
                                'TotalShortFrameRows = GetTotalFrameRows(ShortsFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(5), VTTAG, True)
                                'INSERT THE DELIVERY REFERENCE INTO THE KEY ALSO:
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) = CStr(RecID(1)) 'LINK BACK to tblDeliveryInfo record.
                                    Fieldnames2(TAbleIDX, 1) = "LINKID"
                                    FieldValues2(TAbleIDX, 1) = CStr(RecID(1))
                                Else
                                    'ControlValue = DeliveryRef
                                End If

                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) = DeliveryDAte
                                    Fieldnames2(TAbleIDX, 2) = "DELIVERYDATE"
                                    FieldValues2(TAbleIDX, 2) = DeliveryDAte
                                Else
                                    'ControlValue = DeliveryDate
                                End If

                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) = DeliveryRef
                                    Fieldnames2(TAbleIDX, 3) = "DELIVERYREFERENCE"
                                    FieldValues2(TAbleIDX, 3) = DeliveryRef
                                Else
                                    'ControlValue = DeliveryRef
                                End If

                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) = CStr(RowNumber)
                                    Fieldnames2(TAbleIDX, 4) = "FRAMEROWNUMBER"
                                    FieldValues2(TAbleIDX, 4) = CStr(RowNumber)
                                Else
                                    'ControlValue = DeliveryDate
                                End If

                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) = CStr(TotalShortFrameRows)

                                Else
                                    'ControlValue = DeliveryDate
                                End If

                                If UCase(ControlFieldname) = "PARTNO" Then
                                    If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "PARTNO|" & CStr(RowNumber)) Then
                                        Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "PARTNO|" & CStr(RowNumber)) = ControlValue
                                        Fieldnames2(TAbleIDX, 5) = ControlFieldname
                                        FieldValues2(TAbleIDX, 5) = ControlValue
                                    Else
                                        'Repeated
                                    End If
                                    'FieldOpControls(5, RowIDX) = UCase(ControlFieldname)
                                    'ValueOpControls(5, RowIDX) = UCase(ControlValue)
                                End If

                                If UCase(ControlFieldname) = "QTY" Then
                                    If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "QTY|" & CStr(RowNumber)) Then
                                        Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "QTY|" & CStr(RowNumber)) = ControlValue
                                        Fieldnames2(TAbleIDX, 6) = ControlFieldname
                                        FieldValues2(TAbleIDX, 6) = ControlValue
                                    Else
                                        'Repeated
                                    End If
                                    'FieldOpControls(5, RowIDX) = UCase(ControlFieldname)
                                    'ValueOpControls(5, RowIDX) = UCase(ControlValue)
                                End If

                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "SHORTOREXTRA|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "SHORTOREXTRA|" & CStr(RowNumber)) = ShortOrExtra
                                    Fieldnames2(TAbleIDX, 7) = ControlFieldname
                                    FieldValues2(TAbleIDX, 7) = ControlValue
                                Else
                                    'repeated.
                                End If

                                'If UCase(ControlFieldname) = "UPDATEDBYUSERNAME" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) = UpdatedByUsername
                                    Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                    FieldValues2(TAbleIDX, 9) = ControlValue
                                Else
                                    'repeated:
                                End If
                                'End If
                                'If UCase(ControlFieldname) = "UPDATEDBYEMPNO" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) = UpdatedByEmpNo
                                    Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                    FieldValues2(TAbleIDX, 9) = ControlValue
                                Else
                                    'repeated:
                                End If
                                'End If
                                'If UCase(ControlFieldname) = "UPDATEDBYNAME" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) = UpdatedByName
                                    Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                    FieldValues2(TAbleIDX, 9) = ControlValue
                                Else
                                    'repeated:
                                End If
                                'End If
                                'If UCase(ControlFieldname) = "LASTSAVED" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) Then
                                    dtLastSaved = Now()
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) = dtLastSaved
                                    Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                    FieldValues2(TAbleIDX, 9) = ControlValue
                                Else
                                    'repeated:
                                End If

                            End If
                        End If
                    End If

                    If UCase(ControlDBTable) = UCase(DBTAbles(6)) Then
                        TAbleIDX = 6
                        TotalControlsPerRow = 3
                        '*** tblShortAndExtraParts ***
                        RowNumber = Get_NumericPartOfString(ControlName)
                        ShortsFrame = GetFrameControl(FormName & FormID, "Frame_Short_Parts")
                        ExtraFrame = GetFrameControl(FormName & FormID, "Frame_Extra_Parts")

                        If CLng(TagName) > 2000 Then
                            ShortOrExtra = "EXTRA"
                            VTTAG = 2000

                            TotalExtraFrameRows = GetTotalFrameRows(ExtraFrame, ControlLowerTag, ControlUpperTag, ControlsPerRow(6), VTTAG, False)

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LINKID|" & CStr(RowNumber)) = CStr(RecID(1)) 'LINK BACK to tblDeliveryInfo record.
                                Fieldnames2(TAbleIDX, 1) = "LINKID"
                                FieldValues2(TAbleIDX, 1) = CStr(RecID(1))
                            Else
                                'ControlValue = DeliveryRef
                            End If

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYDATE|" & CStr(RowNumber)) = DeliveryDAte
                                Fieldnames2(TAbleIDX, 2) = "DELIVERYDATE"
                                FieldValues2(TAbleIDX, 2) = DeliveryDAte
                            Else
                                'ControlValue = DeliveryDate
                            End If

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "DELIVERYREFERENCE|" & CStr(RowNumber)) = DeliveryRef
                                Fieldnames2(TAbleIDX, 3) = "DELIVERYREFERENCE"
                                FieldValues2(TAbleIDX, 3) = DeliveryRef
                            Else
                                'ControlValue = DeliveryRef
                            End If

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "FRAMEROWNUMBER|" & CStr(RowNumber)) = CStr(RowNumber)
                                Fieldnames2(TAbleIDX, 4) = "FRAMEROWNUMBER"
                                FieldValues2(TAbleIDX, 4) = CStr(RowNumber)
                            Else
                                'ControlValue = DeliveryDate
                            End If

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "TOTAL_ROWS|" & CStr(RowNumber)) = CStr(TotalExtraFrameRows)

                            Else
                                'ControlValue = DeliveryDate
                            End If

                            If UCase(ControlFieldname) = "PARTNO" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "PARTNO|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "PARTNO|" & CStr(RowNumber)) = ControlValue
                                    Fieldnames2(TAbleIDX, 5) = ControlFieldname
                                    FieldValues2(TAbleIDX, 5) = ControlValue
                                Else
                                    'Repeated
                                End If
                            End If

                            If UCase(ControlFieldname) = "QTY" Then
                                If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "QTY|" & CStr(RowNumber)) Then
                                    Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "QTY|" & CStr(RowNumber)) = ControlValue
                                    Fieldnames2(TAbleIDX, 6) = ControlFieldname
                                    FieldValues2(TAbleIDX, 6) = ControlValue
                                Else
                                    'Repeated
                                End If
                            End If

                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "SHORTOREXTRA|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "SHORTOREXTRA|" & CStr(RowNumber)) = ShortOrExtra
                                Fieldnames2(TAbleIDX, 7) = ControlFieldname
                                FieldValues2(TAbleIDX, 7) = ControlValue
                            Else
                                'repeated.
                            End If

                            'If UCase(ControlFieldname) = "UPDATEDBYUSERNAME" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYUSERNAME|" & CStr(RowNumber)) = UpdatedByUsername
                                Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                FieldValues2(TAbleIDX, 9) = ControlValue
                            Else
                                'repeated:
                            End If
                            'End If
                            'If UCase(ControlFieldname) = "UPDATEDBYEMPNO" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYEMPNO|" & CStr(RowNumber)) = UpdatedByEmpNo
                                Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                FieldValues2(TAbleIDX, 9) = ControlValue
                            Else
                                'repeated:
                            End If
                            'End If
                            'If UCase(ControlFieldname) = "UPDATEDBYNAME" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) Then
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "UPDATEDBYNAME|" & CStr(RowNumber)) = UpdatedByName
                                Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                FieldValues2(TAbleIDX, 9) = ControlValue
                            Else
                                'repeated:
                            End If
                            'End If
                            'If UCase(ControlFieldname) = "LASTSAVED" Then
                            If Not Dic_SaveUs.Exists(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) Then
                                dtLastSaved = Now()
                                Dic_SaveUs(CStr(TAbleIDX) & "|" & DeliveryRef & "|" & "LASTSAVED|" & CStr(RowNumber)) = dtLastSaved
                                Fieldnames2(TAbleIDX, 9) = ControlFieldname
                                FieldValues2(TAbleIDX, 9) = ControlValue
                            Else
                                'repeated:
                            End If

                        End If
                    End If 'If UCase(ControlDBTable) = UCase(DBTAbles(6))

                End If 'If UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "COMBOBOX" Or
                'UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "TEXTBOX"
            Else
                Continue For
            End If 'If UCase(CStr(ControlDeliveryDate)) = UCase(DeliveryDAte)

            'Can be up to 30 iterations etc depending on number of rows
            CountIterations = CountIterations + 1
        Next

        frmMainGIForm.txtMessages.Text = "STILL SAVING ... PLEASE WAIT - Iterations:" & CStr(CountIterations)

        ReDim NeedsUpdate(1)
        TotalOPFrameRows = Get_RecordIDs(DeliveryDAte, DeliveryRef, 43, DBTAbles(4), "Frame_Operatives", 400, 6, Dic_RecIDs, NeedsUpdate,
                                         Dic_Criteria)
        RowIDX = 1
        Do While RowIDX <= TotalOPFrameRows
            Fieldnames(4) = ""
            FieldValues(4) = ""
            Call ConvertDic_ToInsertFields(DeliveryRef, 4, frmMainGIForm.myConnString, DBTAbles(4), Dic_SaveUs, RowIDX, Fieldnames(4), FieldValues(4))

            UpdateCriteria(4) = Dic_Criteria(DBTAbles(4) & "_" & CStr(RowIDX))
            SaveOK(4) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate(RowIDX - 1), "", DBTAbles(4), Fieldnames(4), FieldValues(4),
                        UpdateCriteria(4), ExcludeFields(4), ErrMessages(4), False, ",")
            RowIDX = RowIDX + 1
        Loop

        If TotalOPFrameRows = 0 Then
            SaveOK(4) = True
        End If

        'tblShortsAndExtraParts:

        VTTAG = 1000
        ReDim NeedsUpdate(1)
        SearchSHORTCriteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
        TotalShortFrameRows = Get_RecordIDs(DeliveryDAte, DeliveryRef, 1001, DBTAbles(5), "Frame_Shorts", VTTAG, 3, Dic_RecIDs, NeedsUpdate,
                                            Dic_Criteria, SearchSHORTCriteria)
        Totals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)

        'TotalShortFrameRows =
        RowIDX = 1
        Do While RowIDX <= TotalShortFrameRows
            Fieldnames(5) = ""
            FieldValues(5) = ""
            Call ConvertDic_ToInsertFields(DeliveryRef, 5, frmMainGIForm.myConnString, DBTAbles(5), Dic_SaveUs, RowIDX,
                                           Fieldnames(5), FieldValues(5))

            UpdateCriteria(5) = Dic_Criteria(DBTAbles(5) & "_" & CStr(RowIDX)) & " AND ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
            SaveOK(5) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate(RowIDX - 1), "", DBTAbles(5), Fieldnames(5), FieldValues(5),
                        UpdateCriteria(5), ExcludeFields(5), ErrMessages(5), False, ",")
            RowIDX = RowIDX + 1
        Loop
        If TotalShortFrameRows = 0 Then
            SaveOK(5) = True
        End If

        'tblExtras:
        VTTAG = 2000
        ReDim NeedsUpdate(1)
        SearchEXTRACriteria = "ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
        TotalExtraFrameRows = Get_RecordIDs(DeliveryDAte, DeliveryRef, 2001, DBTAbles(6), "Frame_Extra", VTTAG, 3, Dic_RecIDs, NeedsUpdate,
                                            Dic_Criteria, SearchEXTRACriteria)
        RowIDX = 1
        Do While RowIDX <= TotalExtraFrameRows
            Fieldnames(6) = ""
            FieldValues(6) = ""
            Call ConvertDic_ToInsertFields(DeliveryRef, 6, frmMainGIForm.myConnString, DBTAbles(6), Dic_SaveUs, RowIDX, Fieldnames(6), FieldValues(6))

            UpdateCriteria(6) = Dic_Criteria(DBTAbles(6) & "_" & CStr(RowIDX)) & " AND ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
            SaveOK(6) = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, NeedsUpdate(RowIDX - 1), "", DBTAbles(6), Fieldnames(6), FieldValues(6),
                        UpdateCriteria(6), ExcludeFields(6), ErrMessages(6), False, ",")
            RowIDX = RowIDX + 1
        Loop
        If TotalExtraFrameRows = 0 Then
            SaveOK(6) = True
        End If

        If SaveOK(4) = True And SaveOK(5) = True And SaveOK(6) = True Then
            Save_Dynamic_Controls2 = True
        Else
            Save_Dynamic_Controls2 = False
        End If

        Call Save_TotalHours(strDeliveryDate, strDeliveryRef, TimeSheetHours)

        'Iterarte through the new array created above to extract each row for the Operatives:
        ' AND SAVE after gathering the FIELDS and VALUES:

        'BUT we need the record ID of the parent record first ?????????????????? so the LINKID is populated.

        'OK so we have two loops then. The first to establish and save the STATIC control records and setup record IDs.
        'The second loop will go through and filter on just the CHILD TABLES and will have an inner loop that will detect when a row change happens.
    End Function

    Sub Write_Properties_To_File(PropertyLogFile As String, WriteProperties As clsControls)
        Dim Messages() As String
        Dim IDX As Integer

        If WriteProperties Is Nothing Then
            frmMainGIForm.logger.LogError(PropertyLogFile, Application.StartupPath, "clsControls passed is nothing", "Write_Properties_To_File()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED IN:" & frmMainGIForm.myUsername)
            Exit Sub
        End If

        ReDim Messages(11)
        Messages(0) = "1) Control Name= " & WriteProperties.ControlName
        Messages(1) = "2) Control ID= " & WriteProperties.ControlID
        Messages(2) = "3) Control Value= " & WriteProperties.ControlValue
        Messages(3) = "4) Control Date=" & CStr(WriteProperties.ControlDate)
        Messages(4) = "5) Control Frame no=" & CStr(WriteProperties.ControlFrameRowNumber)
        Messages(5) = "6) Control Primary Key=" & CStr(WriteProperties.ControlPrimaryKey)
        Messages(6) = "7) Control Op Name=" & WriteProperties.ControlOpName
        Messages(7) = "8) Control Op Activity=" & WriteProperties.ControlOpActivity
        Messages(8) = "9) Control Op Start Date=" & CStr(WriteProperties.ControlOpStartDateTime)
        Messages(9) = "10) Control Op End Date=" & CStr(WriteProperties.ControlOpEndDateTime)
        Messages(10) = "11) CONTROL FIELDNAME = " & WriteProperties.ControlFieldName
        Messages(11) = "12) LAST SAVED = " & CStr(WriteProperties.ControlLastSaved)

        For IDX = 0 To 11
            frmMainGIForm.logger.LogError(PropertyLogFile, Application.StartupPath, Messages(IDX), "Write_Properties_To_File()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED IN:" & frmMainGIForm.myUsername)
        Next

    End Sub

    Function Get_NumericPartOfString(TheString As String) As Long
        Dim IDX As Long
        Dim strPart As String
        Dim strNumber As String = ""
        Dim NumericPart As Long

        NumericPart = 0
        strPart = TheString
        IDX = 1
        Do While IDX <= Len(TheString)
            If IsNumeric(Mid(TheString, IDX, 1)) Then
                strNumber = strNumber & Mid(TheString, IDX, 1)
            End If
            IDX = IDX + 1
        Loop
        If Len(strNumber) > 0 Then
            NumericPart = CLng(strNumber)
        End If

        Get_NumericPartOfString = NumericPart

    End Function

    Sub ConvertDic_ToInsertFields(strDeliveryRef As String, TAbleIDX As Long, connString As String, DBTable As String, Dic_Values As Object, RowIDX As Long,
                                  ByRef FinalFields As String, ByRef FinalValues As String, Optional ByVal DelimChar As String = "|")
        Dim FieldIDX As Long
        Dim Fieldname As String
        Dim Fieldnames As String
        Dim FieldValue As String
        Dim FieldValues As String
        Dim SplitRowArr As Object
        Dim FieldsArr() As String
        Dim Criteria As String
        Dim TotalFields As Long
        Dim ErrMessage As String = ""
        Dim TotalRows As Long
        Dim RowNumber As Long

        'Set Dic_Values = Scripting.Dictionary

        Fieldnames = GetMyFields(DBTable, connString, ErrMessage)
        If Len(ErrMessage) > 0 Then
            frmMainGIForm.txtMessages.Text = ErrMessage
        End If
        'STOPPED at FieldIDX = 5 with an Array out of bounds message because SplitRowArr() only has 0 to 4 values for some reason.
        'FieldsArr = strToStringArray(Fieldnames, ",", 1, False, False, False, "_", False)
        FieldsArr = strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
        TotalFields = UBound(FieldsArr)
        FieldIDX = 1
        RowNumber = RowIDX
        'TotalRows = Dic_Values(CStr(TAbleIDX) & "_" & "TOTAL_ROWS" & "_" & CStr(RowNumber))

        Do While FieldIDX <= TotalFields
            If Dic_Values.Exists(CStr(TAbleIDX) & DelimChar & strDeliveryRef & DelimChar & FieldsArr(FieldIDX) & DelimChar & CStr(RowNumber)) Then
                FieldValue = Dic_Values(CStr(TAbleIDX) & DelimChar & strDeliveryRef & DelimChar & FieldsArr(FieldIDX) & DelimChar & CStr(RowNumber))
                'SplitRowArr = Split(Dic_Values(FieldsArr(FieldIDX)), "_")
                If Len(FinalFields) = 0 Then
                    FinalFields = FieldsArr(FieldIDX)
                Else
                    FinalFields = FinalFields & "," & FieldsArr(FieldIDX)
                End If
                If Len(FinalValues) = 0 Then
                    FinalValues = FieldValue
                Else
                    FinalValues = FinalValues & "," & FieldValue
                End If
            End If
            FieldIDX = FieldIDX + 1
        Loop

    End Sub

    Function GetTotalFrameRows(TheFrame As ScrollableControl, ByVal LowestTAG As Long, ByRef HighestTag As Long, ByVal TotalControls As Long, VTTAG As Long,
    Optional Subtract_VTTAG As Boolean = False) As Long
        Dim FrameCtrl As Control
        Dim IDX As Long
        Dim strTAGID As String
        Dim TagID As Long
        Dim TotalRows As Long

        TotalRows = 0
        HighestTag = 0
        If TheFrame Is Nothing Then
            GetTotalFrameRows = 0
            Exit Function
        End If
        For Each FrameCtrl In TheFrame.Controls
            strTAGID = FrameCtrl.Tag
            If IsNumeric(strTAGID) Then
                TagID = CLng(strTAGID)
                If TagID > VTTAG Then
                    If Subtract_VTTAG Then
                        TagID = TagID - VTTAG
                    End If
                End If
                If TagID > HighestTag Then
                    HighestTag = TagID
                End If
            End If
        Next
        If HighestTag > 0 And TotalControls > 0 Then
            TotalRows = (HighestTag - (LowestTAG - 1)) / TotalControls
        End If
        GetTotalFrameRows = TotalRows

    End Function

    Sub ADD_ROW_TO_DATA_GRID_VIEW_OPERATIVES(dgv As Object, RowNum As Long, StartTAGNumber As Long, StartButtonTAG As Long)
        Dim IDX As Int32
        Dim btnStartTime As New DataGridViewButtonColumn()
        Dim btnEndTime As New DataGridViewButtonColumn()
        'Dim cmbName As New DataGridViewComboBoxColumn()
        'Dim cmbActivity As New DataGridViewComboBoxColumn()
        Dim txtName As New DataGridViewTextBoxColumn()
        Dim txtActivity As New DataGridViewTextBoxColumn()
        Dim txtRow As New DataGridViewTextBoxColumn()
        Dim txtStartTime As New DataGridViewTextBoxColumn()
        Dim txtEndTime As New DataGridViewTextBoxColumn()
        Dim txtComments As New DataGridViewTextBoxColumn()

        'dgv.ColumnCount = 8

        txtRow.HeaderText = "#"
        txtRow.Name = "OpRow"
        txtRow.Tag = StartTAGNumber

        txtRow.Visible = True

        'cmbName.HeaderText = "Select Employee"
        'cmbName.Name = "cmbEmployee"
        'cmbName.MaxDropDownItems = 100
        'cmbName.Tag = StartTAGNumber + 1
        'cmbName.Items.Add("Richard Stilgo")
        'cmbName.Visible = True

        txtName.HeaderText = "Select Employee"
        txtName.Name = "txtEmployee"
        txtName.Tag = StartTAGNumber + 1
        txtName.Visible = True

        txtActivity.HeaderText = "Select Activity"
        txtActivity.Name = "txtActivity"
        txtActivity.Tag = StartTAGNumber + 2
        txtActivity.Visible = True



        btnStartTime.HeaderText = "Start Time"
        btnStartTime.Name = "btnStartTime"
        btnStartTime.Text = "@S"
        btnStartTime.Tag = "btn" & CStr(StartButtonTAG)
        btnStartTime.Visible = True

        txtStartTime.HeaderText = "Start Time"
        txtStartTime.Name = "txtStartTime"
        txtStartTime.Tag = StartButtonTAG + 3 + 400
        txtStartTime.Visible = True

        btnEndTime.HeaderText = "End Time"
        btnEndTime.Name = "btnEndTime"
        btnEndTime.Text = "@F"
        btnEndTime.Tag = "btn" & CStr(StartButtonTAG + 1)
        btnEndTime.Visible = True

        txtEndTime.HeaderText = "End Time"
        txtEndTime.Name = "txtEndTime"
        txtEndTime.Tag = StartButtonTAG + 4 + 400
        txtEndTime.Visible = True

        txtComments.HeaderText = "Total Time"
        txtComments.Name = "txtTotalTime"
        txtComments.Tag = StartTAGNumber + 5
        txtComments.Visible = True


        'This ADDS physical COLUMNS to an existing GRID.

        dgv.columns.add(txtRow)
        dgv.columns.add(txtName)
        dgv.columns.add(txtActivity)
        dgv.columns.add(btnStartTime)
        dgv.columns.add(txtStartTime)
        dgv.columns.add(btnEndTime)
        dgv.columns.add(txtEndTime)
        dgv.columns.add(txtComments)
        dgv.visible = True
        'This ADDS a DATA COLUMN to a dynamic grid based on a data table:
        'DIM IDXCOl As DataColumn
        'IDXCol = ds.Tables("srcTable").Columns.Add
        'IDXCol.SetOrdinal(ButtonColumnPos)
        'ds.Tables("srcTable").Rows(RowIDX)(ButtonColumnPos) = value


        btnStartTime.UseColumnTextForButtonValue = True
        btnEndTime.UseColumnTextForButtonValue = True


    End Sub

    Function Change_STATUS(DeliveryRef As String, NewStatus As String,
                           Optional ErrMessages As String = "") As Boolean
        'Change the STATUS field in tblDeliveryInfo:
        Dim SaveOK As Boolean
        Dim Fieldnames As String
        Dim FieldValues As String
        Dim DBTAble As String = "tblDeliveryInfo"
        Dim UpdateCriteria As String = ""
        Dim ExcludeFields As String = ""
        Dim CurrentDAte As DateTime
        Dim AutoComment As String
        Dim OldStatus As String
        Dim LastSaved As String
        Dim EnteredBy As String
        Dim OldComment As String
        Dim UpdatedBy As String
        Dim UpdatedByUsername As String
        Dim UpdatedByEmpNo As String
        Dim UpdatedAt As String
        Dim SearchText As String
        Dim SearchField As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim LoadedOK As Boolean = False
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim SearchCriteria As String
        Dim ErrMessage As String

        Change_STATUS = False
        If Len(DeliveryRef) = 0 Then
            ErrMessages = "NO DELIVERY REFERENCE SUPPLIED"
            Exit Function
        End If

        SearchText = DeliveryRef
        SearchField = "DeliveryReference"
        FieldType = "STRING"
        ReturnField = "ID"
        ReturnValue = ""
        SortFields = "DeliveryDate"
        Reversed = True 'GET LATEST DELIVERY REFERENCE FIRST. BUT WHAT DOES Find_myQuery ACTUALLY do ????? may pick up last found record ?????
        'frmMainGIForm.InsertValueIntoForm(ControlPanelFormName & FormID, "comASNNo", ControlASN)
        SearchCriteria = SearchField & " = " & Chr(34) & SearchText & Chr(34)
        LoadedOK = Find_myQuery(frmMainGIForm.myConnString, DBTAble, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If LoadedOK Then
            OldStatus = GetMYValuebyFieldname(AllValues, AllFields, "STATUS")
            OldComment = GetMYValuebyFieldname(AllValues, AllFields, "Delivery_Comments")
            LastSaved = GetMYValuebyFieldname(AllValues, AllFields, "Last_Saved")
            EnteredBy = GetMYValuebyFieldname(AllValues, AllFields, "UpdatedByName")

            UpdatedAt = Now().ToString("dd/MM/yyyy HH:mm:ss")
            UpdatedBy = frmMainGIForm.myFirstname & " " & frmMainGIForm.myLastname
            AutoComment = "Previous Save: " & LastSaved & vbCrLf
            AutoComment = AutoComment & " Previous STATUS: " & OldStatus & vbCrLf
            AutoComment = AutoComment & " Entered By: " & EnteredBy & vbCrLf
            AutoComment = AutoComment & " NEW STATUS: " & NewStatus & vbCrLf
            AutoComment = AutoComment & " Updated By: " & UpdatedBy & " AT " & UpdatedAt & vbCrLf
            AutoComment = AutoComment & vbCrLf & OldComment

            CurrentDAte = Now()
            UpdateCriteria = "ID = " & ReturnValue
            Fieldnames = "STATUS,UpdatedByName,UpdatedByEmpno,UpdatedByUsername,Last_Saved,Delivery_Comments"
            UpdatedByEmpNo = frmMainGIForm.myEmpNo
            UpdatedByUsername = frmMainGIForm.myUsername
            FieldValues = NewStatus
            FieldValues = FieldValues & "," & UpdatedBy
            FieldValues = FieldValues & "," & UpdatedByEmpNo
            FieldValues = FieldValues & "," & UpdatedByUsername
            FieldValues = FieldValues & "," & CurrentDAte
            FieldValues = FieldValues & "," & AutoComment
            SaveOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "", DBTAble, Fieldnames, FieldValues,
                UpdateCriteria, ExcludeFields, ErrMessages, False, ",")
        Else
            MsgBox("Could not find corresponding Delivery Reference")
            Exit Function
        End If

        If SaveOK Then
            Change_STATUS = True
        End If

    End Function

    Function Change_LockSTATUS(DeliveryRef As String, NewStatus As String, Optional ErrMessages As String = "") As Boolean
        'Change the STATUS field in tblDeliveryInfo:
        Dim SaveOK As Boolean
        Dim Fieldnames As String
        Dim FieldValues As String
        Dim DBTAble As String = "tblDeliveryInfo"
        Dim UpdateCriteria As String = ""
        Dim ExcludeFields As String = ""

        Change_LockSTATUS = False
        If Len(DeliveryRef) = 0 Then

        End If
        UpdateCriteria = "DeliveryReference = " & Chr(34) & DeliveryRef & Chr(34)
        Fieldnames = "LOCK_STATUS"
        FieldValues = NewStatus
        SaveOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "", DBTAble, Fieldnames, FieldValues,
            UpdateCriteria, ExcludeFields, ErrMessages, False, ",")

        If SaveOK Then
            Change_LockSTATUS = True
        End If

    End Function

    Function Get_LOCK_STATUS(ReferenceNumber As String) As String
        Dim FoundRef As Boolean
        Dim SearchField As String
        Dim SearchText As String
        Dim DBTable As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim ALLValues() As Object
        Dim AllFields() As String
        Dim SearchCriteria As String
        Dim SortFields As String
        Dim Reversed As Boolean
        Dim ErrMessage As String
        Dim RefNumber As String
        Dim ASNNumber As String
        Dim LockStatus As String

        DBTable = "tblDeliveryInfo"

        SearchField = "DeliveryReference"
        SearchText = ReferenceNumber
        ReturnField = "ID"
        ReturnValue = ""
        SortFields = "DeliveryDate"
        Reversed = True
        FieldType = "STRING"
        Get_LOCK_STATUS = ""
        FoundRef = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, ALLValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If FoundRef Then
            LockStatus = GetMYValuebyFieldname(ALLValues, AllFields, "LOCK_STATUS")
            RefNumber = GetMYValuebyFieldname(ALLValues, AllFields, "DeliveryReference")
        End If
        Get_LOCK_STATUS = LockStatus

    End Function

    Function Get_STATUS(ReferenceNumber As String) As String
        Dim FoundRef As Boolean
        Dim SearchField As String
        Dim SearchText As String
        Dim DBTable As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim ALLValues() As Object
        Dim AllFields() As String
        Dim SearchCriteria As String
        Dim SortFields As String
        Dim Reversed As Boolean
        Dim ErrMessage As String
        Dim RefNumber As String
        Dim ASNNumber As String
        Dim Status As String

        DBTable = "tblDeliveryInfo"

        SearchField = "DeliveryReference"
        SearchText = ReferenceNumber
        ReturnField = "ID"
        ReturnValue = ""
        SortFields = "DeliveryDate"
        Reversed = True
        FieldType = "STRING"
        Get_STATUS = ""
        FoundRef = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, ALLValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If FoundRef Then
            Status = GetMYValuebyFieldname(ALLValues, AllFields, "STATUS")
            RefNumber = GetMYValuebyFieldname(ALLValues, AllFields, "DeliveryReference")
        End If
        Get_STATUS = Status

    End Function

    Sub Check_For_Completion(strDeliveryDate As String, strDeliveryRef As String, Optional ChangeStatus As Boolean = False)
        Dim Completed As Boolean = False
        Dim FLMMessages As String
        Dim OpMessages As String
        Dim Answer As Integer
        Dim GridMessage As String
        Dim CompleteMessage As String
        Dim ASNno As String = ""
        Dim FormUnique As String
        Dim MDIChildIndex As Integer
        Dim ChildForm As String

        '1) Make sure ALL OP FINISH TIMES are in place.
        '2) Make sure FLM FINISH TIME is in place.
        '3) Issue an are you sure ? dialog to confirm.
        '4) Write to the tblDeliveryInfo - STATUS field - "COMPLETE". - use ChangeStatus()

        'Me.txtRefCompleteComments.Text = ""
        FLMMessages = ""
        OpMessages = ""
        GridMessage = ""
        CompleteMessage = ""
        Completed = Test_For_Completion(strDeliveryDate, strDeliveryRef, FLMMessages, OpMessages, GridMessage)
        If Len(OpMessages) > 0 Then
            If Len(CompleteMessage) > 0 Then
                CompleteMessage = CompleteMessage & vbCrLf & OpMessages
                'Me.txtRefCompleteComments.AppendText(OpMessages)
            Else
                CompleteMessage = OpMessages
            End If
        End If
        If Len(FLMMessages) > 0 Then
            If Len(CompleteMessage) > 0 Then
                'Me.txtRefCompleteComments.AppendText(vbCrLf & FLMMessages)
                CompleteMessage = CompleteMessage & vbCrLf & FLMMessages
            Else
                CompleteMessage = FLMMessages
            End If
        End If
        If Len(GridMessage) > 0 Then
            If Len(CompleteMessage) > 0 Then
                'Me.txtRefCompleteComments.AppendText(vbCrLf & FLMMessages)
                CompleteMessage = CompleteMessage & vbCrLf & GridMessage
            Else
                CompleteMessage = GridMessage
            End If
        End If
        If Completed Then
            'GREAT - change the STATUS to COMPLETED ???
            'SO the STATUS CAN ONLY be changed to COMPLETED IF THE BUTTON IS PRESSED.
            'HERE WE ARE CHECKING the Potential for the reference to be completed.
            'ARE ALL THE CRITERIA / ATTRIBUTES for COMPLETED met ?????
            'NO just stays at NO on the Control Panel.
            'SET Me.txtReferenceCompleted.Text = "YES" only if the button is pressed and the ATTRIBUTES are met.

            frmMainGIForm.InsertValueIntoForm(CPFormName & FormID, "txtReferenceCompleted", "YES")
            If frmMainGIForm.FindMyChild(strDeliveryRef, ASNno, False, FormUnique, MDIChildIndex, ChildForm, "btnReferenceCompleted", "GREEN") = True Then
                'OK GREEN
            End If
            If ChangeStatus Then
                Call Change_STATUS(strDeliveryRef, "COMPLETED")
                Call Change_LockSTATUS(strDeliveryRef, "CLEAR")

            End If
            'Me.btnReferenceCompleted.BackColor = Color.Green
        Else
            'Me.txtReferenceCompleted.Text = "NO"
            frmMainGIForm.InsertValueIntoForm(CPFormName & FormID, "txtReferenceCompleted", "NO")
            If frmMainGIForm.FindMyChild(strDeliveryRef, ASNno, False, FormUnique, MDIChildIndex, ChildForm, "btnReferenceCompleted", "RED") = True Then
                'OK RED
            End If
            If ChangeStatus Then
                Call Change_STATUS(strDeliveryRef, "IN PROGRESS")
                Call Change_LockSTATUS(strDeliveryRef, "LOCKED")
            End If
            'Me.btnReferenceCompleted.BackColor = Color.Red
            'Call Change_LockSTATUS(strDeliveryRef, "LOCKED")
        End If
        frmMainGIForm.InsertValueIntoForm(CPFormName & FormID, "txtRefCompleteComments", CompleteMessage)
    End Sub

    Function Test_For_Completion(strDeliveryDate As String, strDeliveryRef As String, ByRef FLMMessage As String, ByRef OpMessage As String,
                                 Optional ByRef GridMessage As String = "") As Boolean
        'TEST tblOperatives for completion of hours:
        'TEST Op_StartTime > 1970-01-01
        'TEST Op_FinishTime > 1970-01-01
        'TEST for current DeliveryReference
        'TEST for current FrameRowNumber in loop.
        'Do we need to test for the current date ???
        'The Original Delivery Date for the reference can be from a few days ago.
        Dim RowNum As Long
        Dim dtOpStartTime As DateTime
        Dim dtOpFinishTime As DateTime
        Dim dtDeliveryDate As DateTime
        Dim OpTableName As String
        Dim FLMTableName As String
        Dim InfoTableName As String
        Dim SearchField As String
        Dim SearchText As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim FoundFLMRef As Boolean
        Dim FoundInfo As Boolean
        Dim FieldType As String
        Dim SearchCriteria As String
        Dim AllFieldValues() As Object
        Dim AllFieldnames() As String
        Dim SortField As String = "DeliveryDate"
        Dim Reversed As Boolean = True
        Dim ErrMessages As String
        Dim strDateSearch As String
        Dim TotalOps As Long
        Dim OpRowOK() As Boolean
        Dim FLMDatesOK As Boolean = True
        Dim FLM_IsStartDate As Boolean = False
        Dim FLM_IsFinDate As Boolean = False
        Dim FLM_IsName As Boolean = False
        Dim FLMName As String
        Dim FLMStartDate As String
        Dim FLMFinDate As String
        Dim OpName As String
        Dim OpActivity As String
        Dim NoOpError As Boolean = True
        Dim GridNo As String
        Dim GridError As Boolean = False

        'Need to multisearch:
        'Need total number of Ops entered within the REference Passed:
        Test_For_Completion = False

        If Len(strDeliveryRef) = 0 Then
            Test_For_Completion = False
            Exit Function
        End If

        OpTableName = "tblOperatives"
        FLMTableName = "tblLabourHours"
        InfoTableName = "tblDeliveryInfo"
        SearchField = "DeliveryReference"
        SearchText = strDeliveryRef
        ReturnField = "ID"
        FieldType = "STRING"
        ReturnValue = ""
        strDateSearch = CDate("1970-01-01").ToString("yyyy-MM-dd HH:mm:ss")

        'SearchCriteria = "DATE(OP_StartTime) > " & CDate("0001-01-01 00:00:00")
        'SearchCriteria = SearchCriteria & " AND DATE(OP_Finishtime) > " & CDate("0001-01-01 00:00:00")
        TotalOps = Get_TotalRows(OpTableName, strDeliveryRef)
        ReDim OpRowOK(TotalOps)
        RowNum = 1
        Do While RowNum <= TotalOps
            SearchCriteria = "(DATE(" & "OP_StartTime" & ") " & "> " & Chr(39) & strDateSearch & Chr(39) & ")"
            SearchCriteria = SearchCriteria & " AND " & "(DATE(" & "OP_Finishtime" & ") " & "> " & Chr(39) & strDateSearch & Chr(39) & ")"
            SearchCriteria = SearchCriteria & " AND " & " FrameRowNumber = " & RowNum
            OpRowOK(RowNum) = Find_myQuery(frmMainGIForm.myConnString, OpTableName, SearchField, SearchText, "STRING",
                                                ReturnField, ReturnValue, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)

            If OpRowOK(RowNum) = False Then
                OpMessage = OpMessage & vbCrLf & "No Time #" & CStr(RowNum)
                NoOpError = False
            Else
                OpName = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "Op_Name")
                If Len(OpName) = 0 Or UCase(OpName) = UCase("Select Employee") Then
                    OpMessage = OpMessage & vbCrLf & "No Name #" & CStr(RowNum)
                    NoOpError = False
                End If
                OpActivity = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "Activity_Name")
                If Len(OpActivity) = 0 Or UCase(OpActivity) = UCase("Select Activity") Then
                    OpMessage = OpMessage & vbCrLf & "No Activity #" & CStr(RowNum)
                    NoOpError = False
                End If
            End If
            RowNum = RowNum + 1
        Loop
        If NoOpError Then
            OpMessage = "Ops Complete."
        End If
        SearchField = "DeliveryReference"
        SearchText = strDeliveryRef
        ReturnField = "ID"
        FieldType = "STRING"
        ReturnValue = ""
        SortField = "DeliveryDate"
        Reversed = True
        strDateSearch = CDate("1970-01-01").ToString("yyyy-MM-dd HH:mm:ss")

        SearchCriteria = ""
        'SearchCriteria = "(DATE(" & "FLM_StartTime" & ") " & "> " & Chr(39) & strDateSearch & Chr(39) & ")"
        'SearchCriteria = SearchCriteria & " AND " & "(DATE(" & "FLM_Finishtime" & ") " & "> " & Chr(39) & strDateSearch & Chr(39) & ")"
        'SearchCriteria = SearchCriteria & " AND " & " FLM_Name = " & RowNum
        FLMMessage = ""
        FoundFLMRef = Find_myQuery(frmMainGIForm.myConnString, FLMTableName, SearchField, SearchText, "STRING",
                             ReturnField, ReturnValue, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)
        If FoundFLMRef Then
            FLMName = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "FLM_Name")
            FLMStartDate = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "FLM_StartTime")
            FLMFinDate = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "FLM_Finishtime")
        End If
        If Len(FLMName) > 0 Then
            FLM_IsName = True
        Else
            FLMMessage = FLMMessage & vbCrLf & "No FLM Name"
            FLMDatesOK = False
        End If
        If IsDate(FLMStartDate) Then
            If CDate(FLMStartDate) > CDate("1970-01-01") Then
                FLM_IsStartDate = True
            Else
                FLMMessage = FLMMessage & vbCrLf & "No FLM Start Date"
                FLMDatesOK = False
            End If
        End If
        If IsDate(FLMFinDate) Then
            If CDate(FLMFinDate) > CDate("1970-01-01") Then
                FLM_IsFinDate = True
            Else
                FLMMessage = FLMMessage & vbCrLf & "No FLM Finish Date"
                FLMDatesOK = False
            End If
        End If

        SearchField = "DeliveryReference"
        SearchText = strDeliveryRef
        ReturnField = "ID"
        FieldType = "STRING"
        ReturnValue = ""
        strDateSearch = CDate("1970-01-01").ToString("yyyy-MM-dd HH:mm:ss")

        SearchCriteria = ""
        FoundInfo = Find_myQuery(frmMainGIForm.myConnString, InfoTableName, SearchField, SearchText, "STRING",
                             ReturnField, ReturnValue, AllFieldValues, AllFieldnames, SearchCriteria, SortField, Reversed, ErrMessages)
        If FoundInfo Then
            GridNo = GetMYValuebyFieldname(AllFieldValues, AllFieldnames, "Grid_No")
            If Len(GridNo) = 0 Then
                GridMessage = "Grid Number NOT entered."
                GridError = True
            Else
                GridMessage = "Grid OK (" & GridNo & ")"
                GridError = False
            End If
        Else
            GridMessage = "Problem Getting Grid Info"
            GridError = True
        End If

        If FLMDatesOK And NoOpError And GridError = False Then
            Test_For_Completion = True
            FLMMessage = "FLM COMPLETE."
        End If



    End Function

    Function ExtractHeadings(ByVal dgv As DataGridView, ByRef Headings As Object) As String
        Dim ColIDx As Integer
        Dim ColumnName As String
        Dim AllColHeadings As String
        Dim TotalCols As Integer

        ExtractHeadings = ""
        AllColHeadings = ""
        TotalCols = dgv.ColumnCount
        ReDim Headings(TotalCols + 1)
        For ColIDx = 0 To TotalCols
            ColumnName = dgv.Columns(ColIDx).HeaderText
            If Len(AllColHeadings) = 0 Then
                AllColHeadings = ColumnName
            Else
                AllColHeadings = AllColHeadings & "," & ColumnName
            End If
            Headings(ColIDx) = ""
            Headings(ColIDx) = ColumnName
        Next
        ExtractHeadings = AllColHeadings

    End Function

    Function SearchGrid(ByVal dgv As DataGridView, DBTable As String, ByVal ReturnSearchField As String, ByRef ReturnFieldValue As String,
                        ByRef RowNum As Integer, ByRef ColNum As Integer) As Boolean
        Dim SearchCriteria As String
        Dim strDateSearchFrom As String
        Dim strDateSearchTo As String
        Dim SearchField As String
        Dim SearchText As String
        Dim ReturnField As String = "ID"
        Dim ReturnValue As String = ""
        Dim AllFieldValues() As Object
        Dim AllFieldNames() As String
        Dim SortField As String = ""
        Dim Reversed As Boolean
        Dim ErrMessage As String
        Dim LocatedOK As Boolean

        SearchCriteria = "(DATE(" & "OP_StartTime" & ") " & ">= " & Chr(39) & strDateSearchFrom & Chr(39) & ")"
        SearchCriteria = SearchCriteria & " AND " & "(DATE(" & "OP_Finishtime" & ") " & "<= " & Chr(39) & strDateSearchTo & Chr(39) & ")"
        SearchCriteria = SearchCriteria & " AND " & " FrameRowNumber = " & RowNum
        LocatedOK = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING",
                                            ReturnField, ReturnValue, AllFieldValues, AllFieldNames, SearchCriteria, SortField, Reversed, ErrMessage)
        If LocatedOK Then
            ReturnFieldValue = GetMYValuebyFieldname(AllFieldValues, AllFieldNames, ReturnSearchField)
        Else
            ReturnFieldValue = ""
        End If

        SearchGrid = LocatedOK

    End Function

    Sub Populate_Timesheet()
        'READ DATA from tblTimesheetTemplate into Dic_Data() dictionary scripting variable:
        Dim Dic_Data As Object
        Dim DataArray(,) As String
        Dim FaultTemplatePath As String
        Dim xlsFaultForm As New Object
        Dim ErrMessage As String
        Dim varKEY As String
        Dim LastRow As Integer
        Dim LastCol As Integer
        Dim RowIDX As Integer
        Dim FoundPos As String
        Dim FoundArr As String()
        Dim FoundRow As Integer
        Dim FoundCol As Integer
        Dim ColIDX As Integer '32 bit signed integer = should be 2 billion ????
        Dim HashItem As String 'Field Number in template spreadsheet - from #1 to #8
        Dim TemplateFilename As String
        Dim NewFilename As String

        TemplateFilename = "EXCEL TIMESHEET TEMPLATE.XLSX"
        NewFilename = "Timesheet_" & Now().ToString("dd_MM_yyyy_HH_mm_ss")

        'ReDim DataArray(25, 5)
        'LastRow = UBound(DataArray, 1) + 1
        'LastCol = UBound(DataArray, 2) + 1
        FaultTemplatePath = Application.StartupPath
        TemplateFilename = FaultTemplatePath & "\" & TemplateFilename
        NewFileName = FaultTemplatePath & "\" & NewFilename
        Dic_Data("#DeliveryDate") = Now().ToString("dd/MM/yyyy")
        Dic_Data("#DeliveryRef") = "54321"
        Dic_Data("#Name1") = "Dan Goss"
        'DanG_EXCEL_Module.OpenWorkbookAndSetGlobal(TemplateFilename, FaultTemplatePath, True, ErrMessage, True)
        DanG_EXCEL_Module.Insert_Data_Into_Sheet(TemplateFilename, NewFilename, Dic_Data, 51, 0, 0, 1, 1, 54, 16)

        Dic_Data = CreateObject("Scripting.Dictionary")
        Dic_Data.comparemode = vbTextCompare


    End Sub

    Function Get_STATUS_BACKCOLOR(STATUS As String, Optional ByRef ReturnForeColor As System.Drawing.Color = Nothing) As System.Drawing.Color

        Get_STATUS_BACKCOLOR = Color.AliceBlue
        If UCase(STATUS) = "ALL" Then
            Get_STATUS_BACKCOLOR = Color.AliceBlue
            ReturnForeColor = Color.Black
        ElseIf UCase(STATUS) = "ROLLED" Then
            Get_STATUS_BACKCOLOR = Color.Crimson
            ReturnForeColor = Color.AliceBlue
        ElseIf UCase(STATUS) = "ROLLED OVER" Then
            Get_STATUS_BACKCOLOR = Color.Maroon
            ReturnForeColor = Color.White
        ElseIf UCase(STATUS) = "NO SHOW" Then
            Get_STATUS_BACKCOLOR = Color.Firebrick
            ReturnForeColor = Color.White
        ElseIf UCase(STATUS) = "CANCELLED" Then
            Get_STATUS_BACKCOLOR = Color.Red
            ReturnForeColor = Color.Black
        ElseIf UCase(STATUS) = "IN PROGRESS" Then
            Get_STATUS_BACKCOLOR = Color.Orange
            ReturnForeColor = Color.Black
        ElseIf UCase(STATUS) = "COMPLETED" Then
            Get_STATUS_BACKCOLOR = Color.LimeGreen
            ReturnForeColor = Color.Black
        Else
            Get_STATUS_BACKCOLOR = Color.AliceBlue
            ReturnForeColor = Color.Black
        End If


    End Function

    Sub Get_GRID_DETAILS(dgv As DataGridView, RowIndex As Integer, ColumnIndex As Integer,
                         Optional strUniqueCol As String = "",
                         Optional intUniqueCol As Integer = 0,
                         Optional FieldType As String = "INTEGER",
                         Optional DBTable As String = "tblOperatives")
        Dim GridValue As String
        Dim ColName As String
        Dim RefCol As Integer
        Dim AllValues As Object()
        Dim AllFields As String()
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ErrMessage As String
        Dim SearchCriteria As String
        Dim SearchField As String
        Dim SearchText As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim FoundKeyValue As Boolean
        Dim strDeliveryRef As String
        Dim Op_Name As String
        Dim Activity_Name As String
        Dim Op_StartTime As String
        Dim Op_FinishTime As String
        Dim TotalTime As String
        Dim dblTotalHours As Double

        'Need to test if the form is already open - with THAT reference ???
        If RowIndex < 0 Or ColumnIndex < 0 Then
            Exit Sub
        Else
            'Find out which columnIndex holds the reference number in the GRID ?
            RefCol = -1
            For i = 1 To dgv.Columns.Count - 1
                ColName = dgv.Columns(i).HeaderText
                If Len(strUniqueCol) > 0 Then
                    If InStr(UCase(ColName), strUniqueCol) > 0 Then
                        RefCol = i
                    End If
                End If

            Next
            If intUniqueCol > 0 Then
                RefCol = intUniqueCol
            End If


            GridValue = dgv.Rows(RowIndex).Cells(RefCol).Value

            'StrDeliveryRef = CType(RefValue, String)
            'If ColumnIndex = RefCol Then
            'StrDeliveryRef = CType(value, String)
            'frmMainGIForm.SelectedDeliveryRef = StrDeliveryRef

            'End If
            'If ColumnIndex = ASNCol Then
            'ASNNo = CType(value, String)
            'frmMainGIForm.SelectedASN = ASNNo
            'End If
            'MsgBox("value= " & value)

            'Just open up a NEW CHILD FORM for the ControlPanel:
            'txtSelectDeliveryDate.Text
            'If RefCol = ColumnIndex Or ASNCol = ColumnIndex Then
            'If Form cannot be found then create new child form instance:
            'If frmMainGIForm.FindMyChild(StrDeliveryRef, ASNNo, False, FormUnique, MDIChildIndex, ChildForm) = False Then
            'FormID = frmMainGIForm.Get_GUID()
            'frmMainGIForm.ControlPanelIdx = FormID
            'frmMainGIForm.SelectedDeliveryRef = StrDeliveryRef
            'cf.MdiParent = frmMainGIForm
            'cf.StartPosition = FormStartPosition.Manual
            AllValues = Nothing
            AllFields = Nothing
            SortFields = "DeliveryDate"
            Reversed = True
            ErrMessage = ""
            SearchCriteria = ""
            'SearchField = "DELIVERYREFERENCE"
            SearchText = GridValue
            ReturnField = "ID"
            ReturnValue = ""
            'If Len(StrDeliveryRef) > 0 Then
            FoundKeyValue = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
            ReturnValue = GetMYValuebyFieldname(AllValues, AllFields, "ID")

            'cf.Text = "Goods In Control Panel_" & "REF:" & StrDeliveryRef
            'ElseIf Len(ASNNo) > 0 Then
            'Find THE DELIVERY REFERENCE !!!
            'cf.Text = "Goods In Control Panel_" & Now().ToString("dd/MM/yyyy HH:mm:ss")
            'End If

            'cf.Name = CPFormName & FormID
            'cf.Tag = FormID
            'cf.txtDeliveryRef.Text = StrDeliveryRef
            'cf.comDeliveryRef.Text = StrDeliveryRef
            'cf.comASNNo.Text = ASNNo
            'cf.txtASNNum.Text = ASNNo
            'cf.Show()
            'cf.ShowDialog() - gives error
            'Clear_Entry_Controls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, 1, 40)
            'Clear_Entry_Controls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, 801, 807)
            'frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET_" & frmMainGIForm.ControlPanelIdx, "comDeliveryRef", StrDeliveryRef)
            'Application.DoEvents()
            'Else
            'Application.OpenForms.Item(ChildForm).Activate()
            'Need to gather all the fields from the found record.
            'Populate the entry fields:

        End If

    End Sub

End Module
