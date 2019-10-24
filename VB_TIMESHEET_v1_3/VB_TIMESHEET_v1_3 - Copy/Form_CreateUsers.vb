Imports System.Globalization
Public Class frmCreateUsers
    '****************************************************************************************************************************************************************
    '*
    '*  Program Title:                                            Goods Inwards TIMESHEET v1.1
    '*
    '*  Language:                                                 VB.NET - Visual Studio 2015 - © 2017 Microsoft
    '*
    '*  AUTHOR:                                                   Daniel Goss
    '*
    '*  Copyright:                                                Copyright 2018
    '*
    '*  Form Name:                                                Form: frmCreateUsers
    '*
    '*  External Name:                                            Form_CreateUsers.vb
    '*
    '*  Module last amended:                                      02/10/2018 12:10
    '*      First conversion to GI TIMESHEET and using Save with parameters call.
    '*
    '*
    '****************************************************************************************************************************************************************

    Public DefaultFields As String = ""
    Public AllValues As Object = Nothing
    Public AllFields As String() = Nothing
    Public DBUserCreatedOK As Boolean = False
    Public PermissionsCreatedOK As Boolean = False
    Dim strEmpNo As String = ""
    Dim strFirstname As String = ""
    Dim strLastname As String = ""
    Dim strWarehouse As String = ""
    Dim strDepartment As String = ""
    Dim strJobTitle As String = ""
    Dim strDateStarted As String = ""
    Dim strDateCreated As String = ""
    Dim strUserBarcode As String = ""
    Dim strUserBarcodeString As String = ""
    Dim strDateOfBirth As String = ""
    Dim strHostname As String = ""
    Dim strComments As String = ""
    Dim strUsername As String = ""
    Dim strPassword As String = ""
    Dim strAccessRights As String = ""
    Dim DBTable As String = ""
    Dim DBName As String = ""
    Dim SearchField As String = ""
    Dim SearchText As String = ""
    Dim ReturnField As String = ""
    Dim Criteria As String = ""
    Dim Errors As String = ""
    Dim ReturnValue As String = ""
    Dim FieldValues As String = ""
    Dim strUserID As String = ""
    Dim MyMessage As String = ""
    Dim UserExists As Boolean = False
    Dim OperatorExists As Boolean = False
    Dim InsertOK As Boolean = False
    Dim UpdatedOK As Boolean = False
    Dim DeletedOK As Boolean = False
    Dim IsLoggedIN As String = "0"
    Dim strSessionID As String = ""

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub tabUserEntry_MouseClick(sender As Object, e As MouseEventArgs) Handles tabUserEntry.MouseClick
        ' MsgBox("Selected Index= " & CStr(tabUserEntry.SelectedIndex))
        'Me.txtTitle.Text = "Please Enter Account Details"
    End Sub

    Private Sub tabUserEntry_Selected(sender As Object, e As TabControlEventArgs) Handles tabUserEntry.Selected
        ' MsgBox("SELECTED")
    End Sub

    Public Function ValidateDateTimeForError(ByVal checkInputValue As String, FormatString As String) As Boolean
        Dim returnError As Boolean
        Dim dateVal As DateTime

        returnError = False
        If DateTime.TryParseExact(checkInputValue, FormatString,
            System.Globalization.CultureInfo.CurrentCulture,
            DateTimeStyles.None, dateVal) Then
            returnError = True
        End If
        Return returnError
    End Function

    Public Sub Execute_Search()
        'SEARCH FOR EMPLOYEE NUMBER as in : txtSearchEmpNo.Text
        Dim SpacePos As Integer = 0

        strEmpNo = ""
        strEmpNo = Me.txtSearchEmpNo.Text
        SearchField = "EmpNo"
        SearchText = strEmpNo
        ReturnField = "ID"
        ReturnValue = ""
        AllValues = Nothing
        AllFields = Nothing
        DBTable = "tblUsers"
        Criteria = ""
        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            UserExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "", False, MyMessage)
        Else
            UserExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue,
                                      AllValues, AllFields, Criteria, "", False, MyMessage)
        End If
        If Len(MyMessage) > 0 Then
            Me.txtMessage.Text = MyMessage
        End If
        If UserExists Then
            'Need to populate all the fields on the form now:
            strUserID = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "ID")
            strEmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "EmpNo")
            strFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Firstname")
            strLastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Lastname")
            strWarehouse = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Warehouse")
            strDepartment = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Department")
            strJobTitle = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "JobTitle")
            strDateStarted = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "DateStarted")
            strUserBarcode = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "UserBarcode")
            strUserBarcodeString = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "UserBarcodeString")
            strComments = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Comments")
            strUsername = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Username")
            strPassword = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Password")
            strAccessRights = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "AccessRights")

            strDateStarted = Replace(strDateStarted, "-", "/")
            If IsDate(strDateStarted) Then
                strDateStarted = CDate(strDateStarted).ToString("dd/MM/yyyy HH:mm:ss")
            End If
            Me.txtID.Text = strUserID
            Me.txtEmpNum.Text = strEmpNo
            Me.txtFirstname_view.Text = strFirstname
            Me.txtLastname_view.Text = strLastname
            Me.txtWarehouse.Text = strWarehouse
            Me.txtDepartment.Text = strDepartment
            Me.txtJobTitle.Text = strJobTitle
            Me.txtDateStarted.Text = strDateStarted
            Me.txtUserBarcode.Text = strUserBarcode
            Me.rtbComments.Text = strComments
            Me.txtUsername.Text = strUsername
            Me.txtPassword.Text = strPassword
            Me.txtAccessRights.Text = strAccessRights

            If Not strUserID = ReturnValue Then
                MsgBox("UserID and ReturnValue mis-match")
                Me.txtMessage.AppendText("UserID and ReturnValue mis-match: " & strUserID & " " & ReturnValue)
            End If

            strEmpNo = Me.txtSearchEmpNo.Text
            SearchField = "EmpNo"
            SearchText = strEmpNo
            ReturnField = "ID"
            ReturnValue = ""
            Criteria = ""
            DBTable = "tblOperators"
            If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
                OperatorExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "", False, MyMessage)
            Else
                OperatorExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING",
                                              ReturnField, ReturnValue, AllValues, AllFields, Criteria, "", False, MyMessage)
            End If
            If Len(MyMessage) > 0 Then
                Me.txtMessage.Text = MyMessage
            End If
            If OperatorExists Then
                'MsgBox("Employee Already an Operator")
                'Show Remove Button:
                Me.btnRemoveOperator.Visible = True
                Me.btnAddOperator.Visible = False
            Else
                Me.btnRemoveOperator.Visible = False
                Me.btnAddOperator.Visible = True
            End If
        Else
            MsgBox("USER: " & Me.txtEmpNum.Text & " Not Found")
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim TestDT As DateTime
        Dim strTestDT As String = ""
        Dim ExcludeFields As String = ""
        Dim RecID As String
        Dim DBUsersTAble As String = "tblUsers"
        Dim FieldNames As String
        Dim ErrMessage As String
        Dim strRepeatPassword As String
        Dim strExistingPassword As String

        'SAVE User Details:

        If Len(Me.txtEmpNo.Text) > 0 Then
            strEmpNo = Me.txtEmpNo.Text
        Else
            strEmpNo = Me.txtEmpNum.Text
        End If
        strFirstname = Me.txtFirstname_view.Text
        strLastname = Me.txtLastname_view.Text
        strWarehouse = Me.txtWarehouse.Text
        strDepartment = Me.txtDepartment.Text
        strJobTitle = Me.txtJobTitle.Text
        strRepeatPassword = Me.txtPasswordRepeat.Text
        If Len(Me.txtDateStarted.Text) > 0 Then
            strDateStarted = Me.txtDateStarted.Text
            If IsDate(strDateStarted) Then
                strDateStarted = CDate(strDateStarted).ToString("dd/MM/yyyy HH:mm:ss")
            Else
                strDateStarted = CDate("1970-01-01 00:00:00").ToString("dd/MM/yyyy HH:mm:ss")
            End If
        Else
            strDateStarted = CDate("1970-01-01 00:00:00").ToString("dd/MM/yyyy HH:mm:ss")
        End If
        strUserBarcode = Me.txtUserBarcode.Text
        strUserBarcodeString = Me.txtUserBarcode.Text

        strComments = Me.rtbComments.Text
        strUsername = Me.txtUsername.Text
        strPassword = Me.txtPassword.Text
        strAccessRights = Me.txtAccessRights.Text
        strUserID = Me.txtID.Text 'IF aVailable and populated ???



        DBTable = "tblUsers"
        DBName = "AssetRegister.accdb"


        SearchField = "UserBarcode"
        SearchText = strUserBarcode
        ReturnField = "ID"
        ReturnValue = ""

        FieldValues = strEmpNo & "," & strFirstname & "," & strLastname & "," & strWarehouse
        FieldValues = FieldValues & "," & strDepartment & "," & strJobTitle & "," & strDateStarted & "," & strUserBarcode
        FieldValues = FieldValues & "," & strUserBarcodeString & "," & strComments & "," & strUsername
        FieldValues = FieldValues & "," & strPassword
        FieldValues = FieldValues & "," & strAccessRights

        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            Criteria = ""
            UserExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "")
            If Not UserExists Then 'INSERT NEW RECORD:
                ExcludeFields = "ID"
                InsertOK = InsertUpdateRecord(False, DBName, "", DBTable, "", FieldValues, Criteria, ExcludeFields)
                If InsertOK Then
                    MsgBox("FINISHED - INSERT OK")
                Else
                    MsgBox("Finished - Error during SAVE.")
                End If
            Else 'UserExists then UPDATE RECORD WHERE ID = passed ID - will need to be as a text box.
                ExcludeFields = "ID,EmpNo" 'EmpNo is EXCLUDED due to the tblUsers being populated already.
                UpdatedOK = InsertUpdateRecord(True, DBName, "", DBTable, "", FieldValues, Criteria, ExcludeFields)
                If UpdatedOK Then
                    MsgBox("Finished Update - OK.")
                Else
                    MsgBox("Finished Update - But there was a problem")
                End If
            End If
        Else
            If frmMainGIForm.myLoggedIn Then

                UserExists = Find_myQuery(frmMainGIForm.myConnString, DBUsersTAble, SearchField, SearchText, "STRING", ReturnField, ReturnValue, AllValues, AllFields, Criteria, "", False, MyMessage)
                If Not UserExists Then 'INSERT NEW RECORD:
                    ExcludeFields = "ID"
                    strExistingPassword = Me.txtPasswordRepeat.Text
                    If Not strExistingPassword = strPassword Then
                        MsgBox("Passwords Do Not Match")
                        Exit Sub
                    End If
                    InsertOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, RecID, DBUsersTAble,
                                                                     FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage, False)
                    If InsertOK Then
                        MsgBox("OK NEW USER ADDED.")

                    Else
                        MsgBox("Error during SAVE.")
                    End If
                Else 'UserExists then UPDATE RECORD WHERE ID = passed ID - will need to be as a text box.
                    'FieldValues = Chr(34) & strEmpNo & Chr(34)
                    Criteria = "ID = " & ReturnValue
                    ExcludeFields = "ID,EmpNo" 'EmpNo is EXCLUDED due to the tblUsers being populated already.
                    UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, RecID, DBUsersTAble,
                                                                     FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage, False)
                    If UpdatedOK Then
                        MsgBox("Finished Update - OK.")

                    Else
                        MsgBox("Error During Update")
                        Me.txtMessage.AppendText(vbCrLf & MyMessage)
                    End If
                End If
            End If
        End If
        Criteria = ""
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear all the fields - so user can enter another name:

        Me.txtEmpNo.Text = ""
        Me.txtEmpNum.Text = ""
        Me.txtFirstname_view.Text = ""
        Me.txtLastname_view.Text = ""
        Me.txtWarehouse.Text = ""
        Me.txtDepartment.Text = ""
        Me.txtJobTitle.Text = ""
        Me.txtDateStarted.Text = "1970/01/01"
        Me.txtUserBarcode.Text = ""
        Me.rtbComments.Text = ""
        Me.txtUsername.Text = ""
        Me.txtPassword.Text = ""
        Me.txtPasswordRepeat.Text = ""
        Me.txtAccessRights.Text = ""
        Me.txtFirstname.Text = ""
        Me.txtLastname.Text = ""

        Me.btnAddOperator.Visible = False
        Me.btnRemoveOperator.Visible = False

        Me.txtEmpNum.Focus()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmCreateUsers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Autoload any default fields - such as when the form is being used to update a record:
        Dim FieldArray As String()
        Dim AccessTypes As String = ""
        Dim AccessTypesArray As String()
        Dim IDX As Integer = 0
        'Populate combo box with Database Types:

        ReDim AccessTypesArray(1)
        Me.txtID.Visible = False
        Me.txtOutput.Visible = False
        AccessTypes = "Normal,Admin"
        LoadMyConfigKey("DatabaseTypes", AccessTypes)
        If Len(AccessTypes) > 0 Then
            AccessTypesArray = strToStringArray(AccessTypes, ",")
        Else
            AccessTypesArray = {"Normal", "Admin"}
        End If
        For IDX = 0 To UBound(AccessTypesArray)
            comAccessRights.Items.Add(AccessTypesArray(IDX))

        Next

        If Len(DefaultFields) > 0 Then
            ReDim FieldArray(1)
            FieldArray = strToStringArray(DefaultFields, ",")
            Me.txtID.Visible = True
            Me.txtID.Text = FieldArray(0)
            Me.txtEmpNum.Text = FieldArray(1)
            Me.txtFirstname_view.Text = FieldArray(2)
            Me.txtLastname_view.Text = FieldArray(3)
            Me.txtWarehouse.Text = FieldArray(4)
            Me.txtDepartment.Text = FieldArray(5)
            Me.txtJobTitle.Text = FieldArray(6)
            Me.txtDateStarted.Text = FieldArray(7)
            Me.txtUserBarcode.Text = FieldArray(8)
            Me.rtbComments.Text = FieldArray(11)
            Me.txtUsername.Text = FieldArray(12)
            Me.txtPassword.Text = FieldArray(13)
            Me.txtAccessRights.Text = FieldArray(14)

        End If
        If UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
            Me.btnPopulateUsers.Visible = True
            Me.txtStartNum.Visible = True
            Me.txtTotalQty.Visible = True
        End If
        Me.btnSearchEmpNo.Focus()
        txtSearchEmpNo.Focus()
        KeyPreview = True

    End Sub

    Private Sub txtFirstname_TextChanged(sender As Object, e As EventArgs) Handles txtFirstname_view.TextChanged
        txtFirstname.Text = txtFirstname_view.Text
    End Sub

    Private Sub txtLastname_TextChanged(sender As Object, e As EventArgs) Handles txtLastname_view.TextChanged
        txtLastname.Text = txtLastname_view.Text
    End Sub

    Private Sub btnPopulateUsers_Click(sender As Object, e As EventArgs) Handles btnPopulateUsers.Click
        'This will populate the User Table - only will need to be called ONCE.
        Dim IDX As Integer = 0
        'Dim LimitNumber As Integer = 999999
        Dim LimitNumber As Integer = 100
        Dim NewCode As String = ""
        Dim CodeNumber As Integer = 1
        Dim StartNumber As Integer = 1
        Dim Percentage As Single = 0.0F
        Dim OldCode As String = ""
        Dim Fieldnames As String = ""
        Dim ErrMessage As String = ""
        Dim RecID As String = ""

        If UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
            If IsNumeric(Me.txtStartNum.Text) And Len(Me.txtStartNum.Text) > 0 Then
                CodeNumber = CInt(Me.txtStartNum.Text)
            Else
                CodeNumber = 1
            End If

            If IsNumeric(Me.txtTotalQty.Text) And Len(Me.txtTotalQty.Text) > 0 Then
                LimitNumber = CInt(Me.txtTotalQty.Text)
            Else
                LimitNumber = 100
            End If

            KeyPreview = True

            For IDX = 1 To LimitNumber
                'NEW - 7 DIGITS - INSERT users from 0000001 to 1000000 and AGY0001 to AGY9999 - well AGY2000 for now.
                If Me.rbAgency.Checked Then
                    If IDX + (CodeNumber - 1) < 10 Then
                        OldCode = "AGY" & "00" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100 Then
                        OldCode = "AGY" & "0" & CStr((CodeNumber - 1) + IDX)
                    Else
                        OldCode = "AGY" & CStr((CodeNumber - 1) + IDX)
                    End If

                    If IDX + (CodeNumber - 1) < 10 Then
                        NewCode = "AGY" & "000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100 Then
                        NewCode = "AGY" & "00" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 1000 Then
                        NewCode = "AGY" & "0" & CStr((CodeNumber - 1) + IDX)
                    Else
                        NewCode = "AGY" & CStr((CodeNumber - 1) + IDX)
                    End If
                Else
                    If IDX + (CodeNumber - 1) < 10 Then
                        OldCode = "00000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100 Then
                        OldCode = "0000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 1000 Then
                        OldCode = "000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 10000 Then
                        OldCode = "00" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100000 Then
                        OldCode = "0" & CStr((CodeNumber - 1) + IDX)
                    Else
                        OldCode = CStr((CodeNumber - 1) + IDX)
                    End If

                    If IDX + (CodeNumber - 1) < 10 Then
                        NewCode = "000000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100 Then
                        NewCode = "00000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 1000 Then
                        NewCode = "0000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 10000 Then
                        NewCode = "000" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 100000 Then
                        NewCode = "00" & CStr((CodeNumber - 1) + IDX)
                    ElseIf IDX + (CodeNumber - 1) < 1000000 Then
                        NewCode = "0" & CStr((CodeNumber - 1) + IDX)
                    Else
                        NewCode = CStr((CodeNumber - 1) + IDX)
                    End If
                End If

                strEmpNo = NewCode
                strFirstname = Me.txtFirstname_view.Text
                strLastname = Me.txtLastname_view.Text
                strWarehouse = Me.txtWarehouse.Text
                strDepartment = Me.txtDepartment.Text
                strJobTitle = Me.txtJobTitle.Text
                strUserBarcode = NewCode
                strUserBarcodeString = OldCode 'NEEDS ENCODING.
                strComments = Me.rtbComments.Text
                strUsername = Me.txtUsername.Text & NewCode
                strPassword = Me.txtPassword.Text
                strAccessRights = Me.txtAccessRights.Text
                'strDateStarted = "1970-01-01 01:00:00"
                strDateStarted = CStr(Now())
                'strUserID = Me.txtID.Text

                DBTable = "tblUsers"
                DBName = "AssetRegister.accdb"

                SearchField = "EmpNo"
                SearchText = OldCode
                ReturnField = "ID"
                ReturnValue = ""
                Fieldnames = ""
                FieldValues = strEmpNo & "," & strFirstname & "," & strLastname & "," & strWarehouse
                FieldValues = FieldValues & "," & strDepartment & "," & strJobTitle & "," & strDateStarted & "," & strUserBarcode
                FieldValues = FieldValues & "," & strUserBarcodeString & "," & strComments & "," & strUsername
                FieldValues = FieldValues & "," & strPassword
                FieldValues = FieldValues & "," & strAccessRights

                If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
                    Criteria = ""
                    UserExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "")

                    If Not UserExists Then 'INSERT NEW RECORD:

                        InsertOK = InsertUpdateRecord(False, DBName, "", DBTable, "", FieldValues, Criteria, "ID")
                        If InsertOK Then
                            txtMessage.AppendText("INSERT Created -  " & strEmpNo & " OK")
                        Else
                            txtMessage.AppendText("Error with EMP NO: " & strEmpNo)
                        End If
                    Else 'UserExists then UPDATE RECORD WHERE ID = passed ID - will need to be as a text box.
                        Criteria = ReturnField & " = " & ReturnValue
                        'UpdatedOK = InsertUpdateRecord(True, DBName, "", DBTable, "", FieldValues, Criteria, "ID")
                        'If UpdatedOK Then
                        'txtMessage.AppendText("UPDATE Performed -  " & strEmpNo & " OK")
                        'Else
                        'txtMessage.AppendText("Error UPDATE with EMP NO: " & strEmpNo)
                        'End If
                    End If
                Else
                    If Module_DanG_MySQL_Tools.mylogged_in Then
                        Criteria = ""
                        UserExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField,
                                                  ReturnValue, AllValues, AllFields, Criteria, "", False, MyMessage)
                        If Not UserExists Then 'INSERT NEW RECORD:

                            InsertOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, RecID, DBTable,
                                                                     FieldNames, FieldValues, Criteria, "ID", ErrMessage, False)
                            If InsertOK Then

                                'DBUserCreatedOK = CreateDBUser(frmMainGIForm.myConnString, strHostname, strUsername, strPassword, MyMessage)
                                'If DBUserCreatedOK = True Then
                                ' txtMessage.AppendText("Error with EMP NO: " & strEmpNo)
                                ''PermissionsCreatedOK = myPermissions("assetregister", frmMainGIForm.connString, strUsername, strPassword, strAccessRights, strHostname, MyMessage)
                                'If PermissionsCreatedOK = True Then
                                'Me.txtMessage.AppendText(vbCrLf & " The Permissions were created for user: " & strAccessRights)
                                'MsgBox("FINISHED - INSERT USER OK")
                                'End If
                                'txtMessage.AppendText("INSERT Created -  " & strEmpNo & " OK")
                            Else
                                txtMessage.AppendText("Error During Insert Name: " & strEmpNo)
                            End If
                        Else 'UserExists then UPDATE RECORD WHERE ID = passed ID - will need to be as a text box.
                            FieldValues = strEmpNo
                            FieldValues = FieldValues & "," & strUserBarcode
                            Fieldnames = "EmpNo,UserBarcode"
                            Criteria = "EmpNo" & " = " & Chr(34) & OldCode & Chr(34)
                            UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, RecID, DBTable,
                                                                     Fieldnames, FieldValues, Criteria, "ID", ErrMessage, False)
                            'If UpdatedOK Then
                            'txtMessage.AppendText("UPDATE Performed -  " & strEmpNo & " OK")

                            'Else
                            'txtMessage.AppendText("Error UPDATE with EMP NO: " & strEmpNo)
                            'End If
                        End If
                    End If
                End If

                Percentage = (IDX / LimitNumber) * 100
                pbBar_Accounts.Value = CInt(Percentage)
                Application.DoEvents()

            Next
            MsgBox("FINISHED POPULATING")
        Else
            MsgBox("Insufficient Access Level")
        End If
    End Sub

    Sub Convert_OldEmpNo_NewEmpNo(ByVal CodeType As String, ByVal TheCodeNumber As Long, ByRef OldCode As String, ByRef NewCode As String, Optional LimitNumberIndex As Integer = 1)
        Dim CodeNumber As Long = 0
        Dim StartNumber As Integer = 1
        Dim Percentage As Single = 0.0F

        OldCode = ""
        NewCode = ""
        CodeNumber = TheCodeNumber
        If UCase(CodeType) = "AGENCY" Then
            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                OldCode = "AGY" & "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                OldCode = "AGY" & "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                OldCode = "AGY" & CStr((CodeNumber - 1) + LimitNumberIndex)
            End If

            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                NewCode = "AGY" & "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                NewCode = "AGY" & "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                NewCode = "AGY" & "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                NewCode = "AGY" & CStr((CodeNumber - 1) + LimitNumberIndex)
            End If
        Else
            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                OldCode = "00000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                OldCode = "0000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                OldCode = "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 10000 Then
                OldCode = "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100000 Then
                OldCode = "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                OldCode = CStr((CodeNumber - 1) + LimitNumberIndex)
            End If

            If LimitNumberIndex + (CodeNumber - 1) < 10 Then
                NewCode = "000000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100 Then
                NewCode = "00000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000 Then
                NewCode = "0000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 10000 Then
                NewCode = "000" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 100000 Then
                NewCode = "00" & CStr((CodeNumber - 1) + LimitNumberIndex)
            ElseIf LimitNumberIndex + (CodeNumber - 1) < 1000000 Then
                NewCode = "0" & CStr((CodeNumber - 1) + LimitNumberIndex)
            Else
                NewCode = CStr((CodeNumber - 1) + LimitNumberIndex)
            End If
        End If


    End Sub

    Private Sub btnSearchEmpNo_Click(sender As Object, e As EventArgs) Handles btnSearchEmpNo.Click
        Me.Execute_Search()

    End Sub

    Private Sub txtFirstname_copy_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub txtFirstname_TextChanged_1(sender As Object, e As EventArgs) Handles txtFirstname.TextChanged
        Me.txtFirstname_view.Text = Me.txtFirstname.Text
    End Sub

    Private Sub txtLastname_TextChanged_1(sender As Object, e As EventArgs) Handles txtLastname.TextChanged
        Me.txtLastname_view.Text = Me.txtLastname.Text

    End Sub

    Private Sub comAccessRights_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comAccessRights.SelectedIndexChanged
        Me.txtAccessRights.Text = Me.comAccessRights.Text
    End Sub

    Private Sub btnAddOperator_Click(sender As Object, e As EventArgs) Handles btnAddOperator.Click
        Dim RecID As String = ""
        Dim ErrMessage = ""
        'Add user to Operators Table:
        'Add the current details found to the operators table - check if exists already -
        ' either perform INSERT operation OR UPDATE.
        'If logged in remotely - need to remove from local database ALSO.
        DBTable = "tblOperators"
        If (UCase(frmMainGIForm.myAccessRights) = "ADMIN") Or (UCase(frmMainGIForm.myAccessRights) = "SUPER") Then
            'TEST if USER NOT ALREADY in the OPERATORS TABLE:
            strEmpNo = Me.txtSearchEmpNo.Text

            SearchField = "EmpNo"
            SearchText = strEmpNo
            ReturnField = "ID"
            ReturnValue = ""
            strUserID = Me.txtID.Text
            If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
                OperatorExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "", False, MyMessage)
            Else
                OperatorExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, AllValues, AllFields, Criteria, "", False, MyMessage)
            End If
            If Len(MyMessage) > 0 Then
                Me.txtMessage.Text = MyMessage
            End If
            If OperatorExists Then
                MsgBox("Employee Already an Operator")
            Else
                'Remember - AllValues contains the values from the field positions of tblUsers still - NOT tblOperators - because not found !!!
                'strEmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "EmpNo")
                'strFirstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Firstname")
                'strLastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Lastname")
                'strWarehouse = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Warehouse")
                'strDepartment = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Department")
                'strJobTitle = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "JobTitle")
                'strDateStarted = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "DateStarted")
                'strUserBarcode = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "UserBarcode")
                'strUserBarcodeString = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "UserBarcodeString")
                'strDateOfBirth = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "DateOfBirth")
                'strComments = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Comments")
                'strUsername = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Username")
                'strPassword = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "Password")
                'strAccessRights = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "AccessRights")
                strDateCreated = CStr(Now())
                IsLoggedIN = "0"
                strSessionID = frmMainGIForm.mySessionID
                FieldValues = strUserID & "," & strUsername & "," & strFirstname & "," & strLastname & "," & strEmpNo
                FieldValues = FieldValues & "," & strDateCreated
                FieldValues = FieldValues & "," & IsLoggedIN
                FieldValues = FieldValues & "," & strComments
                FieldValues = FieldValues & "," & strAccessRights
                FieldValues = FieldValues & "," & strSessionID

                If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
                    Criteria = ""
                    InsertOK = InsertUpdateRecord(False, DBName, "", DBTable, "", FieldValues, Criteria, "ID")
                    If InsertOK Then
                        txtMessage.AppendText("INSERT Created -  " & strEmpNo & " OK")
                    Else
                        txtMessage.AppendText("Error with EMP NO: " & strEmpNo)
                    End If
                Else
                    If frmMainGIForm.myLoggedIn Then
                        Criteria = ""
                        InsertOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, RecID, DBTable,
                                                                     "", FieldValues, Criteria, "ID", ErrMessage, False)
                        If InsertOK Then
                            txtMessage.AppendText("INSERT Created -  " & strEmpNo & " OK")
                            Me.btnAddOperator.Visible = False
                            Me.btnRemoveOperator.Visible = True
                            MsgBox("Employee " & strFirstname & " " & strLastname & " Added to Operators")
                        Else
                            txtMessage.AppendText("Error with EMP NO: " & strEmpNo)
                            MsgBox("Problem with Adding Employee to Operators")
                        End If
                    End If 'If logged_IN
                End If 'Database Type - ADODB or MySQL
            End If 'If USER EXISTS
        Else
            MsgBox("NOT ENOUGH USER RIGHTS")
            Exit Sub
        End If 'IF USER AT ADMIN OR SUPER LEVEL
    End Sub

    Private Sub btnRemoveOperator_Click(sender As Object, e As EventArgs) Handles btnRemoveOperator.Click
        'REMOVE USER FROM OPERATORS TABLE:
        'DELETE FROM tblOperators - using current details.
        'If logged in remotely - need to remove from local database ALSO.
        Dim Criteria As String = ""

        DBTable = "tblOperators"
        strEmpNo = Me.txtSearchEmpNo.Text

        SearchField = "EmpNo"
        SearchText = strEmpNo
        ReturnField = "ID"
        ReturnValue = ""
        strUserID = Me.txtID.Text
        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            OperatorExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "", False, MyMessage)
        Else
            OperatorExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue,
                                          AllValues, AllFields, Criteria, "", False, MyMessage)
        End If
        If Len(MyMessage) > 0 Then
            Me.txtMessage.Text = MyMessage
        End If

        'strEmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllValues, AllFields, "EmpNo")
        If Len(ReturnValue) > 0 And IsNumeric(ReturnValue) Then
            Criteria = "ID = " & CLng(ReturnValue)
            If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
                DeletedOK = DeleteRecord(DBName, frmMainGIForm.DBProvider, DBTable, Criteria)
            Else
                DeletedOK = DeleteMyRecord(DBTable, frmMainGIForm.myConnString, Criteria, MyMessage)
            End If
            If DeletedOK = True Then
                MsgBox("OK EMPLOYEE REMOVED FROM OPERATORS")
                Me.btnRemoveOperator.Visible = False
                Me.btnAddOperator.Visible = True
            Else
                MsgBox("Problem Removing Employee from Operators")
            End If
        End If

    End Sub

    Private Sub frmCreateUsers_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Me.txtSearchEmpNo.ContainsFocus Then
                Me.Execute_Search()
            End If
        End If
    End Sub

    Private Sub frmCreateUsers_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        'UPDATE NAMES with selected EXCEL sheet - either XLSX or CSV - should have ALIAS and NAME field names:
        ' Run through each NAME (employee Number) from the loaded sheet - compare against Firstname and Lastname to see if blank.
        '  OR if a different name ?? - ASK USER if they want to change to the NEW name ??
        'First Need to allow user to open up the EXCEL sheet location:
        UpdateNames()


    End Sub

    Sub UpdateNames()
        Dim TheFilename As String = ""
        Dim TheNames(,) As String
        Dim TotalFields As Long
        Dim TotalRows As Long
        Dim ErrMessage As String = ""
        Dim RowIDX As Long
        Dim strName As String = ""
        Dim strEmpno As String = ""
        Dim TitleName As String = ""
        Dim TitleEmpNo As String = ""
        Dim FirstAnswer As Integer = 0
        Dim SecondAnswer As Integer = 0
        Dim lngEmpNOColumn As Long = 0
        Dim lngNameColumn As Long = 0
        Dim strEmpNoCol As String = ""
        Dim strNameCol As String = ""
        Dim DBUsersTAble As String = ""
        'Dim DBRegisterTable As String = ""
        'Dim DBHistoryTable As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = "ID"
        Dim ExcludeFields As String = ""
        Dim Criteria As String = ""
        Dim SearchCriteria As String = ""
        Dim InsertMessage As String = ""
        Dim FieldNames As String = ""
        Dim FieldValues As String = ""
        Dim ReturnValue As String = ""
        Dim Errors As String = ""
        Dim ExtractErrors As String = ""
        Dim FindUser As Boolean = False
        Dim AllFieldValues As Object() = Nothing
        Dim AllFieldNames() As String = Nothing
        Dim SortField As String = ""
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim SheetFirstname As String = ""
        Dim SheetLastname As String = ""
        Dim SpacePos As Integer = 0
        Dim UpdatedOK As Boolean = False
        Dim strNewEmpNO As String = ""
        Dim Percentage As Single = 0.0F
        Dim OldEmpNo As String = ""
        Dim NewEmpNo As String = ""
        Dim lngEmpNo As Long = 0
        Dim strNewName As String = ""
        Dim strComments As String = ""
        Dim strDBEmpNo As String = ""
        Dim strUserID As String = ""
        Dim strUsername As String
        Dim strPassword As String
        Dim strAccessRights As String
        Dim RecID As String = ""
        Dim dtDateModified As DateTime

        ReDim TheNames(1, 1)
        TotalRows = 0
        TotalFields = 0
        DBUsersTAble = "tblUsers"
        Me.txtOutput.Visible = True
        Me.txtOutput.Text = ""
        FieldNames = "EmpNo,Firstname,Lastname"
        'DBRegisterTable = "tblRegister" 'may need to compare to CURRENT DATE - as the names can change Daily - particularly AGency.
        'DBHistoryTable = "tblHistory" 'careful - as what if the names have changed the FOLLOWING DAY ?? will need to check the DATEOut also.
        dtDateModified = Now()

        TheFilename = DanG_DB_Tools.Get_Filename()

        TotalRows = Module_DanG_MySQL_Tools.CSVFileToArray(TheNames, TheFilename, TotalFields, ErrMessage)

        RowIDX = 0
        If TotalRows = 0 Then
            MsgBox("No names returned")
            Exit Sub
        End If
        TitleName = TheNames(1, 0) 'ALIAS
        TitleEmpNo = TheNames(0, 0) 'NAME / Emp No
        lngEmpNOColumn = 0
        lngNameColumn = 1
        If UCase(TitleEmpNo) = "NAME" Then
            'OK EmpNo Column
            lngEmpNOColumn = 0
        Else
            FirstAnswer = MsgBox("Is this the Employee Number Column ?", vbYesNo, "Emp No Column ? " & TitleEmpNo)
            If FirstAnswer = vbNo Then
                strEmpNoCol = InputBox("Enter the Employee Number Column: 0 to " & CStr(TotalFields))
                If Len(strEmpNoCol) = 0 Then
                    Exit Sub
                End If
                If IsNumeric(CLng(strEmpNoCol)) Then
                    'OK VALID number.
                    lngEmpNOColumn = CLng(strEmpNoCol)
                Else
                    MsgBox("Not a valid column number")
                    Exit Sub
                End If
            End If
        End If
        If UCase(TitleName) = "ALIAS" Then
            'OK NAME column
            lngNameColumn = 1
        Else
            SecondAnswer = MsgBox("Is this the Name Column ?", vbYesNo, "Name Column ? " & TitleName)
            If SecondAnswer = vbNo Then
                strNameCol = InputBox("Enter the Name Column: 0 to " & CStr(TotalFields))
                If Len(strNameCol) = 0 Then
                    Exit Sub
                End If
                If IsNumeric(CLng(strNameCol)) Then
                    'OK VALID number.
                    lngNameColumn = CLng(strNameCol)
                Else
                    MsgBox("Not a valid column number")
                    Exit Sub
                End If
            End If
        End If
        ExcludeFields = "ID"
        Criteria = ""
        Do While RowIDX <= TotalRows
            strEmpno = TheNames(lngEmpNOColumn, RowIDX) 'From Loaded sheet - will either have AGY at start or an ECP number without leading zeros.
            strName = TheNames(lngNameColumn, RowIDX) 'From Loaded Sheet
            SpacePos = InStr(strName, " ")
            SheetFirstname = ""
            SheetLastname = ""
            If SpacePos > 0 Then
                SheetFirstname = Mid(strName, 1, SpacePos - 1)
                SheetLastname = Mid(strName, SpacePos + 1, Len(strName))
            Else
                SheetFirstname = strName
            End If
            'OK now we have the two values - name and Emp no. Search database for EmpNO.
            'CONVERT the EMployee Number from the sheet to the same format as saved in tblUsers i.e. ADD ZEROS to beginning to make 7-DIGITS.

            SearchField = "EmpNo" 'from tblUsers
            Errors = ""
            lngEmpNo = 0
            If UCase(Mid(strEmpno, 1, 3)) = "AGY" Then
                lngEmpNo = CInt(Mid(strEmpno, 4, Len(strEmpno)))
                Convert_OldEmpNo_NewEmpNo("AGENCY", lngEmpNo, OldEmpNo, NewEmpNo)
            Else
                If IsNumeric(strEmpno) Then
                    lngEmpNo = CInt(strEmpno)
                    Convert_OldEmpNo_NewEmpNo("ECP", lngEmpNo, OldEmpNo, NewEmpNo)
                End If

            End If

            SearchText = NewEmpNo
            SearchCriteria = ""
            strNewEmpNO = NewEmpNo
            If Len(NewEmpNo) > 0 Then
                FindUser = Find_myQuery(frmMainGIForm.myConnString, DBUsersTAble, SearchField, SearchText, "STRING", ReturnField,
                                    ReturnValue, AllFieldValues, AllFieldNames, SearchCriteria, SortField, False, Errors)
            Else
                FindUser = False
            End If
            If Len(Errors) > 0 Then
                MsgBox("Error: " & Errors)
                Exit Sub
            End If
            ExtractErrors = ""

            If FindUser Then 'Userbarcode is found in tblUSers that matches the first col on the sheet - Employee Number
                'OK now check if the Firstname and Lastname is blank :
                strUserID = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "ID", ExtractErrors)
                Firstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Firstname", ExtractErrors)
                Lastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Lastname", ExtractErrors)
                strDBEmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "EmpNo", ExtractErrors) 'NEW 7 DIGITS
                strComments = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Comments", ExtractErrors)
                strNewName = SheetFirstname & " " & SheetLastname
                If Len(ExtractErrors) > 0 Then
                    MsgBox("Error: " & ExtractErrors)
                    Exit Sub
                End If
                If Len(Firstname) > 0 Then 'NAME ALREADY IN DATABASE:
                    'CHECK if the name matches:
                    If SpacePos > 0 Then 'The Name - 2nd Col - has a space
                        If UCase(SheetFirstname) = UCase(Firstname) And UCase(SheetLastname) = UCase(Lastname) Then
                            'Nothing to update
                        Else
                            'Firstname is DIFFERENT:
                            FirstAnswer = MsgBox("New NAME (from sheet) is different - UPDATE name in database ???", vbYesNoCancel, "New NAME (from sheet): " & strNewName & " NOT match " & Firstname & " " & Lastname)
                            If FirstAnswer = vbYes Then
                                strComments = strComments & ". INSERTED " & strNewName & " AT " & CStr(Now()) & " - UPDATED BY: " & frmMainGIForm.myFirstname & " " & frmMainGIForm.myLastname
                                FieldNames = "Firstname,Lastname,Comments"
                                'FieldValues = Chr(34) & strNewEmpNO & Chr(34)
                                FieldValues = SheetFirstname
                                FieldValues = FieldValues & "," & SheetLastname
                                FieldValues = FieldValues & "," & strComments
                                Criteria = "EmpNo = " & strDBEmpNo
                                UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, RecID, DBUsersTAble,
                                                                                 FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage,
                                                                                 False)
                                If UpdatedOK Then
                                    'MsgBox("OK changed Firstname")
                                    Me.txtOutput.AppendText(vbCrLf & "Changed NAME:" & strNewEmpNO & " " & strNewName & vbCrLf)
                                Else
                                    MsgBox("Could NOT change the Firstname")
                                    Me.txtOutput.AppendText(vbCrLf & "NO CHANGE to " & strDBEmpNo & " " & Firstname & " " & Lastname & vbCrLf)
                                End If
                            End If
                            If FirstAnswer = vbCancel Then
                                Exit Sub
                            End If
                        End If
                        If UCase(SheetLastname) = UCase(Lastname) Then
                            'nothing to update
                        Else
                            'Lastname is DIFFERENT:
                        End If
                    Else 'No SPACE FOUND:
                        'Just one name contained in FIRSTNAME from DATABASE:
                        If UCase(SheetFirstname) = UCase(Firstname) And UCase(SheetLastname) = UCase(Lastname) Then
                            'OK nothing to update
                        Else
                            FirstAnswer = MsgBox("New NAME (from sheet) is different - UPDATE name in database ???", vbYesNoCancel, "New NAME (from sheet): " & strNewName & " NOT match " & Firstname & " " & Lastname)
                            If FirstAnswer = vbYes Then
                                strComments = strComments & ". INSERTED " & strNewName & " AT " & CStr(Now()) & " - UPDATED BY: " & frmMainGIForm.myFirstname & " " & frmMainGIForm.myLastname
                                FieldNames = "Firstname,Lastname,Comments"
                                'FieldValues = Chr(34) & strNewEmpNO & Chr(34)
                                FieldValues = SheetFirstname
                                FieldValues = FieldValues & "," & SheetLastname
                                FieldValues = FieldValues & "," & strComments
                                Criteria = "EmpNo = " & strDBEmpNo
                                UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, RecID, DBUsersTAble,
                                                                                 FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage,
                                                                                 False)
                                If UpdatedOK Then
                                    'MsgBox("OK changed Firstname")
                                    Me.txtOutput.AppendText(vbCrLf & "Changed NAME:" & strNewEmpNO & " " & strNewName & vbCrLf)
                                Else
                                    MsgBox("Could NOT change the Firstname")
                                    Me.txtOutput.AppendText(vbCrLf & "NO CHANGE to " & strDBEmpNo & " " & Firstname & " " & Lastname & vbCrLf)
                                End If
                            End If
                            If FirstAnswer = vbCancel Then
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    'Firstname is BLANK in DATABASE:
                    'Need to show the EMployee ID in the comments also.
                    strComments = strComments & ". INSERTED " & strNewName & " AT " & CStr(Now()) & " - UPDATED BY: " & frmMainGIForm.myFirstname & " " & frmMainGIForm.myLastname
                    FieldNames = "Firstname,Lastname,Comments"
                    'FieldValues = Chr(34) & strNewEmpNO & Chr(34)
                    FieldValues = SheetFirstname
                    FieldValues = FieldValues & "," & SheetLastname
                    FieldValues = FieldValues & "," & strComments
                    Criteria = "EmpNo = " & strDBEmpNo
                    UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, RecID, DBUsersTAble,
                                                                                 FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage,
                                                                                 False)
                    If UpdatedOK Then
                        'MsgBox("OK INSERTED NAME")
                        Me.txtOutput.AppendText(vbCrLf & "INSERTED NAME:" & strNewEmpNO & " " & strNewName & vbCrLf)
                    Else
                        MsgBox("Could NOT change the NAME")
                        Me.txtOutput.AppendText(vbCrLf & "NAME NOT INSERTED " & strNewEmpNO & " " & Firstname & " " & Lastname & vbCrLf)
                    End If
                End If 'Length of Firstname > 0
            Else
                'CANNOT FIND THE EMPLOYEE NUMBER in DATABASE:
                strUsername = ""
                If Len(SheetFirstname) > 0 And Len(SheetLastname) > 0 Then
                    strUsername = LCase(Mid(SheetFirstname, 1, 1))
                    strUsername = LCase(strUsername & SheetLastname)
                    strPassword = "password"
                    strAccessRights = "normal"
                    FieldNames = "EmpNo,Firstname,Lastname,Username,Password,AccessRights,DateModified"
                    FieldValues = NewEmpNo
                    FieldValues = FieldValues & "," & SheetFirstname
                    FieldValues = FieldValues & "," & SheetLastname
                    FieldValues = FieldValues & "," & strUsername
                    FieldValues = FieldValues & "," & strPassword
                    FieldValues = FieldValues & "," & strAccessRights
                    FieldValues = FieldValues & "," & dtDateModified

                    UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, False, RecID, DBUsersTAble,
                                                                 FieldNames, FieldValues, Criteria, ExcludeFields, ErrMessage, False)
                    If UpdatedOK Then
                        Me.txtOutput.AppendText(vbCrLf & "ADDED: " & strEmpno & " " & strName)
                    Else
                        Me.txtOutput.AppendText(vbCrLf & "FAILED to ADD: " & strEmpno & " " & strName)
                    End If
                End If
            End If

            Percentage = (RowIDX / TotalRows) * 100
            pbBar_Accounts.Value = CInt(Percentage)
            Application.DoEvents()
            RowIDX = RowIDX + 1
        Loop

        MsgBox(" OK FINISHED UPDATE")

    End Sub

    Private Sub frmCreateUsers_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
            Me.btnUpdate.Visible = True
        Else
            Me.btnUpdate.Visible = False
        End If


    End Sub

    Private Sub btnAddExtraLogin_Click(sender As Object, e As EventArgs) Handles btnAddExtraLogin.Click
        'ADD EXTRA LOGIN:
        Dim SearchText As String
        Dim SearchCriteria As String
        Dim DBTable As String
        Dim SearchField As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim AllFieldValues() As Object
        Dim AllFieldNames() As String
        Dim SortField As String
        Dim Errors As String
        Dim FindUser As Boolean
        Dim FindLogin As Boolean
        Dim strNewEmpNO As String
        Dim NewEmpNo As String
        Dim lngEmpNo As Long
        Dim OldEmpNo As String
        'Need to work out next login number:
        'Search tblOperators:
        DBTable = "tblOperators"
        lngEmpNo = 0
        If UCase(Mid(strEmpNo, 1, 3)) = "AGY" Then
            lngEmpNo = CInt(Mid(strEmpNo, 4, Len(strEmpNo)))
            Convert_OldEmpNo_NewEmpNo("AGENCY", lngEmpNo, OldEmpNo, NewEmpNo)
        Else
            If IsNumeric(strEmpNo) Then
                lngEmpNo = CInt(strEmpNo)
                Convert_OldEmpNo_NewEmpNo("ECP", lngEmpNo, OldEmpNo, NewEmpNo)
            End If

        End If


        SearchText = NewEmpNo
        SearchCriteria = ""
        strNewEmpNO = NewEmpNo
        If Len(NewEmpNo) > 0 Then
            FindUser = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField,
                                ReturnValue, AllFieldValues, AllFieldNames, SearchCriteria, SortField, False, Errors)
        Else
            FindUser = False
        End If


    End Sub
End Class