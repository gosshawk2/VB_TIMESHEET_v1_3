Public Class frmMysqlLogin

    '****************************************************************************************************************************************************************
    '*
    '*  Program Title:                                            VB GoodsIn Timesheet v1.1
    '*
    '*  Language:                                                 VB.NET - Visual Studio 2015 - © 2017 Microsoft
    '*
    '*  AUTHOR:                                                   Daniel Goss
    '*
    '*  Copyright:                                                Copyright 2016-2018
    '*
    '*  Module Code:                                              Form: frmMysqlLogin
    '*
    '*  External Name:                                            Form_Mysql_Login.vb
    '*
    '*  Module last amended:                                      17/06/2018 20:00
    '*
    '*      TITLE CHANGED TO REFLECT THIS IS THE NEW VERSION IS UNDER TESTING.
    '*
    '*      Corrected Password problem : a user could log in using any password ! - now corrected - Exact password must be used.
    '*
    '*                                                              LINK to frmMainGIForm
    '****************************************************************************************************************************************************************

    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Me.Close()

    End Sub

    Private Sub btnLOGIN_Click(sender As Object, e As EventArgs) Handles btnLOGIN.Click
        LoginProcess()
    End Sub

    Sub LoginProcess()
        'LOGIN to database - gather user entry and pass to setupConnection method:
        'Login locally as well as remotely:
        Dim Server As String = ""
        Dim username As String = ""
        Dim Password As String = ""
        Dim Database As String = ""
        Dim PORT As String = ""
        Dim myMessage As String = ""
        Dim EmpNo As String = ""
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim ConnectMessage As String = ""
        Dim OPTable As String = "tblOperators"
        Dim UserTable As String = "tblUsers"
        Dim DBTable As String = ""
        Dim theUserID As String = ""
        Dim Criteria As String = ""
        Dim theAccessLevel As String = ""
        Dim SearchField As String = ""
        Dim SearchValue As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim AllFieldNames As String() = Nothing
        Dim AllFieldValues As Object = Nothing
        Dim FindUser As Boolean = False
        Dim FindPassword As Boolean = False
        Dim connString As String = ""
        Dim ConnMessage As String = ""
        Dim LocalConnString As String = ""
        Dim LocalConnMessage As String = ""
        Dim FindOperator As Boolean = False
        Dim FoundPassword As String = ""

        If Len(Me.txtServer.Text) > 0 Then
            frmMainGIForm.myServer = Me.txtServer.Text
        Else
            frmMainGIForm.myServer = "127.0.0.1"
        End If
        If Len(Me.txtDatabase.Text) > 0 Then
            frmMainGIForm.myDatabase = Me.txtDatabase.Text
        Else
            frmMainGIForm.myDatabase = "timesheet"
        End If
        If Len(Me.txtPort.Text) > 0 Then
            frmMainGIForm.myPort = Me.txtPort.Text
        Else
            frmMainGIForm.myPort = "3306"
        End If
        If Len(Me.txtUsername.Text) > 0 Then
            frmMainGIForm.myUsername = Me.txtUsername.Text 'FROM USER JUST LOGGED IN (OPERATOR)
        Else
            MsgBox("Username must be entered")
            Application.Exit()
        End If
        If Len(Me.txtPassword.Text) > 0 Then
            frmMainGIForm.myPassword = Me.txtPassword.Text 'FROM USER JUST LOGGED IN (OPERATOR)
        Else
            MsgBox("Password cannot be blank")
            Application.Exit()
        End If
        DBTable = "tblUsers"
        Server = frmMainGIForm.myServer 'FROM LOGIN TEXT BOX
        Password = frmMainGIForm.myPassword 'FROM USER JUST LOGGED IN
        Database = frmMainGIForm.myDatabase 'FROM LOGIN TEXT BOX
        username = frmMainGIForm.myUsername 'FROM USER JUST LOGGED IN
        PORT = frmMainGIForm.myPort 'FROM PORT TEXT JUST LOGGED IN
        SearchField = "Username"
        SearchValue = username
        ReturnField = "ID"
        Criteria = ""
        Me.TopMost = False
        frmMainGIForm.SYSTEMUsername = "giuser"
        frmMainGIForm.SYSTEMPassword = "giuser"
        Module_DanG_MySQL_Tools.myChecked = TestLogin(Server, Database, frmMainGIForm.SYSTEMUsername, frmMainGIForm.SYSTEMPassword, PORT, "timesheet", "tblDeliveryInfo", "tblDeliveryInfo", myMessage, ConnectMessage)
        If Module_DanG_MySQL_Tools.myChecked Then

            'Now check corresponding USER details in tblUsers and get the UserID:
            connString = setupMySQLconnection(Server, Database, frmMainGIForm.SYSTEMUsername, frmMainGIForm.SYSTEMPassword, PORT, ConnMessage)

            LocalConnString = setupMySQLconnection("127.0.0.1", Database, frmMainGIForm.SYSTEMUsername, frmMainGIForm.SYSTEMPassword, "3306", LocalConnMessage)

            SearchField = "Username"
            SearchValue = username


            FindOperator = Module_DanG_MySQL_Tools.Find_myQuery(connString, OPTable, SearchField, SearchValue,
                                                             "STRING", ReturnField, ReturnValue, AllFieldValues, AllFieldNames, Criteria,
                                                            "Username", False, myMessage, "=")
            If FindOperator = True Then
                'Load up the preferences from the local ACCESS TABLE ?? - no from MySQL table - as that is central.

                EmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "EmpNo")
                theUserID = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "UserID")
                SearchField = "ID"
                SearchValue = CLng(theUserID)
                ReturnField = "ID"
                ReturnValue = ""
                'Criteria = "password = '" & Password & "'"
                Criteria = ""
                FindUser = Module_DanG_MySQL_Tools.Find_myQuery(connString, UserTable, SearchField, SearchValue,
                                                             "INTEGER", ReturnField, ReturnValue, AllFieldValues, AllFieldNames, Criteria,
                                                            "Username", False, myMessage, "=")
                If FindUser = True Then
                    'Now test if their password matches :

                    'theUserID = ReturnValue
                    frmMainGIForm.myUserID = ReturnValue
                    frmMainGIForm.myConnString = connString
                    frmMainGIForm.myLOCALConnString = LocalConnString
                    theAccessLevel = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "AccessRights")
                    EmpNo = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "EmpNo")
                    Firstname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Firstname")
                    Lastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Lastname")
                    'frmMainGIForm.myUsername = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Username")
                    FoundPassword = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllFieldValues, AllFieldNames, "Password")
                    frmMainGIForm.myAccessRights = theAccessLevel
                    frmMainGIForm.myEmpNo = EmpNo
                    frmMainGIForm.myFirstname = Firstname
                    frmMainGIForm.myLastname = Lastname
                    frmMainGIForm.myUserID = theUserID

                    If FoundPassword = Password Then
                        'MsgBox("SUCCESS !!!")
                        frmMainGIForm.Initialised = True
                        frmMainGIForm.myLoggedIn = True
                        frmMainGIForm.myUsername = username
                        Module_DanG_MySQL_Tools.mylogged_in = True
                        'MsgBox("YES - Successful Login ! - WELCOME to the TIMESHEET APP" & frmMainGIForm.myVersion)
                        'MsgBox("Localhost: " & frmMainGIForm.myLOCALConnString & " " & frmMainGIForm.myAccessRights)
                    Else
                        MsgBox("Password does NOT MACTCH")
                        Application.Exit()
                    End If

                Else
                    myMessage = "Operator Found in Operators Table"
                    MsgBox("BUT Corresponding USER not Found in Users table")


                End If


            Else
                MsgBox("Sorry - Your name is not in the Operators Table")
                Application.Exit()
            End If

        Else
            MsgBox("LOGIN FAILURE: " & ConnectMessage)
            If Len(myMessage) > 0 Then
                MsgBox("Login Error: " & myMessage)
            End If
            frmMainGIForm.myLoggedIn = False
            Application.Exit()

        End If
        Me.Close()

    End Sub

    Private Sub btnCANCEL_Click(sender As Object, e As EventArgs) Handles btnCANCEL.Click
        'CANCEL: - send a message back to the main program first - so it does not activate the SEssion Save routine.
        frmMainGIForm.EXITProcess = True
        Application.Exit()

    End Sub

    Private Sub rb6080_CheckedChanged(sender As Object, e As EventArgs) Handles rb6080.CheckedChanged
        txtPort.Text = "6080"
    End Sub

    Private Sub rb3306_CheckedChanged(sender As Object, e As EventArgs) Handles rb3306.CheckedChanged
        txtPort.Text = "3306"
    End Sub

    Private Sub rb3307_CheckedChanged(sender As Object, e As EventArgs) Handles rb3307.CheckedChanged
        txtPort.Text = "3307"
    End Sub

    Private Sub frmMysqlLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DBName As String = ""
        Dim DBTable As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim ALLFieldNames As String() = Nothing
        Dim ALLFieldValues As Object = Nothing
        Dim strHostname As String = ""
        Dim strPortNumber As String = ""
        Dim Criteria As String = ""
        Dim UserExists As Boolean = False
        Dim myMessage As String = ""
        Dim NumRows As Integer = 0
        Dim PrefArray As String(,) = Nothing
        Dim TotalFields As Integer = 0
        Dim FormTitle As String = ""
        Dim ComputerName As String

        DBTable = "tblPreferences"
        DBName = "GoodsInTimesheetRecords.accdb"

        ComputerName = frmMainGIForm.GetComputerName()
        If UCase(ComputerName) = "ECP-99-KPATEL" Then
            btnAutoLOGIN.Visible = True
        End If

        FormTitle = Me.Text & " " & frmMainGIForm.myVersion
        Me.Text = FormTitle
        txtUsername.Focus()
        SearchField = "UserID"
        SearchText = "0"
        ReturnField = "Server_IPv4Addr"
        ReturnValue = ""
        strHostname = frmMainGIForm.myServer
        strPortNumber = frmMainGIForm.myPort
        Criteria = ""
        If System.IO.File.Exists("CSV Preference Files\Preferences.csv") Then
            NumRows = Module_DanG_MySQL_Tools.CSVFileToArray(PrefArray, "CSV Preference Files\Preferences.csv", TotalFields)
            'PrefArray(FieldNo,RowNo) so PrefArray(0,0) is the ID label. (0,1) is the ID number from the next row. so (2,1) should give the IPv4Addr 
            'UserExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "INTEGER", ReturnField, ReturnValue, ALLFieldValues, ALLFieldNames, Criteria, "", "", False, myMessage)
            ReturnValue = Module_DanG_MySQL_Tools.GetFieldValue_From_Fieldname(PrefArray, 1, "Server_IPv4Addr")

            strHostname = ReturnValue
            strPortNumber = Module_DanG_MySQL_Tools.GetFieldValue_From_Fieldname(PrefArray, 1, "PortNumber")
        Else
            MsgBox("CANNOT FIND PREFERENCES CSV FILE")
        End If
        Me.txtServer.Text = strHostname
        Me.txtPort.Text = strPortNumber
        Me.TopMost = True
        Me.btnAdvanced.Text = "Advanced"
        Me.txtServer.Visible = False
        Me.txtDatabase.Visible = False
        Me.txtPort.Visible = False
        lblServer.Visible = False
        lblDatabase.Visible = False
        lblUsername.Visible = True
        lblPassword.Visible = True
        lblPort.Visible = False
        Me.rb3306.Visible = False
        Me.rb3307.Visible = False
        Me.rb6080.Visible = False
        Me.txtUsername.Visible = True
        Me.txtPassword.Visible = True

        Me.btnLOGIN.Focus()
        txtUsername.Focus()
        KeyPreview = True
    End Sub

    Private Sub btnAdvanced_Click(sender As Object, e As EventArgs) Handles btnAdvanced.Click
        Dim TheComputerName As String = ""

        If Me.btnAdvanced.Text = "Advanced" Then
            Me.btnAdvanced.Text = "Simple"
            lblServer.Visible = True
            lblDatabase.Visible = True
            lblUsername.Visible = True
            lblPassword.Visible = True
            lblPort.Visible = True
            Me.txtServer.Visible = True
            Me.txtDatabase.Visible = True
            Me.txtPort.Visible = True
            Me.txtUsername.Visible = True
            Me.txtPassword.Visible = True
            Me.txtPort.Visible = True
            Me.rb3306.Visible = True
            Me.rb3307.Visible = True
            Me.rb6080.Visible = True
            'Me.btnAdvanced.Visible = True
            Me.btnRemote.Visible = True
            Me.btnLocal.Visible = True
        Else
            Me.btnAdvanced.Text = "Advanced"
            lblServer.Visible = False
            lblDatabase.Visible = False
            lblUsername.Visible = True
            lblPassword.Visible = True
            lblPort.Visible = False
            Me.txtServer.Visible = False
            Me.txtDatabase.Visible = False
            Me.txtPort.Visible = False
            Me.txtUsername.Visible = True
            Me.txtPassword.Visible = True
            Me.txtPort.Visible = False
            Me.rb3306.Visible = False
            Me.rb3307.Visible = False
            Me.rb6080.Visible = False
            'Me.btnAdvanced.Visible = False
            Me.btnRemote.Visible = True
            Me.btnLocal.Visible = False
        End If
        TheComputerName = frmMainGIForm.GetComputerName
        'If UCase(TheComputerName) = UCase("ecp-99-kpatel") Then
        Me.btnLocal.Visible = True
        'Else
        'Me.btnLocal.Visible = False
        'End If
    End Sub

    Private Sub lblServer_Click(sender As Object, e As EventArgs) Handles lblServer.Click

    End Sub

    Private Sub frmMysqlLogin_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.txtUsername.Focus()

            LoginProcess()
        End If
    End Sub

    Private Sub btnRemote_Click(sender As Object, e As EventArgs) Handles btnRemote.Click
        Dim strHostname As String = ""
        Dim strPortNumber As String = ""
        Dim DBTable As String = ""
        Dim DBName As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim Criteria As String = ""
        Dim NumRows As Integer = 0
        Dim PrefArray As String(,) = Nothing
        Dim TotalFields As Integer = 0
        'Populate fields with REMOTE settings:
        DBTable = "tblPreferences"
        DBName = "GoodsInTimesheetRecords.accdb"

        SearchField = "UserID"
        SearchText = "0"
        ReturnField = "Server_IPv4Addr"
        ReturnValue = ""
        strHostname = frmMainGIForm.myServer
        strPortNumber = frmMainGIForm.myPort
        Criteria = ""
        If System.IO.File.Exists("CSV Preference Files\Preferences.csv") Then
            NumRows = Module_DanG_MySQL_Tools.CSVFileToArray(PrefArray, "CSV Preference Files\Preferences.csv", TotalFields)
            'PrefArray(FieldNo,RowNo) so PrefArray(0,0) is the ID label. (0,1) is the ID number from the next row. so (2,1) should give the IPv4Addr 
            'UserExists = Find_me(DBName, "", DBTable, SearchField, SearchText, "INTEGER", ReturnField, ReturnValue, ALLFieldValues, ALLFieldNames, Criteria, "", "", False, myMessage)
            ReturnValue = Module_DanG_MySQL_Tools.GetFieldValue_From_Fieldname(PrefArray, 1, "Server_IPv4Addr")

            strHostname = ReturnValue
            strPortNumber = Module_DanG_MySQL_Tools.GetFieldValue_From_Fieldname(PrefArray, 1, "PortNumber")
        End If
        Me.txtServer.Text = strHostname
        Me.txtPort.Text = strPortNumber
    End Sub

    Private Sub btnLocal_Click(sender As Object, e As EventArgs) Handles btnLocal.Click
        Dim strHostname As String = ""
        Dim strPortNumber As String = ""
        'Populate fields with LOCAL settings:
        strHostname = "localhost"
        strPortNumber = "3306"
        Me.txtServer.Text = strHostname
        Me.txtPort.Text = strPortNumber
    End Sub

    Private Sub txtTitle_TextChanged(sender As Object, e As EventArgs) Handles txtTitle.TextChanged

    End Sub

    Private Sub frmMysqlLogin_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.txtUsername.Focus()

    End Sub

    Private Sub btnAutoLOGIN_Click(sender As Object, e As EventArgs) Handles btnAutoLOGIN.Click
        Me.txtUsername.Text = "giuser"
        Me.txtPassword.Text = "giuser"
    End Sub
End Class