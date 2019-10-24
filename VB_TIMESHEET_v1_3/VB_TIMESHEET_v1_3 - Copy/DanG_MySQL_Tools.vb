Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient
Imports System.Collections.Generic

'****************************************************************************************************************************************************************
'*
'*  Program Title:                                            VB GOODS INWARDS TIMESHEET V 1.0
'*
'*  Language:                                                 VB.NET - Visual Studio 2015 - © 2017 Microsoft
'*
'*  AUTHOR:                                                   Daniel Goss
'*
'*  Copyright:                                                Copyright © 2016-2018
'*
'*  Module Code:                                              Module: Module_DanG_MySQL_Tools
'*
'*  External Name:                                            DanG_MySQL_Tools.vb
'*
'*  Previous Amendment:                                       17/06/2018 17:50
'*                                                              LINK TO frmMainGIForm
'*
'*
'*  Module last amended:                                      22/07/2018 17:50
'*    Added code to GetFieldValue_From_Fieldname() to detect if value passed is nothing - if so EXIT function.
'*    Now using GetmyFields()
'*
'*                                                            10/AUG/2018 12:55
'* Modified InsertUpdateRecords() to include ElementStartIDX,EncaseFields,RemoveBadChars,IncludeComma,IncludeQuotes
'*
'****************************************************************************************************************************************************************

Module Module_DanG_MySQL_Tools
    Public myConnection As String = ""
    Public mylogged_in As Boolean = False
    Public myChecked As Boolean = False

    Function setupMySQLconnection(ByVal Server As String, ByVal DbaseName As String, ByVal USERNAME As String,
                                    ByVal password As String, ByVal port As String, ByRef Message As String) As String
        'Dim DatabaseName As String = "assetregister"
        'Dim server As String = "localhost"
        'Dim Port As String = "6080"
        'Dim User as String = "root"
        'Dim Values(,) As String
        Dim conn As MySqlConnection
        Dim myCMD As MySqlCommand
        Dim myDR As MySqlDataReader
        Dim Output As String
        Dim ZeroDatetime As Boolean

        Output = ""
        conn = New MySqlConnection()
        If Not conn Is Nothing Then conn.Close()
        ZeroDatetime = True
        'Output = String.Format("server={0}; user id={1}; password={2}; database={3}; port={4}; pooling=false; Convert Zero Datetime=True; Allow Zero Datetime=True;", Server, USERNAME, password, DbaseName, port)
        'Output = String.Format("server={0}; user id={1}; password={2}; database={3}; port={4}; pooling=false; convert zero datetime=True", Server, USERNAME, password, DbaseName, port)
        Output = String.Format("server={0}; user id={1}; password={2}; database={3}; Convert Zero Datetime={4}; port={5}; pooling=false", Server, USERNAME, password, DbaseName, ZeroDatetime, port)
        conn.ConnectionString = Output
        Try
            conn.Open()
            'myCMD = New MySqlCommand(SQLstr, conn)
            'myDR = myCMD.ExecuteReader()
            'While myDR.Read
            'values() = myDR(fieldname).ToString()
            'End While
            'MsgBox("Connected")
        Catch ex As Exception
            'MsgBox("Error in setupMySQLconnection: " & ex.Message)
            Message = "Error in setupMySQLconnection: " & ex.Message
            frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "setupMySQLconnection()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("AR_messages_v2.7.log", Application.StartupPath, Message, "CONNECTION to DB NOT SUCCESSFUL, Server=" & Server & ",DB Name= " & DbaseName & ", Username: " & USERNAME & ", Port= " & port)
        End Try
        conn.Close()

        setupMySQLconnection = Output

    End Function

    Function TestLogin(ByVal Server As String, ByVal DbaseName As String, ByVal USERNAME As String,
                       ByVal password As String, ByVal port As String, TestDBName As String, ByVal TestTable As String,
                       ByVal RealTableName As String, ByRef Message As String, MessageLogin As String) As Boolean
        Dim IsChecked As Boolean = False
        Dim IsConnectedToTestDB = False
        Dim IsConnectedToMainDB = False
        Dim conn As MySqlConnection = Nothing
        Dim myCMD As MySqlCommand = Nothing
        Dim myDR As MySqlDataReader = Nothing
        Dim connString As String = ""
        Dim SQLstr As String = ""
        Dim ConnMessage As String = ""

        'Dim Output As String

        Message = ""
        MessageLogin = ""
        If Len(TestDBName) = 0 Then
            'MsgBox("No Test TableName passed", vbOK, "Error ")
            Message = Message & vbCrLf & "No Test Database passed"
        End If
        If Len(Server) = 0 Then
            'MsgBox("No server passed", vbOK, "Error ")
            Message = Message & vbCrLf & "No server passed"
        End If
        If Len(DbaseName) = 0 Then
            'MsgBox("No database name passed", vbOK, "Error ")
            Message = Message & vbCrLf & "No database name passed"
        End If
        If Len(USERNAME) = 0 Then
            'MsgBox("No username passed", vbOK, "Error  ")
            Message = Message & vbCrLf & "No username passed"
        End If
        If Len(port) = 0 Then
            'MsgBox("No port passed", vbOK, "Error ")
            Message = Message & vbCrLf & "No port passed"
        End If
        If Len(Message) > 0 Then
            frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "TestLogin()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function
        End If
        SQLstr = "Select * FROM " & TestTable
        connString = setupMySQLconnection(Server, TestDBName, USERNAME, password, port, ConnMessage)
        If Len(ConnMessage) > 0 Then
            Message = Message & vbCrLf & " Connection Error: " & ConnMessage
            frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "TestLogin()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function
        End If
        'Output = String.Format("server={0}; user id={1}; password={2}; database={3}; port={4}; pooling=False", Server, USERNAME, password, DbaseName, port)
        Try
            conn = New MySqlConnection(connString)
            conn.Open()
            myCMD = New MySqlCommand(SQLstr, conn)
            myDR = myCMD.ExecuteReader()
            'While myDR.Read
            'values() = myDR(fieldname).ToString()
            'End While
            'MsgBox("Connected")
        Catch ex As Exception
            'MsgBox("Error In setupMySQLconnection: " & ex.Message)
            Message = "Error In Test Table with DB Login: " & ex.Message
            TestLogin = False
            frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "TestLogin()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function
        End Try
        conn.Close()
        MessageLogin = MessageLogin & vbCrLf & " OK - Test DB Connected OK."
        SQLstr = "Select * FROM " & RealTableName
        Message = "DB LOGIN SUCCESSFUL"
        frmMainGIForm.logger.logmessage("AR_messages_v2_7.log", Application.StartupPath, Message, "DB LOGIN SUCCESSFUL SQL=" & SQLstr)
        connString = setupMySQLconnection(Server, DbaseName, USERNAME, password, port, ConnMessage)
        Try
            conn = New MySqlConnection(connString)
            conn.Open()
            myCMD = New MySqlCommand(SQLstr, conn)
            myDR = myCMD.ExecuteReader()
            'While myDR.Read
            'values() = myDR(fieldname).ToString()
            'End While
            'MsgBox("Connected")
        Catch ex As Exception
            'MsgBox("Error In setupMySQLconnection: " & ex.Message)
            Message = "Exception Error In Test Asset-Register DB Login: " & ex.Message
            frmMainGIForm.logger.LogError("AR_Error_v2_7.log", Application.StartupPath, Message, "TestLogin()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            frmMainGIForm.logger.logmessage("AR_messages_v2_7.log", Application.StartupPath, Message, "DB LOGIN NOT SUCCESSFUL SQL=" & SQLstr)
            TestLogin = False
            Exit Function
        End Try
        IsChecked = True
        conn.Close()
        myDR.Close()
        TestLogin = IsChecked


    End Function

    Function saveSession(ByVal connString As String, ByVal UserID As String, ByVal Username As String,
                            ByVal ComputerName As String, ByVal IPv4Addr As String, ByVal IPv6Addr As String,
                            ByVal SessionID As String, ByVal AccessRights As String, ByRef MyMessage As String,
                            ByVal EmpNo As String, ByVal Firstname As String, ByVal Lastname As String,
                            ByVal Comments As String, ByVal Location As String, ByVal WorkstationName As String,
                            ByVal ProgramVersion As String,
                            Optional ByRef MyOpsMessage As String = "") As Boolean
        Dim SessionUpdatedOK As Boolean = False
        Dim IsOpsUpdatedOK As Boolean = False
        Dim IsOpsOnlineUpdatedOK As Boolean = False
        Dim ConnMessage As String = ""
        Dim DBTable As String = ""
        Dim FieldValues As String = ""
        Dim OpsOnlineFieldValues As String = ""
        Dim FieldNames As String = ""
        Dim OpsOnlineFieldNames As String = ""
        Dim ExcludeFields As String = ""
        Dim OpsOnlineExcludeFields As String = ""
        Dim Criteria As String = ""
        Dim OpsOnlineCriteria As String = ""
        Dim strComputerName As String = ""
        Dim strIP4Addr As String = ""
        Dim strIP6Addr As String = ""
        Dim strGUID As String = ""
        Dim strSessionID As String = ""
        Dim strAccessRights As String = ""
        Dim strUsername As String = ""
        Dim strDateLoggedIn As String = ""
        Dim strDateLoggedOut As String = ""
        Dim strUserID As String = ""
        Dim strIsLoggedIn As String = ""
        Dim strLoggedInDuration As String = ""
        Dim OpsDBTable As String = ""
        Dim OpsFieldnames As String = ""
        Dim OpsExcludeFields As String = ""
        Dim OpsFieldValues As String = ""
        Dim OpsCriteria As String = ""
        Dim dtMinutesDiff As Long = 0
        Dim dtDateLoggedIn As DateTime
        Dim dtDateLoggedOut As DateTime
        Dim strHours As String = ""
        Dim strMins As String = ""
        Dim OnlineTable As String = ""
        Dim strEmpNo As String = ""
        Dim strFirstname As String = ""
        Dim strLastname As String = ""
        Dim strComments As String = ""
        Dim strLocation As String = ""
        Dim strWorkstationName As String = ""
        Dim strProgramVersion As String

        saveSession = False
        DBTable = "tblSessions"
        OpsDBTable = "tblOperators"
        OnlineTable = "tblOperatorsOnline"
        strUserID = UserID
        'Need to get the Login Username INDEX from login global var - not the value from the table.
        strUsername = Username
        strEmpNo = EmpNo
        strFirstname = Firstname
        strLastname = Lastname
        strComments = Comments
        strLocation = Location
        strComputerName = ComputerName
        strWorkstationName = WorkstationName
        strIP4Addr = IPv4Addr
        strIP6Addr = IPv6Addr
        strSessionID = SessionID
        strAccessRights = AccessRights
        strIsLoggedIn = "YES"
        strDateLoggedIn = Now().ToString("yyyy-MM-dd HH:mm:ss")
        strDateLoggedOut = "1970-01-01 01:00:00"
        strProgramVersion = ProgramVersion
        'dtDateLoggedIn = CDate(strDateLoggedIn)
        'dtDateLoggedOut = CDate(strDateLoggedOut)
        'dtMinutesDiff = DateDiff(DateInterval.Minute, dtDateLoggedIn, dtDateLoggedOut)
        'strHours = CStr(Math.Abs(dtMinutesDiff) / 60)
        'strMins = CStr(Math.Abs(dtMinutesDiff) Mod 60)

        strLoggedInDuration = strHours & " Hours " & strMins & " Mins"
        FieldNames = "ID,UserID,Username,ComputerName,IPv4Addr,IPv6Addr,SessionID,AccessRights,LogInDateTime,LogOffDateTime,IsLoggedIn"
        FieldNames = FieldNames & ",LoggedInDuration,ProgramVersion"
        'FieldNames = Module_DanG_MySQL_Tools.GetMyFields("tblSessions", frmMainGIForm.myConnString, "")
        ExcludeFields = "ID"
        FieldValues = strUserID
        FieldValues = FieldValues & "," & strUsername
        FieldValues = FieldValues & "," & strComputerName
        FieldValues = FieldValues & "," & strIP4Addr
        FieldValues = FieldValues & "," & strIP6Addr
        FieldValues = FieldValues & "," & strSessionID
        FieldValues = FieldValues & "," & strAccessRights
        FieldValues = FieldValues & "," & strDateLoggedIn
        FieldValues = FieldValues & "," & strDateLoggedOut
        FieldValues = FieldValues & "," & strIsLoggedIn
        FieldValues = FieldValues & "," & strLoggedInDuration
        FieldValues = FieldValues & "," & strProgramVersion

        SessionUpdatedOK = InsertUpdateRecords_Using_Parameters(connString, False, "", DBTable, FieldNames, FieldValues, Criteria,
                                                                ExcludeFields)

        'SessionUpdatedOK = InsertUpdateMyRecord(False, connString, DBTable, FieldNames, FieldValues, MyMessage, Criteria, ExcludeFields)
        If Not SessionUpdatedOK Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "saveSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If
        saveSession = SessionUpdatedOK
        'UPDATE the OPERATORS table:
        '=============================
        OpsCriteria = "Username = " & Chr(34) & strUsername & Chr(34)
        OpsFieldnames = "IsLoggedIn,SessionID"
        OpsExcludeFields = "ID,UserID,Username,Firstname,Lastname,EmpNo,DateCreated,Comments,AccessRights"
        OpsFieldValues = strIsLoggedIn & "," & strSessionID

        IsOpsUpdatedOK = InsertUpdateRecords_Using_Parameters(connString, True, "", OpsDBTable, OpsFieldnames, OpsFieldValues, OpsCriteria, ExcludeFields, MyOpsMessage)
        'IsOpsUpdatedOK = InsertUpdateMyRecord(True, connString, OpsDBTable, OpsFieldnames, OpsFieldValues, MyOpsMessage, OpsCriteria, OpsExcludeFields)
        If Not IsOpsUpdatedOK Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "saveSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If

        'INSERT NEW OPERATOR JUST LOGGED IN - INTO Operators Online table:
        IsOpsOnlineUpdatedOK = InsertIntoOperatorsTable(connString, strUserID, strUsername,
                                       strEmpNo, strFirstname, strLastname,
                                       strComputerName, strWorkstationName, strLocation,
                                       strIP4Addr, strIP6Addr, strSessionID,
                                       strAccessRights, strDateLoggedIn, strDateLoggedOut,
                                       strIsLoggedIn, strLoggedInDuration, strComments,
                                        MyMessage)


    End Function

    Function InsertIntoOperatorsTable(connString As String, ByVal strUserID As String, ByVal strUsername As String,
                                      ByVal strEmpNo As String, ByVal strFirstname As String, ByVal strLastname As String,
                                      ByVal strComputerName As String, ByVal strWorkstationName As String, ByVal strLocation As String,
                                      ByVal strIP4Addr As String, ByVal strIP6Addr As String, ByVal strSessionID As String,
                                      ByVal strAccessRights As String, ByVal strDateLoggedIn As String, ByVal strDateLoggedOut As String,
                                      ByVal strIsLoggedIn As String, ByVal strLoggedInDuration As String, ByVal strComments As String,
                                      Optional ByRef MyMessage As String = "") As Boolean
        Dim OnlineTable As String = "tblOperatorsOnline"
        Dim OpsOnlineFieldNames = ""
        Dim OpsOnlineFieldValues = ""
        Dim OpsOnlineCriteria = ""
        Dim OpsOnlineExcludeFields = ""
        Dim IsOpsOnlineUpdatedOK As Boolean = False

        InsertIntoOperatorsTable = False
        OpsOnlineFieldNames = ""
        OpsOnlineCriteria = ""
        OpsOnlineExcludeFields = "ID"
        OpsOnlineFieldValues = strUserID
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strUsername
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strEmpNo
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strFirstname
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strLastname
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strComputerName
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strWorkstationName
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strLocation
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strIP4Addr
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strIP6Addr
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strSessionID
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strAccessRights
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strDateLoggedIn
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strDateLoggedOut
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strIsLoggedIn
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strLoggedInDuration
        OpsOnlineFieldValues = OpsOnlineFieldValues & "," & strComments

        IsOpsOnlineUpdatedOK = InsertUpdateRecords_Using_Parameters(connString, False, "", OnlineTable, OpsOnlineFieldNames, OpsOnlineFieldValues,
                                                                    OpsOnlineCriteria, OpsOnlineExcludeFields, MyMessage)
        'IsOpsOnlineUpdatedOK = InsertUpdateMyRecord(False, connString, OnlineTable, OpsOnlineFieldNames, OpsOnlineFieldValues, MyMessage, OpsOnlineCriteria, OpsOnlineExcludeFields)
        If Not IsOpsOnlineUpdatedOK Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "saveSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        Else
            InsertIntoOperatorsTable = True
        End If
    End Function

    Sub UpdateMySession(ByVal connString As String, ByVal UserID As String, ByVal Username As String,
                    ByVal ComputerName As String, ByVal IPv4Addr As String, ByVal IPv6Addr As String,
                    ByVal SessionID As String, ByVal AccessRights As String, ByVal LogOutDate As String,
                    ByVal IsLoggedIn As String, ByVal LoggedInDuration As String, ByVal Criteria As String, ByRef MyMessage As String,
                      Optional ByRef MyOpsMessage As String = "")


        Dim ConnMessage As String = ""
        Dim DBTable As String = ""
        Dim FieldValues As String = ""
        Dim FieldNames As String = ""
        Dim ExcludeFields As String = ""
        Dim strComputerName As String = ""
        Dim strIP4Addr As String = ""
        Dim strIP6Addr As String = ""
        Dim strGUID As String = ""
        Dim strSessionID As String = ""
        Dim strAccessRights As String = ""
        Dim strUsername As String = ""
        Dim strDateLoggedIn As String = ""
        Dim strDateLoggedOut As String = ""
        Dim strUserID As String = ""
        Dim strIsLoggedIn As String = ""
        Dim strLoggedInDuration As String = ""
        Dim OpsDBTable As String = ""
        Dim OpsFieldnames As String = ""
        Dim OpsExcludeFields As String = ""
        Dim OpsFieldValues As String = ""
        Dim OpsCriteria As String = ""
        Dim dtMinutesDiff As Long = 0
        Dim dtDateLoggedIn As DateTime
        Dim dtDateLoggedOut As DateTime
        Dim strHours As String = ""
        Dim strMins As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = ""
        Dim ReturnValue As String = ""
        Dim OpSearchCriteria As String = ""
        Dim OpMessage As String = ""
        Dim OnlineTable As String = ""
        Dim strConnectionString As String = ""
        Dim OpDeleteCriteria As String = ""
        Dim IsUpdatedOK As Boolean = False
        Dim IsOpsUpdatedOK As Boolean = False
        Dim FoundSessionID As Boolean = False
        Dim DeletedOK As Boolean = False
        Dim AllFieldNames As String() = Nothing
        Dim AllFieldValues As Object() = Nothing
        Dim DecMinutes As Decimal = 0.0
        Dim DecHours As Decimal = 0.0
        Dim dblTotalTime As Double
        Dim intHours As Integer
        Dim SPAN As TimeSpan

        DBTable = "tblSessions"
        OpsDBTable = "tblOperators"
        OnlineTable = "tblOperatorsOnline"
        strUserID = UserID
        strUsername = Username
        strComputerName = ComputerName
        strIP4Addr = IPv4Addr
        strIP6Addr = IPv6Addr
        strSessionID = SessionID
        strAccessRights = AccessRights
        strIsLoggedIn = IsLoggedIn
        strDateLoggedOut = LogOutDate

        OpDeleteCriteria = "SessionID = " & Chr(39) & strSessionID & Chr(39)
        If Len(connString) > 0 Then
            strConnectionString = connString
        Else
            strConnectionString = frmMainGIForm.myConnString
        End If
        strHours = "0"
        strMins = "0"
        strLoggedInDuration = "0 mins"
        strDateLoggedOut = CStr(Now())
        SearchField = "SessionID"
        SearchText = strSessionID
        ReturnField = "LogInDateTime"
        'Find logindate from tblSessions
        FoundSessionID = Find_myQuery(strConnectionString, DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, AllFieldValues, AllFieldNames, OpSearchCriteria, "", False, OpMessage)
        If Len(OpMessage) > 0 Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, OpMessage, "UpdateSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If
        If FoundSessionID Then
            strDateLoggedIn = ReturnValue
            dtDateLoggedIn = CDate(strDateLoggedIn)
            dtDateLoggedOut = CDate(strDateLoggedOut)
            SPAN = dtDateLoggedOut.Subtract(dtDateLoggedIn)
            dtMinutesDiff = DateDiff(DateInterval.Minute, dtDateLoggedIn, dtDateLoggedOut)
            DecHours = Math.Abs(dtMinutesDiff) / 60
            DecMinutes = Math.Abs(dtMinutesDiff) Mod 60
            If DecHours < 1 Then
                strHours = "0"
            Else
                Dim theRounded = Math.Sign(DecHours) * Math.Floor(Math.Abs(DecHours) * 100) / 100.0
                intHours = CInt(theRounded)
                strHours = CStr(intHours)
                'hmmm hours still ending up as 1.34 and 5.6 etc.
            End If
            strMins = CStr(DecMinutes)

            strLoggedInDuration = SPAN.Hours & "h " & SPAN.Minutes & "m " & SPAN.Seconds & "s "

        End If

        'OpsCriteria = "Username = " & Chr(39) & strUsername & Chr(39)
        OpsCriteria = "SessionID = " & Chr(39) & strSessionID & Chr(39)
        FieldNames = "LogOffDateTime,LogInDateTime,IsLoggedIn,LoggedInDuration"
        OpsFieldnames = "IsLoggedIn,SessionID"
        ExcludeFields = ""
        OpsExcludeFields = "ID,UserID,Username,Firstname,Lastname,EmpNo,DateCreated,Comments,AccessRights"
        FieldValues = dtDateLoggedOut & "," & dtDateLoggedIn & "," & strIsLoggedIn & "," & strLoggedInDuration

        IsUpdatedOK = InsertUpdateRecords_Using_Parameters(strConnectionString, True, "", DBTable, FieldNames, FieldValues, OpsCriteria,
                                                           ExcludeFields, MyMessage, False, ",")
        If Not IsUpdatedOK Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "UpdateSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If
        'UPDATE tblOperators with Logged In info:
        OpsFieldValues = strIsLoggedIn & "," & strSessionID
        IsOpsUpdatedOK = InsertUpdateRecords_Using_Parameters(strConnectionString, True, "", OpsDBTable, OpsFieldnames, OpsFieldValues, OpsCriteria,
                                                              OpsExcludeFields, MyOpsMessage, False, ",")
        If Not IsOpsUpdatedOK Then
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyOpsMessage, "UpdateSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If

        'REMOVE THE OPERATOR FROM THE OperatorsOnline table:
        DeletedOK = DeleteMyRecord(OnlineTable, strConnectionString, OpDeleteCriteria, MyMessage)
        If DeletedOK = True Then
            'Me.txtMessages.AppendText(vbCrLf & " *Entry moved to History: " & strAssetName)
            frmMainGIForm.logger.LogMessage("GI_Error_v1_1.log", Application.StartupPath, "User Logged Out: " & frmMainGIForm.myUsername, frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address())
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, "User Logged Out: " & frmMainGIForm.myUsername, "UpdateSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged OUT:" & frmMainGIForm.myUsername)
        Else
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, "Error Logging Out: " & MyMessage, "UpdateSession()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged IN:" & frmMainGIForm.myUsername)
        End If


    End Sub

    Function Check_myLoggedIN(Optional ByRef Message As String = "") As Boolean
        Dim Processed As Boolean = False

        If mylogged_in = True Then
            Processed = True
            Message = "Successfully Logged In"
        Else
            Message = "Cannot show form until successfully logged in."
        End If


        Check_myLoggedIN = Processed
    End Function

    Function PopulateUsersTable(ByVal connStr As String, ByVal dbTable As String, ByVal EmpNo As String, ByVal UserBarCode As String) As Boolean
        'Intended use ? - i think has been superceeded by another function ? D.G. 30-08-2017 at 20:46
    End Function

    Function GetMyNumFields(connStr As String, Sqlstr As String,
                          Optional ByRef FieldList As String = "", Optional ByRef FieldTypes As String = "",
                          Optional ByRef ColumnWidths As String = "", Optional ByRef Message As String = "") As Long
        Dim Numfields As Long
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim myReader As MySqlDataReader
        Dim dt As DataTable
        Dim fldIDX As Integer
        Dim myRowData As DataRow
        Dim myColData As DataColumn

        GetMyNumFields = 0
        FieldList = ""
        FieldTypes = ""
        Numfields = 0
        If Len(Sqlstr) = 0 Then
            Message = "Error in GetMyNumFields: query string passed is empty."
            frmMainGIForm.logger.LogError("Error.log", Application.StartupPath, Message, "GetMyNumFields()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            GetMyNumFields = 0
            Exit Function
        End If
        Try
            con = New MySqlConnection(connStr)
            cmd = con.CreateCommand()
            cmd.CommandText = Sqlstr
            con.Open()
            myReader = cmd.ExecuteReader
            Numfields = myReader.FieldCount
            For fldIDX = 0 To Numfields - 1
                If fldIDX = 0 Then
                    FieldList = myReader.GetName(fldIDX)
                    FieldTypes = myReader.GetName(fldIDX).GetType.ToString
                    ColumnWidths = CStr(myReader.GetSchemaTable().Columns(fldIDX).MaxLength)
                Else
                    FieldList = FieldList & "," & myReader.GetName(fldIDX)
                    FieldTypes = FieldTypes & "," & myReader.GetName(fldIDX).GetType.ToString
                    ColumnWidths = ColumnWidths & "," & CStr(myReader.GetSchemaTable().Columns(fldIDX).MaxLength)
                End If
            Next
            'dt = myReader.GetSchemaTable()
            myReader.Close()
            con.Close()

        Catch ex As Exception
            Message = "Exception Error In GetMyNumFields: " & ex.Message.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetMyNumFields()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        GetMyNumFields = Numfields
    End Function

    Function IsFieldExcluded(Fieldname As String, ExcludedList As String) As Boolean
        Dim IDX As Integer
        Dim ExcludedListArr As String()

        ExcludedListArr = Split(ExcludedList, ",")
        IsFieldExcluded = False
        For IDX = 0 To UBound(ExcludedListArr) 'DOES THIS LIST ALL THE FIELDS IN THE PASSED LIST ?????
            If UCase(Fieldname) = UCase(ExcludedList(IDX)) Then
                IsFieldExcluded = True
            End If
        Next

        Return IsFieldExcluded

    End Function

    Function GetMyFields(DBTable As String, ByVal myconnString As String, ByRef Message As String, Optional ExcludeFields As String = "",
                         Optional ByRef MyTypes As String = "", Optional ByVal StartTypeIDX As Long = 0, Optional ByRef Dic_Types As Object = Nothing) As String
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim dt As DataTable
        Dim DataSource As String
        Dim strSQL As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Provider As String = ""
        Dim colIDX As Integer = 0
        Dim Fieldnames As String = ""
        Dim dr1 As MySqlDataReader
        'Dim FieldType As Type
        Dim FieldType As String
        Dim TypeIDX As Long
        Dim Fieldname As String

        GetMyFields = ""
        NumBlankFields = 0
        NumFields = 0
        Message = ""
        NumRows = 0
        strSQL = ""
        DataSource = ""
        TypeIDX = StartTypeIDX

        Dic_Types = CreateObject("Scripting.Dictionary")
        Dic_Types.comparemode = vbTextCompare

        If Len(DBTable) = 0 Then
            'MsgBox("Error in Insert Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            Message = "Error in Insert Record: Database Table not provided"
            Exit Function
        End If
        Try
            con = New MySqlConnection(myconnString)

            con.Open()
            strSQL = "SELECT * FROM " & DBTable
            cmd = New MySqlCommand(strSQL, con)
            dr1 = cmd.ExecuteReader()

            NumFields = dr1.FieldCount - 1
            If Len(ExcludeFields) = 0 Then
                While colIDX <= NumFields
                    Fieldname = dr1.GetName(colIDX)
                    FieldType = dr1.GetDataTypeName(colIDX)
                    If Not Dic_Types.exists(Fieldname) Then
                        Dic_Types(Fieldname) = FieldType
                    End If
                    If colIDX = 0 Then
                        Fieldnames = Fieldname
                        MyTypes = FieldType.ToString
                    Else
                        Fieldnames = Fieldnames & "," & Fieldname
                        MyTypes = MyTypes & "," & FieldType.ToString
                    End If
                    colIDX = colIDX + 1
                End While
            Else
                While colIDX <= NumFields
                    If IsFieldExcluded(dr1.GetName(colIDX), ExcludeFields) Then
                        'ignore
                    Else
                        If colIDX = 0 Then
                            Fieldnames = dr1.GetName(colIDX)
                        Else
                            Fieldnames = Fieldnames & "," & dr1.GetName(colIDX)
                        End If
                    End If
                    colIDX = colIDX + 1
                End While
            End If

            con.Close()
            dr1.Close()
            GetMyFields = Fieldnames

        Catch ex As Exception
            Message = "Exception Error In GetMyFields: " & ex.Message.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetMyFields()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

    End Function

    Function Get_Field_Index_From_Fieldname(ByVal DBTable As String, ByVal ConnString As String, ByVal Fieldname As String, ByRef ErrorMessage As String,
                                            Optional ByVal PassbackNegativeIfNotFound As Boolean = False) As Integer
        Dim Fieldnames As String = ""
        Dim FieldArr As String()
        Dim FieldIDX As Integer
        Dim NumFields As Integer = 0
        Dim ThisField As String = ""

        If PassbackNegativeIfNotFound = True Then
            Get_Field_Index_From_Fieldname = -1
        Else
            Get_Field_Index_From_Fieldname = 0
        End If

        ErrorMessage = ""
        If Len(DBTable) = 0 Then
            ErrorMessage = "Error in Get_Field_Index_From_Fieldname: Table name not specified"
            Exit Function
        End If
        If Len(Fieldname) = 0 Then
            ErrorMessage = "Error in Get_Field_Index_From_Fieldname: No Fieldname specified"
            Exit Function
        End If
        Try
            Fieldnames = GetMyFields(DBTable, ConnString, ErrorMessage)
            FieldArr = strToStringArray(Fieldnames, ",", 0)
            NumFields = UBound(FieldArr)
            For FieldIDX = 0 To NumFields
                ThisField = FieldArr(FieldIDX)
                If UCase(Fieldname) = UCase(ThisField) Then
                    Get_Field_Index_From_Fieldname = FieldIDX
                    Exit Function
                End If
            Next
        Catch ex As Exception
            ErrorMessage = "Error in Export " & ex.Message
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrorMessage, "Get_Field_Index_From_Fieldname()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

    End Function

    Function GetPreferences(ByRef DefaultExportDir As String, ByRef UserID As String, ByRef Username As String, ByRef EmpNo As String, ByRef DatabaseType As String,
                            ByRef MainFormBackColor As String, ByRef DateLastChanged As String, Optional ByRef Messages As String = "",
                            Optional ByRef WaitTime As String = "30", Optional ByRef AllowImmediateAlerts As String = "1") As Boolean
        Dim strUsername As String = ""
        Dim strMainBackColor As String = ""
        Dim strDatabaseType As String = ""
        Dim strDateLastChanged As String = ""
        Dim strEmpNo As String = ""
        Dim strDefaultExportDirectory As String = ""
        Dim strWaitTime As String = ""
        Dim strAllowImmediateAlerts As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = "ID"
        Dim Criteria As String = ""
        Dim InsertMessage As String = ""
        Dim FieldValues As String = ""
        Dim DBTable As String = ""
        Dim ReturnValue As String = ""
        Dim Errors As String = ""
        Dim UserExists As Boolean = False
        Dim InsertOK As Boolean = False
        Dim AllFieldNames As String()
        Dim AllPreferenceValues As Object()
        Dim SortField As String = ""
        Dim FindUser As Boolean = False

        GetPreferences = False
        DBTable = "tblPreferences"
        SearchField = "userID" 'TYPE = INTEGER is from tblUser ID field.
        If Len(UserID) = 0 Then
            UserID = frmMainGIForm.myUserID 'ID in string form
        End If
        SearchText = UserID
        ReturnField = "Username"
        'Criteria = "userID = '" & SearchText & "'"

        ReDim AllFieldNames(1)
        AllPreferenceValues = Nothing
        SortField = "userID"
        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            FindUser = Find_me(frmMainGIForm.DBName, "", DBTable, SearchField, SearchText, "INTEGER", ReturnField,
                               ReturnValue, AllPreferenceValues, AllFieldNames, Criteria, "", SortField, False)
        Else
            FindUser = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "INTEGER", ReturnField,
                                    ReturnValue, AllPreferenceValues, AllFieldNames, Criteria, SortField, False, Messages)
        End If
        If FindUser = True Then
            strUsername = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "Username")
            strEmpNo = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "EmpNo")
            strMainBackColor = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "MainFormBackColor")
            strDatabaseType = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "DatabaseType")
            strDefaultExportDirectory = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "DefaultExportDirectory")
            strDateLastChanged = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "DateLastChanged")
            strWaitTime = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "WaitTime")
            strAllowImmediateAlerts = GetValuebyFieldname(AllPreferenceValues, AllFieldNames, "AllowImmediateAlerts")
            GetPreferences = True

        Else
            strUsername = ""
            strMainBackColor = ""
            strEmpNo = ""
            strDatabaseType = ""
            strDefaultExportDirectory = ""
            strDateLastChanged = ""
            strWaitTime = ""
            strAllowImmediateAlerts = ""
        End If

        DefaultExportDir = strDefaultExportDirectory
        DatabaseType = strDatabaseType
        MainFormBackColor = strMainBackColor
        Username = strUsername
        EmpNo = strEmpNo
        DateLastChanged = strDateLastChanged
        WaitTime = strWaitTime
        AllowImmediateAlerts = strAllowImmediateAlerts

    End Function

    Function SaveToPreferencesTable(ByVal DefaultExportDir As String, ByVal UserID As String, ByVal Username As String,
                                    ByVal EmpNo As String, ByVal DatabaseType As String, ByVal MainFormBackColor As String, Optional ByRef Messages As String = "",
                                    Optional WaitTime As String = "", Optional AllowImmediateAlerts As String = "") As Boolean
        Dim strUserID As String = ""
        Dim strUsername As String = ""
        Dim strMainBackColor As String = ""
        Dim strDatabaseType As String = ""
        Dim strDateLastChanged As String = ""
        Dim strDefaultExportDirectory As String = ""
        Dim strEmpNo As String = ""
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = "ID"
        Dim Criteria As String = ""
        Dim InsertMessage As String = ""
        Dim FieldValues As String = ""
        Dim DBTable As String = ""
        Dim ReturnValue As String = ""
        Dim Errors As String = ""
        Dim UserExists As Boolean = False
        Dim InsertOK As Boolean = False
        Dim strWaitTime As String = ""
        Dim StrAllowImmediateAlerts As String = ""

        'check if username already exists
        SaveToPreferencesTable = False
        SearchField = "UserID"
        SearchText = UserID
        ReturnField = "ID"
        Criteria = "userID = '" & SearchText & "'"
        strUserID = UserID
        strDefaultExportDirectory = DefaultExportDir
        strDatabaseType = DatabaseType
        strMainBackColor = MainFormBackColor
        strUsername = Username
        strWaitTime = WaitTime
        StrAllowImmediateAlerts = AllowImmediateAlerts
        DBTable = "tblPreferences"
        strDateLastChanged = CStr(Now())
        FieldValues = Chr(34) & strUserID & Chr(34) _
                & "," & Chr(34) & strUsername & Chr(34) _
                & "," & Chr(34) & strEmpNo & Chr(34) _
                & "," & Chr(34) & strMainBackColor & Chr(34) _
                & "," & Chr(34) & strDatabaseType & Chr(34) _
                & "," & Chr(34) & strDateLastChanged & Chr(34) _
                & "," & Chr(34) & strDefaultExportDirectory & Chr(34) _
                & "," & Chr(34) & strWaitTime & Chr(34) _
                & "," & Chr(34) & StrAllowImmediateAlerts & Chr(34)

        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            UserExists = Find_me(frmMainGIForm.DBName, "", DBTable, SearchField, SearchText, "INTEGER", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "")
            If Not UserExists Then
                InsertOK = InsertUpdateRecord(False, frmMainGIForm.DBName, "", DBTable, "", FieldValues, Criteria, "ID")
                If Not InsertOK Then
                    Messages = "Insert Error"
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "SaveToPreferencesTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
            Else
                InsertOK = InsertUpdateRecord(True, frmMainGIForm.DBName, "", DBTable, "", FieldValues, Criteria, "ID")
                If Not InsertOK Then
                    Messages = "Update Error"
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "SaveToPreferencesTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
            End If
        Else
            If Module_DanG_MySQL_Tools.mylogged_in Then
                UserExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "INTEGER", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", False, InsertMessage)
                If Len(InsertMessage) > 0 Then
                    Messages = InsertMessage
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "SaveToPreferencesTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
                If Not UserExists Then
                    InsertOK = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, DBTable, "", FieldValues, InsertMessage, Criteria, "ID")
                    If Len(InsertMessage) > 0 Then
                        Messages = InsertMessage
                        frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "SaveToPreferencesTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                        Exit Function
                    End If
                Else
                    InsertOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, DBTable, "", FieldValues, InsertMessage, Criteria, "ID")
                    If Len(InsertMessage) > 0 Then
                        Messages = InsertMessage
                        frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "SaveToPreferencesTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                        Exit Function
                    End If
                End If
            Else
                'get user to log in:
                Module_DanG_MySQL_Tools.Check_myLoggedIN()

            End If

        End If
        Application.DoEvents()
        SaveToPreferencesTable = True
    End Function

    Function ExportDGVToCSVFile(ByRef dgv As DataGridView, ByVal CSVFilename As String, ByVal Optional Extension As String = ".csv", Optional ByRef anyMessages As String = "") As Boolean

        Dim headers = (From header As DataGridViewColumn In dgv.Columns.Cast(Of DataGridViewColumn)()
                       Select header.HeaderText).ToArray
        Dim dlgSaveCSV As New SaveFileDialog()
        Dim CSVPath As String = ""
        Dim strUserID As String = ""
        Dim strUsername As String = ""
        Dim strDatabaseType As String = ""
        Dim strMainFormBackColor As String = ""
        Dim strDefaultExportPath As String = ""
        Dim strDateLastChanged As String = ""
        Dim ErrMessages As String = ""
        Dim strFilename As String = ""
        Dim strEmpNo As String = ""

        Dim rows = From row As DataGridViewRow In dgv.Rows.Cast(Of DataGridViewRow)()
                   Where Not row.IsNewRow
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        'Load up current preferences first: - if strUserID is passed blank - then the frmMainGIForm.myUserID is used (Operator signed in)
        ' - The rest of the fields will be populated - IF there is a record for the current User/OPerator in the preferences table:
        ' - Do we save the path and filename - or just the path - because the new filename is passed into this function ???
        ' - Why are we getting an Access Denied message when trying to save to ANY specific directory ?

        'HOW DO WE EXCLUDE Certain Fields passed - that we do NOT want sent to the CSV file output ?????????????????????

        GetPreferences(strDefaultExportPath, strUserID, strUsername, strEmpNo, strDatabaseType, strMainFormBackColor, strDateLastChanged, ErrMessages)
        dlgSaveCSV.Filter = "CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
        dlgSaveCSV.Title = "Save data to CSV File"
        dlgSaveCSV.FilterIndex = 1
        dlgSaveCSV.RestoreDirectory = False
        If Len(strDefaultExportPath) > 0 Then
            dlgSaveCSV.InitialDirectory = strDefaultExportPath
            dlgSaveCSV.FileName = CSVFilename
        Else
            dlgSaveCSV.InitialDirectory = frmMainGIForm.ProgramData
            dlgSaveCSV.FileName = CSVFilename
        End If
        If Len(strDatabaseType) = 0 Then
            strDatabaseType = "MYSQL"
        End If
        If dlgSaveCSV.ShowDialog() = DialogResult.OK Then
            CSVPath = IO.Path.GetDirectoryName(dlgSaveCSV.FileName)
            CSVPath = CSVPath & "\" & CSVFilename & Extension
            strFilename = Replace(CSVPath, "\", "\\")
            'strFilename = strFilename & Extension
            'Save path to preferences table:
            strUsername = frmMainGIForm.myUsername
            strEmpNo = frmMainGIForm.myEmpNo
            SaveToPreferencesTable(strFilename, strUserID, strUsername, strEmpNo, strDatabaseType, strMainFormBackColor, anyMessages)
            frmMainGIForm.logger.LogMessage("GI_Error_v1_1.log", Application.StartupPath, "Preference Path:" & CSVPath, frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        Else
            CSVPath = CSVFilename & Extension

        End If
        'MsgBox("CSV Save Path= " & CSVPath & ", CSVFilename= " & strFilename)
        ExportDGVToCSVFile = False
        Try
            'Using sw As New IO.StreamWriter(CSVFilename & Extension)
            Using sw As New IO.StreamWriter(strFilename)
                sw.WriteLine(String.Join(",", headers))
                For Each r In rows
                    sw.WriteLine(String.Join(",", r))
                Next
            End Using
            'Process.Start(CSVPath)
            Process.Start(strFilename)
        Catch ex As Exception
            anyMessages = "Error in Export " & ex.Message
            'Getting a file is denied error here ???? !!!! - the filename was wrong in the StreamWriter() parameter.
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, anyMessages, "ExportDGVToCSVFile()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        ExportDGVToCSVFile = True
    End Function

    Function ExportDGVToPDFFile(ByRef dgv As DataGridView, ByVal PDFFilename As String, ByVal Optional Extension As String = ".pdf", Optional Widths() As Single = Nothing,
                                Optional PageTitle As String = "FAULT LIST REPORT", Optional ByVal BarCodeColumn As Integer = 3, Optional ByVal PlotSecondRow As Boolean = False,
                                Optional ByVal RowsPerPage As Long = 10, Optional ByVal BarCodeFontSize As Single = 15,
                                           Optional ByVal TitleFontSize As Single = 20, Optional ByVal HeadingFontSize As Single = 20,
                                Optional ByVal BodyFontSize As Single = 12, Optional ByVal Landscape As Boolean = True, Optional ByVal RowHeight As Single = 30,
                                Optional ByVal RowHeight2 As Single = 20, Optional StartArrayIndex As Long = 0, Optional ByRef anyMessages As String = "") As Boolean
        Dim headers = (From header As DataGridViewColumn In dgv.Columns.Cast(Of DataGridViewColumn)()
                       Select header.HeaderText).ToArray
        Dim dlgSavePDF As New SaveFileDialog()
        Dim PDFPath As String = ""
        Dim strUserID As String = ""
        Dim strUsername As String = ""
        Dim strDatabaseType As String = ""
        Dim strMainFormBackColor As String = ""
        Dim strDefaultExportPath As String = ""
        Dim strDateLastChanged As String = ""
        Dim ErrMessages As String = ""
        Dim strFilename As String = ""
        Dim strEmpNo As String = ""
        Dim strFields As String = ""
        Dim strWholerow As String = ""
        Dim PDFArr(,) As String
        Dim Fields As Long
        Dim GridRows As Long
        Dim myRow As Long
        Dim myCol As Long
        Dim NumCols As Integer
        Dim WidthIDX As Integer
        Dim MaxWidth As Single

        Dim rows = From row As DataGridViewRow In dgv.Rows.Cast(Of DataGridViewRow)()
                   Where Not row.IsNewRow
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        'Load up current preferences first: - if strUserID is passed blank - then the frmMainGIForm.myUserID is used (Operator signed in)
        ' - The rest of the fields will be populated - IF there is a record for the current User/OPerator in the preferences table:
        ' - Do we save the path and filename - or just the path - because the new filename is passed into this function ???
        ' - Why are we getting an Access Denied message when trying to save to ANY specific directory ?
        GetPreferences(strDefaultExportPath, strUserID, strUsername, strEmpNo, strDatabaseType, strMainFormBackColor, strDateLastChanged, ErrMessages)
        dlgSavePDF.Filter = "PDF File (*.pdf)|*.pdf|All Files (*.*)|*.*"
        dlgSavePDF.Title = "Save data to PDF File"
        dlgSavePDF.FilterIndex = 1
        dlgSavePDF.RestoreDirectory = False
        If Len(strDefaultExportPath) > 0 Then
            dlgSavePDF.InitialDirectory = strDefaultExportPath
            dlgSavePDF.FileName = PDFFilename
        Else
            dlgSavePDF.InitialDirectory = frmMainGIForm.ProgramData
            dlgSavePDF.FileName = PDFFilename
        End If
        If Len(strDatabaseType) = 0 Then
            strDatabaseType = "MYSQL"
        End If
        If dlgSavePDF.ShowDialog() = DialogResult.OK Then
            PDFPath = IO.Path.GetDirectoryName(dlgSavePDF.FileName)
            'save the default export path here before it gets the filename added:
            strDefaultExportPath = PDFPath
            PDFPath = PDFPath & "\" & PDFFilename & Extension
            strFilename = Replace(PDFPath, "\", "\\")
            strDefaultExportPath = Replace(strDefaultExportPath, "\", "\\")
            'strFilename = strFilename & Extension
            'Save path to preferences table:
            strUsername = frmMainGIForm.myUsername
            strEmpNo = frmMainGIForm.myEmpNo
            SaveToPreferencesTable(strDefaultExportPath, strUserID, strUsername, strEmpNo, strDatabaseType, strMainFormBackColor, anyMessages)
            frmMainGIForm.logger.LogMessage("GI_Error_v1_1.log", Application.StartupPath, "Preference Path:" & PDFPath, frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        Else
            PDFPath = PDFFilename & Extension

        End If
        'MsgBox("CSV Save Path= " & CSVPath & ", CSVFilename= " & strFilename)
        ExportDGVToPDFFile = False
        ReDim PDFArr(headers.Count, rows.Count)
        Try
            'Using sw As New IO.StreamWriter(CSVFilename & Extension)
            'Now we really need to call a general purpose Export to PDF file procedure which takes parameters like number per page and column widths
            'Number of columns in each row .

            For myRow = 0 To rows.Count - 1
                For myCol = 0 To headers.Count - 1
                    PDFArr(myCol, myRow) = rows(myRow)(myCol)
                Next
            Next
            If IsNothing(Widths) Then
                'CALCULATE WIDTHS. Need num columns first.
                NumCols = headers.Count
                ReDim Widths(NumCols - 1)
                If Landscape = True Then
                    MaxWidth = 32
                Else
                    MaxWidth = 20
                End If
                For WidthIDX = 0 To NumCols - 1
                    Widths(WidthIDX) = CSng(MaxWidth \ NumCols)
                Next
            End If
            If Len(strFilename) > 0 Then
                'CreatePDFDoc_Fault_Report_List(strFilename, PageTitle, headers, PDFArr, BarCodeColumn, Widths, PlotSecondRow, RowsPerPage, BarCodeFontSize,
                'TitleFontSize, HeadingFontSize, BodyFontSize, Landscape, RowHeight, RowHeight2, StartArrayIndex, anyMessages)
            Else
                anyMessages = "PDF SAVE was Cancelled"
            End If

        Catch ex As Exception
            anyMessages = "Error in Export " & ex.Message
            'Getting a file is denied error here ???? !!!! - the filename was wrong in the StreamWriter() parameter.
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, anyMessages, "ExportDGVToPDFFile()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        ExportDGVToPDFFile = True
    End Function

    Function ArrayToCSVFile(ByRef HeaderArr As String(), RowArr As String(), CSVFilename As String, ByVal Optional Extension As String = ".csv",
                                Optional ByRef Message As String = "")
        Try
            Using sw As New IO.StreamWriter(CSVFilename & Extension)
                sw.WriteLine(String.Join(",", HeaderArr))
                For Each r In RowArr
                    sw.WriteLine(String.Join(",", r))
                Next
            End Using
            Process.Start(CSVFilename & Extension)
        Catch ex As Exception
            Message = "Error in Export " & ex.Message
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "ArrayToCSVFile()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

    End Function

    Function Get_Total_Rows_From_CSVFile(CSVFilename As String, Optional ByRef TotalFields As Long = 0) As Long
        Dim rownum As Long
        Dim TotalRows As Long
        Dim wholeRows As String()

        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(CSVFilename)
            csvReader.TextFieldType = FileIO.FieldType.Delimited
            csvReader.SetDelimiters(",")
            TotalRows = 0
            rownum = 0
            While Not csvReader.EndOfData
                Try
                    wholeRows = csvReader.ReadFields()

                    TotalFields = wholeRows.Length
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is NOT valid and will be skipped.")
                End Try
                rownum = rownum + 1
            End While

        End Using
        TotalRows = rownum
        Get_Total_Rows_From_CSVFile = TotalRows
    End Function

    Function CSVFileToArray(ByRef NewstrArray As String(,), ByVal CSVFilename As String, Optional ByRef TotalFields As Long = 0,
                                Optional ByRef Message As String = "") As Long
        Dim wholeRows As String()
        Dim numRowsRead As Long
        Dim colnum As Long
        Dim RowNum As Long
        Dim currentFields As String
        Dim Percentage As Long
        Dim TotalRows As Long
        Dim RowMessage As String
        Dim messages As String

        ReDim NewstrArray(1, 1)
        CSVFileToArray = 0
        numRowsRead = 0
        colnum = 0
        RowNum = 0
        messages = ""
        ReDim wholeRows(1)
        wholeRows(0) = ""
        'ProgressBarControl.Maximum = 100
        'ProgressBarControl.Value = 0
        'ProgressLabel.Text = "0%"
        'Import from CSV File TEXT field and check validity:
        Try
            Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(CSVFilename)
                csvReader.TextFieldType = FileIO.FieldType.Delimited
                csvReader.SetDelimiters(",")
                TotalRows = Get_Total_Rows_From_CSVFile(CSVFilename, TotalFields)
                ReDim NewstrArray(TotalFields, TotalRows)
                'frmProgressBar.Show()
                While Not csvReader.EndOfData
                    Try
                        wholeRows = csvReader.ReadFields()
                        colnum = 0
                        For Each currentFields In wholeRows
                            'Me.rtbOutput.Text = currentFields
                            'MsgBox(CStr(RowNum) & ": Field " & CStr(colnum) & " = " & currentFields)
                            '
                            '****************************************************
                            '
                            'AUTHOR : DANIEL GOSS . COPYRIGHT 24-FEB-2017. V1.2
                            'MODIFIED: DANIEL GOSS . COPYRIGHT 01-SEP-2017. v1.3
                            ' - added extra parameter - Messages
                            'MODIFIED: DANIEL GOSS . COPYRIGHT 03-SEP-2017. v1.4 now.
                            ' - removed extra parameter - Messages
                            ' - removed reference to GetFieldValue_From_Index below.
                            ' - added error logging and message logging 12-09-2017 01:50
                            '*****************************************************

                            If TotalRows > 0 Then
                                Percentage = CLng((RowNum / TotalRows) * 100)
                            Else
                                Percentage = 0
                            End If
                            RowMessage = "Read " & CStr(RowNum) & " ROW / " & CStr(TotalRows) & " ROWS"
                            'RichTextBoxMessageControl.Text = "Reading from CSV file ..."
                            'RichTextBoxMessageControl.Text = vbCrLf & RowMessage
                            'ProgressBarControl.Value = Percentage
                            'ProgressLabel.Text = CStr(ProgressBarControl.Value) & "%"
                            NewstrArray(colnum, RowNum) = currentFields

                            Application.DoEvents()

                            colnum = colnum + 1
                        Next
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        If Len(messages) > 0 Then
                            messages = messages & ","
                        End If
                        messages = messages & " Line " & ex.Message & "is NOT valid and will be skipped."
                        Message = messages
                        frmMainGIForm.logger.LogMessage("GI_Error_v1_1.log", Application.StartupPath, Message, frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    End Try
                    RowNum = RowNum + 1
                End While
            End Using
        Catch ex As Exception
            Message = "Error In CSVFileToArray: " & ex.Message.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetFieldValue_From_Index()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        numRowsRead = RowNum
        'currentRow is an array containing all the data now.
        'Me.rtbOutput.Text = ""
        'frmProgressBar.Hide()
        CSVFileToArray = numRowsRead
    End Function

    Function GetFieldValue_From_Index(AllRows As Object, RowIDX As Long, FieldIDX As Long, Optional ByRef Message As String = "") As String
        Dim ReturnValue As String
        Dim colIDX As Long
        Dim ArrayValues As String

        ReturnValue = ""
        Try
            For colIDX = 0 To UBound(AllRows, 1)
                ArrayValues = AllRows(colIDX, RowIDX)
                If colIDX = FieldIDX Then
                    ReturnValue = ArrayValues
                    Exit For
                End If
                'colIDX = colIDX + 1
            Next
        Catch ex As Exception
            Message = "Error In GetFieldValue_From_Index: " & ex.Message.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetFieldValue_From_Index()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

        GetFieldValue_From_Index = ReturnValue
    End Function

    Function GetFieldValue_From_Fieldname(ByVal AllRows As Object, ByVal RowIDX As Long, ByVal FieldName As String, Optional ByRef FieldPos As Long = 0) As String
        Dim ReturnValue As String
        Dim colIDX As Long
        Dim ArrayValues As String

        'AllRows(Fields,0) - row 0 contains all the fieldnames
        GetFieldValue_From_Fieldname = ""
        ReturnValue = ""
        If IsNothing(AllRows) Then
            GetFieldValue_From_Fieldname = ReturnValue
            Exit Function
        End If
        For colIDX = 0 To UBound(AllRows, 1)
            ArrayValues = AllRows(colIDX, RowIDX)
            If UCase(FieldName) = UCase(AllRows(colIDX, 0)) Then
                ReturnValue = ArrayValues
                FieldPos = colIDX
                Exit For
            End If
            'colIDX = colIDX + 1
        Next
        GetFieldValue_From_Fieldname = ReturnValue
    End Function

    Function MySQLToArray(ByVal myConnString As String, ByVal SQLstr As String, ByRef Message As String, Optional ScriptKeyFieldNumber As Long = -1, Optional ByRef dic_TAble As Object = Nothing) As Object
        Dim NumRows As Long
        Dim NumFields As Long
        Dim RowNumber As Long
        Dim FieldNum As Integer
        Dim myArr(,) As String = {{0}, {"NONE"}}
        Dim myconn As MySqlConnection
        Dim myda As MySqlDataAdapter
        Dim myds As DataSet
        Dim mycmd As MySqlCommand
        Dim Percent As Single = 0.0F
        Dim KeyField As String = ""
        Dim KeyFieldValue As String
        Dim FieldList As String
        Dim FieldTypes As String
        Dim ColWidths As String
        Dim FieldArr() As String
        Dim FieldTypeArr() As String
        Dim ColWidthArr() As String

        If ScriptKeyFieldNumber > -1 Then
            dic_TAble = New Scripting.Dictionary
            dic_TAble.removeall
        End If

        MySQLToArray = Nothing
        frmMainGIForm.pbarMain.Value = 0
        If Len(SQLstr) = 0 Then
            'MsgBox("Error in mySQLToArray: QUERY (strSQL) is blank.", vbOK, "Error in Asset-Register v")
            Message = "Error in mySQLToArray: QUERY (strSQL) is blank."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "MySQLToArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            MySQLToArray = Nothing
            Exit Function
        End If
        If Len(myConnString) = 0 Then
            'MsgBox("Error in mySQLToArray: Connection String is blank.", vbOK, "Error in Asset-Register v")
            Message = "Error in mySQLToArray: CONNECTION STRING is BLANK."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "MySQLToArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            MySQLToArray = Nothing
            Exit Function
        End If
        Try
            myconn = New MySqlConnection(myConnString)
            myconn.Open()
            mycmd = New MySqlCommand(SQLstr, myconn)
            myda = New MySqlDataAdapter(mycmd)
            myds = New DataSet()
            myda.Fill(myds, "srcTable")
            NumRows = myds.Tables("srcTable").Rows.Count
            NumFields = GetMyNumFields(myConnString, SQLstr, FieldList, FieldTypes, ColWidths)

            FieldArr = Split(FieldList, ",")
            FieldTypeArr = Split(FieldTypes, ",")
            ColWidthArr = Split(ColWidths, ",")
            ReDim myArr(NumFields, NumRows)
            For RowNumber = 0 To NumRows - 1
                'NewValue = UBound(Arr, 2)
                For FieldNum = 0 To NumFields - 1
                    If Not IsDBNull(myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)) Then
                        myArr(FieldNum, RowNumber) = myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)
                    Else
                        myArr(FieldNum, RowNumber) = ""
                    End If
                    If ScriptKeyFieldNumber > -1 Then
                        KeyFieldValue = myArr(ScriptKeyFieldNumber, RowNumber)
                        If Not dic_TAble.exists(KeyFieldValue) Then
                            dic_TAble(KeyFieldValue) = myArr(FieldNum, RowNumber)
                        Else
                            'dic_TAble(KeyFieldValue) = dic_TAble(KeyFieldValue) & "_" & myArr(FieldNum, RowNumber)
                        End If

                    End If
                    'Call UpdateProgressBar Form here: while it populates the array
                    Percent = (RowNumber / NumRows) * 100
                    frmMainGIForm.pbarMain.Value = CInt(Percent)
                    Application.DoEvents()
                Next

            Next
            myconn.Close()
            mycmd = Nothing


        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            Message = "Exception Error: " & ex.ToString & ", SQL= " & SQLstr
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "MySQLToArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        MySQLToArray = myArr

    End Function

    Function GetUniqueItemsToSingleArray(ByVal DBTable As String, ByVal myConnString As String, ByVal SQLstr As String, ByVal Fieldname As String, ByRef Message As String) As Object
        Dim NumRows As Long
        Dim NumFields As Long
        Dim RowNumber As Long
        Dim FieldNum As Integer
        Dim myArr() As String = {"NONE"}
        Dim myconn As MySqlConnection
        Dim myda As MySqlDataAdapter
        Dim myds As DataSet
        Dim mycmd As MySqlCommand
        Dim GroupName As String = ""
        Dim GroupNameList As New Dictionary(Of String, String)
        Dim ErrMessage As String = ""
        Dim GroupElements As Integer

        GetUniqueItemsToSingleArray = Nothing

        'Do we need to specify WHICH field we want to save to the array ????
        'If just a single field - then do we just need a single array and not a 2-dim array ??
        'OR IF one fieldname is passed - return SINGLE array.
        ' - if more than one fieldname passed from same table - return as 2D array ???

        If Len(SQLstr) = 0 Then
            'MsgBox("Error in mySQLToArray: QUERY (strSQL) is blank.", vbOK, "Error in Asset-Register v")
            Message = "Error in GetUniqueItemsToSingleArray: QUERY (strSQL) is blank."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetUniqueItemsToSingleArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            GetUniqueItemsToSingleArray = Nothing
            Exit Function
        End If
        If Len(myConnString) = 0 Then
            'MsgBox("Error in mySQLToArray: Connection String is blank.", vbOK, "Error in Asset-Register v")
            Message = "Error in GetUniqueItemsToSingleArray: CONNECTION STRING is BLANK."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetUniqueItemsToSingleArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            GetUniqueItemsToSingleArray = Nothing
            Exit Function
        End If
        If Len(DBTable) = 0 Then
            Message = "Error in GetUniqueItemsToSingleArray: Table Name is BLANK."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetUniqueItemsToSingleArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            GetUniqueItemsToSingleArray = Nothing
            Exit Function
        End If
        'TEST if fieldname passed exists:
        FieldNum = Get_Field_Index_From_Fieldname(DBTable, myConnString, Fieldname, ErrMessage, True)
        If FieldNum = -1 Then
            FieldNum = 0 'SET DEFAULT to first field = ID - if passed fieldname not found.
            Message = "Error in GetUniqueItemsToSingleArray: Fieldname passed is NOT FOUND."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetUniqueItemsToSingleArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End If
        'so if fieldnum is 0 - means that fieldname cannot be found within the table passed - 0 is also the ID field. could pass back -1 instead of 0 ???
        Try
            myconn = New MySqlConnection(myConnString)
            myconn.Open()
            mycmd = New MySqlCommand(SQLstr, myconn)
            myda = New MySqlDataAdapter(mycmd)
            myds = New DataSet()
            myda.Fill(myds, "srcTable")
            NumRows = myds.Tables("srcTable").Rows.Count
            NumFields = GetMyNumFields(myConnString, SQLstr)
            ReDim myArr(1)
            GroupNameList.Clear()
            GroupElements = 0
            For RowNumber = 0 To NumRows - 1
                'NewValue = UBound(Arr, 2)
                'For FieldNum = 0 To NumFields - 1
                If Not IsDBNull(myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)) Then
                    'Check if not in the script dictionary already ?
                    GroupName = myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)
                    If GroupNameList.ContainsKey(GroupName) Then
                        'Already in dictionary.
                    Else
                        myArr(GroupElements) = myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)
                        ReDim Preserve myArr(UBound(myArr) + 1)
                        GroupNameList(GroupName) = GroupName
                        GroupElements = GroupElements + 1
                    End If
                Else
                    myArr(GroupElements) = ""
                End If
                'Call UpdateProgressBar Form here: while it populates the array

                'Next
            Next
            myconn.Close()

        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            Message = "Exception Error: " & ex.ToString & ", SQL= " & SQLstr
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetUniqueItemsToSingleArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        GetUniqueItemsToSingleArray = myArr
    End Function

    Function ArchiveData(NewTableName As String, OldTableName As String, myConnString As String, SQLstr As String, ByRef ErrMessages As String) As Object
        Dim NumRows As Long
        Dim NumFields As Long
        Dim RowNumber As Long
        Dim FieldNum As Integer
        Dim myArr(,) As String = {{0}, {"NONE"}}
        Dim myconn As MySqlConnection
        Dim myda As MySqlDataAdapter
        Dim myds As DataSet
        Dim mycmd As MySqlCommand
        Dim Dict_OldTable As Scripting.Dictionary
        Dim Dict_NewTable As Scripting.Dictionary

        ArchiveData = Nothing
        'Use the KEY as the ID into the Target Table to tell if a record has already been archived.


        If Len(SQLstr) = 0 Then
            'MsgBox("Error in mySQLToArray: QUERY (strSQL) is blank.", vbOK, "Error in Asset-Register v")
            ErrMessages = "Error in mySQLToArray: QUERY (strSQL) is blank."
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessages, "MySQLToArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            ArchiveData = Nothing
            Exit Function
        End If
        'need to test if records already exist in target table.
        Try
            myconn = New MySqlConnection(myConnString)
            myconn.Open()
            mycmd = New MySqlCommand(SQLstr, myconn)
            myda = New MySqlDataAdapter(mycmd)
            myds = New DataSet()
            myda.Fill(myds, "srcTable")
            NumRows = myds.Tables("srcTable").Rows.Count
            NumFields = GetMyNumFields(myConnString, SQLstr)
            ReDim myArr(NumFields, NumRows)
            For RowNumber = 0 To NumRows - 1
                'NewValue = UBound(Arr, 2)
                For FieldNum = 0 To NumFields - 1
                    If Not IsDBNull(myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)) Then

                        myArr(FieldNum, RowNumber) = myds.Tables("srcTable").Rows(RowNumber).Item(FieldNum)
                    Else
                        myArr(FieldNum, RowNumber) = ""
                    End If
                    'Call UpdateProgressBar Form here: while it populates the array

                Next
            Next
            myconn.Close()

        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            ErrMessages = "Exception Error: " & ex.ToString & ", SQL= " & SQLstr
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessages, "MySQLToArray()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        ArchiveData = myArr

    End Function

    Function SearchInArray() As Boolean

        SearchInArray = False

    End Function

    Sub PopulateMyDataSource(ByRef DSource As Object, ByVal connString As String, ByRef SQLStr As String,
                             Optional ByRef NumRows As Integer = 0, Optional ByRef Message As String = "")
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim IDX As Integer = 0
        Dim tblName As String = ""
        Dim PosFrom As Integer = 0
        Dim PosSpace As Integer = 0
        Dim IdxCol As DataColumn

        PosFrom = InStr(UCase(SQLStr), "FROM")
        PosSpace = InStr(PosFrom + 5, SQLStr, " ")
        If (PosSpace - PosFrom) - 5 > 0 Then
            tblName = Mid(SQLStr, PosFrom + 5, (PosSpace - PosFrom) - 5)
        Else
            tblName = "PosFrom= " & CStr(PosFrom) & ", PosSpace= " & CStr(PosSpace)
        End If

        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(SQLStr, con)
            con.Open()
            da = New MySqlDataAdapter(cmd)
            ds = New DataSet()

            da.Fill(ds, "srcTable") 'MYSQL DATE CONVERSION PROBLEM HERE !
            If IsDBNull(ds.Tables("srcTable")) Then
                'MsgBox("Problem: Empty Rows", vbOK, "Problem in Timesheet v" & frmMainGIForm.myVersion)
                Message = "Problem: Empty Rows"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource()",
                                            " TableName=" & tblName & ", " & frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                            ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                NumRows = 0
            Else
                NumRows = ds.Tables("srcTable").Rows.Count
                If NumRows > 0 Then
                    DSource = ds.Tables("srcTable") 'THIS DOES NOT REFER TO THE DATAGRIDVIEW PROPERTIES. THIS IS JUST THE DATA SOURCE.
                    If Not IsNothing(DSource) Then
                        IdxCol = ds.Tables("srcTable").Columns.Add 'yea this does add a blank column to each grid.
                        IdxCol.SetOrdinal(0) 'Yep this puts the new blank column as the first column in the data grid.
                        IdxCol.ColumnName = "#"
                        'IdxCol.width = 10
                        For RowIDX = 0 To NumRows - 1
                            'not valid:
                            ds.Tables("srcTable").Rows(RowIDX)(0) = CStr(RowIDX + 1)
                        Next
                    End If
                Else
                    Message = "NOTHING FOUND"
                End If

            End If
            con.Close()
            'frmMainGIForm.AnyEvents = False
        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            Message = Message & vbCrLf & "Exception Error in PopulateMyDataSource: " & ex.ToString & ", SQL= " & SQLStr
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource()",
                                        " TableName=" & tblName & ", " & frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try


    End Sub

    Sub PopulateMyDataSource_With_ColumnWidths(ByRef DGV As DataGridView, ByRef DSource As Object, ByVal connString As String, ByRef SQLStr As String,
                             Optional ByRef NumRows As Integer = 0, Optional ByRef Message As String = "", Optional ByVal ColumnWidths As String = "")
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim ColwidthArr As String() = Nothing
        Dim IDX As Integer = 0

        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(SQLStr, con)
            con.Open()
            da = New MySqlDataAdapter(cmd)
            ds = New DataSet()
            da.Fill(ds, "srcTable") 'MYSQL DATE CONVERSION PROBLEM HERE !
            ReDim ColwidthArr(1)
            If IsDBNull(ds.Tables("srcTable")) Then
                'MsgBox("Problem: Empty Rows", vbOK, "Problem in Asset Register v" & frmMainGIForm.myVersion)
                Message = "Problem: Empty Rows"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource_With_ColumnWidths()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                NumRows = 0
            Else
                NumRows = ds.Tables("srcTable").Rows.Count
                'ds.Tables("srcTable").Columns.Add.Caption = "Row Num"
                For RowIDX = 1 To NumRows
                    'ds.Tables("srcTable").item(rowidx).value = CStr(RowIDX)
                Next
            End If
            If Len(ColumnWidths) > 0 Then
                ColwidthArr = strToStringArray(ColumnWidths, ",")
                If DGV.ColumnCount < UBound(ColwidthArr) Then
                    Message = "Problem: Mismatched columns - Column Count is LESS THAN Number of Columns Passed"
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource_With_ColumnWidths()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                End If
                If DGV.ColumnCount > UBound(ColwidthArr) Then
                    Message = "Problem: Mismatched columns - Column Count is GREATER THAN Number of Columns Passed"
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource_With_ColumnWidths()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                End If
                'DataGridViewAutoSizeColumnsMode = FILL
                IDX = 0
                While IDX < UBound(ColwidthArr)
                    DGV.Columns(IDX).FillWeight = CInt(ColwidthArr(IDX))
                    IDX = IDX + 1
                End While
            End If
            DSource.DataSource = ds.Tables("srcTable")
            con.Close()
        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            Message = Message & vbCrLf & "Exception Error in PopulateMyDataSource: " & ex.ToString & ", SQL= " & SQLStr
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "PopulateMyDataSource_With_ColumnWidths()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)

        End Try


    End Sub

    Function Find_myDate(ByVal connString As String, ByVal DBTable As String, ByVal SearchDateFromField As String,
                            ByVal SearchDateFromText As String, SearchDateToField As String, SearchDateToText As String,
                         ByRef AllFieldNames As String(), ByRef AllFieldValues As Object(), ByVal Criteria As String,
                         Optional ByVal SortField As String = "", Optional ByVal Reversed As Boolean = False,
                         Optional ByRef Message As String = "") As Boolean

        Dim QryStr As String
        Dim cmd As MySqlCommand
        Dim con As MySqlConnection
        Dim WholeRow As Object()
        Dim Fieldnames As String()
        Dim IDX As Integer
        Dim numFields As Integer
        Dim dr1 As MySqlDataReader
        Dim myReader As MySqlDataReader
        Dim SearchFieldType As Type = Nothing
        Dim ReturnValue As String = ""

        Find_myDate = False
        'ReDim WholeRow(1)
        If Len(DBTable) = 0 Then
            'MsgBox("Error In FIND_ME: No DB Table specified")
            Message = "Error in Find_myQuery: No DB Table specified"
            Find_myDate = False
            Exit Function
        End If
        ReturnValue = ""
        Find_myDate = False
        Try
            con = New MySqlConnection(connString)
            con.Open()
            QryStr = "SELECT * FROM " & DBTable & " WHERE (DATE(" & SearchDateFromField & ") "
            QryStr = QryStr & " BETWEEN " & Chr(39) & SearchDateFromText & Chr(39)
            QryStr = QryStr & " AND " & Chr(39) & SearchDateToText & Chr(39)


            If Len(Criteria) > 0 Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " = '" & SearchText & "')"
                QryStr = QryStr & " AND " & Criteria
            End If
            If Len(SortField) > 0 Then
                QryStr = QryStr + " ORDER BY " + SortField
                If Reversed = False Then
                    QryStr = QryStr + " ASC"
                Else
                    QryStr = QryStr + " DESC"
                End If
            End If
            '
            cmd = New MySqlCommand(QryStr, con)
            dr1 = cmd.ExecuteReader()

            numFields = dr1.FieldCount - 1
            ReDim Fieldnames(numFields)
            IDX = 0
            While IDX <= numFields
                Fieldnames(IDX) = dr1.GetName(IDX)
                IDX = IDX + 1
            End While
            AllFieldNames = Fieldnames
            dr1.Close()
            ReDim WholeRow(numFields + 1)
            myReader = cmd.ExecuteReader()
            If myReader.HasRows Then
                'myReader.GetValues(WholeRow)
                'AllFieldValues = WholeRow
                While myReader.Read()
                    'DescriptionText.Text = dr("Description").ToString
                    'CostText.Text = dr("Cost").ToString
                    'PriceText.Text = dr("Price").ToString

                    'WHAT IF THE SEARCH TYPE is not STRING here ??????
                    Find_myDate = True
                    myReader.GetValues(WholeRow)
                    AllFieldValues = WholeRow

                End While
            Else
                'No Rows
                Find_myDate = False
            End If
            con.Close()
            myReader.Close()

        Catch ex As Exception
            Message = "Exception Error in Find_myDate : " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "Find_myDate()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
    End Function

    Public Function Find_myQuery(ByVal connString As String, ByVal DBTable As String, ByVal SearchField As String,
                            ByVal SearchText As String, ByVal SearchType As String, ByVal ReturnField As String,
                            ByRef ReturnValue As String, Optional ByRef AllFieldValues As Object() = Nothing,
                            Optional ByRef AllFieldNames As String() = Nothing, Optional ByVal Criteria As String = "",
                            Optional SortField As String = "ID", Optional ByVal Reversed As Boolean = False,
                            Optional ByRef Message As String = "", Optional ByVal Operation As String = "=",
                            Optional ByRef numRows As Long = 0) As Boolean
        Dim QryStr As String
        Dim cmd As MySqlCommand
        Dim con As MySqlConnection
        Dim WholeRow As Object()
        Dim Fieldnames As String()
        Dim FieldTypes As Type()
        Dim Fieldindex As Integer = 0
        Dim IDX As Integer
        Dim Rows As Long = 0
        Dim numFields As Integer
        Dim dr1 As MySqlDataReader
        Dim myReader As MySqlDataReader
        Dim SearchFieldType As Type = Nothing
        Dim CurValue As String = ""

        '**********************************************************************************************************************************
        '*************************** SEARCH CRITERIA MAY BE DIFFERENT TO UPDATE CRITERIA !!!!!!!! *****************************************
        '**********************************************************************************************************************************

        Find_myQuery = False
        'ReDim WholeRow(1)
        Message = ""
        If Len(DBTable) = 0 Then
            'MsgBox("Error In FIND_ME: No DB Table specified")
            Message = "Error in Find_myQuery: No DB Table specified"
            Find_myQuery = False
            Exit Function
        End If
        ReturnValue = ""
        Try
            con = New MySqlConnection(connString)
            con.Open()
            If UCase(SearchType) = "STRING" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " " & Operation & " '" & SearchText & "')"
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & "'" & SearchText & "')"
            ElseIf UCase(SearchType) = "INTEGER" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & ")"
            ElseIf UCase(SearchType) = "DATE" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (DATE(" & SearchField & ") " & Operation & Chr(39) & SearchText & Chr(39) & ")"


            Else
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & "f" & ")"
            End If
            If Len(Criteria) > 0 Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " = '" & SearchText & "')"
                QryStr = QryStr & " AND " & Criteria
            End If
            If Len(SortField) > 0 Then
                QryStr = QryStr + " ORDER BY " + SortField
                If Reversed = False Then
                    QryStr = QryStr + " ASC"
                Else
                    QryStr = QryStr + " DESC"
                End If
            End If
            '
            cmd = New MySqlCommand(QryStr, con)
            dr1 = cmd.ExecuteReader()

            numFields = dr1.FieldCount - 1
            ReDim Fieldnames(numFields + 1)
            ReDim FieldTypes(numFields + 1)
            IDX = 0
            While IDX <= numFields
                Fieldnames(IDX) = dr1.GetName(IDX)
                FieldTypes(IDX) = dr1.GetFieldType(IDX)
                If UCase(Fieldnames(IDX)) = UCase(SearchField) Then
                    SearchFieldType = dr1.GetFieldType(IDX)
                    'Exit While - need to loop completely to get all of the FIELDNAMES. else use dr1.Fields.ColumnName(index)
                    'Could use FieldnameExists = dr1.GetSchemaTable.Columns.Contains(Fieldname)
                End If
                IDX = IDX + 1
            End While
            AllFieldNames = Fieldnames
            dr1.Close()
            ReDim WholeRow(numFields + 1)
            myReader = cmd.ExecuteReader()
            If IsNothing(myReader) Then
                Exit Function
            End If
            If myReader.HasRows Then
                'myReader.GetValues(WholeRow)
                'AllFieldValues = WholeRow
                While myReader.Read()
                    'DescriptionText.Text = dr("Description").ToString
                    'CostText.Text = dr("Cost").ToString
                    'PriceText.Text = dr("Price").ToString

                    'WHAT IF THE SEARCH TYPE is not STRING here ??????
                    If InStr(Operation, ">") > 0 Then
                        ReturnValue = myReader(ReturnField).ToString
                        Find_myQuery = True
                        myReader.GetValues(WholeRow)
                        AllFieldValues = WholeRow
                        'myReader.Close()
                        'The following should only be executed if ONE value - the FIRST value found is required.
                        'Otherwise - store in a 2-dim array and return that in the array passed.
                        Exit While
                    Else
                        If UCase(myReader(SearchField).ToString) = UCase(SearchText) Then
                            If Len(ReturnField) > 0 Then
                                ReturnValue = myReader(ReturnField).ToString
                                Find_myQuery = True
                            End If
                            'Got to extract the whole row:
                            myReader.GetValues(WholeRow)
                            'myReader.FieldCount
                            CurValue = myReader.GetString(Fieldindex)
                            AllFieldValues = WholeRow
                            Find_myQuery = True
                            'myReader.Close()
                            Exit While
                        End If
                    End If
                    numRows = numRows + 1
                End While

            Else
                'No Rows
                Find_myQuery = False
                numRows = 0
            End If
            If (Not myReader Is Nothing) Then
                myReader.Close()
            End If
            con.Close()

        Catch ex As Exception
            Message = "Exception Error in Find_myQuery : " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "Find_myQuery()", frmMainGIForm.GetComputerName() &
                                        ", DB Table:" & DBTable & ", " & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try


    End Function

    Public Function Find_myMultiQuery(ByVal connString As String, ByVal DBTable As String, ByVal SearchField As String,
                            ByVal SearchText As String, ByVal SearchType As String, ByVal ReturnField As String,
                            ByRef ReturnValue As String, Optional ByRef AllFieldValues As Object() = Nothing,
                            Optional ByRef AllFieldNames As String() = Nothing, Optional ByVal Criteria As String = "",
                            Optional SortField As String = "ID", Optional ByVal Reversed As Boolean = False,
                            Optional ByRef Message As String = "", Optional ByVal Operation As String = " = ",
                            Optional ByVal ReturnSingleValue As Boolean = True, Optional ByRef Return_2Dim_Array As Object(,) = Nothing,
                                 Optional ByRef numRows As Long = 0, Optional ByRef LastReturnRow As Long = -1) As Boolean

        Dim QryStr As String
        Dim cmd As MySqlCommand
        Dim con As MySqlConnection
        Dim WholeRow As Object()
        Dim Fieldnames As String()
        Dim Fieldindex As Integer = 0
        Dim IDX As Integer
        Dim TotalRows As Long = 0
        Dim numFields As Integer
        Dim dr1 As MySqlDataReader
        Dim myReader As MySqlDataReader
        Dim SearchFieldType As Type = Nothing
        Dim CurValue As String = ""
        Dim ReturnRowIdx As Long = -1

        Find_myMultiQuery = False
        numRows = -1
        'ReDim WholeRow(1)
        Message = ""
        If Len(DBTable) = 0 Then
            'MsgBox("Error In FIND_ME: No DB Table specified")
            Message = "Error in Find_myQuery: No DB Table specified"
            Find_myMultiQuery = False
            Exit Function
        End If
        ReturnValue = ""
        Find_myMultiQuery = False
        Try

            If UCase(SearchType) = "STRING" Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " " & Operation & " '" & SearchText & "')"
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & "'" & SearchText & "')"
            ElseIf UCase(SearchType) = "INTEGER" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & ")"
            ElseIf UCase(SearchType) = "DATE" Then
                QryStr = "SELECT * FROM " & DBTable & " WHERE (DATE(" & SearchField & ") " & Operation & Chr(39) & SearchText & Chr(39) & ")"
            Else
                QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & Operation & SearchText & "f" & ")"
            End If
            If Len(Criteria) > 0 Then
                'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " = '" & SearchText & "')"
                QryStr = QryStr & " AND " & Criteria
            End If
            If Len(SortField) > 0 Then
                QryStr = QryStr + " ORDER BY " + SortField
                If Reversed = False Then
                    QryStr = QryStr + " ASC"
                Else
                    QryStr = QryStr + " DESC"
                End If
            End If
            '
            con = New MySqlConnection(connString)
            con.Open()
            cmd = New MySqlCommand(QryStr, con)
            dr1 = cmd.ExecuteReader()
            While dr1.Read()
                TotalRows = TotalRows + 1
            End While
            numFields = dr1.FieldCount - 1
            ReDim Fieldnames(numFields)
            IDX = 0
            LastReturnRow = -1
            While IDX <= numFields
                Fieldnames(IDX) = dr1.GetName(IDX)
                If UCase(Fieldnames(IDX)) = UCase(SearchField) Then
                    SearchFieldType = dr1.GetFieldType(IDX)
                    'Exit While
                End If
                IDX = IDX + 1
            End While
            AllFieldNames = Fieldnames
            dr1.Close()
            ReDim WholeRow(numFields + 1)
            ReDim Return_2Dim_Array(TotalRows + 1, numFields + 1)
            myReader = cmd.ExecuteReader()
            If myReader.HasRows Then
                ReturnRowIdx = 0
                numRows = 0
                'myReader.GetValues(WholeRow)
                'AllFieldValues = WholeRow
                While myReader.Read()
                    'DescriptionText.Text = dr("Description").ToString
                    'CostText.Text = dr("Cost").ToString
                    'PriceText.Text = dr("Price").ToString

                    'WHAT IF THE SEARCH TYPE is not STRING here ??????
                    If InStr(Operation, ">") > 0 Then 'Return the details of the last record found.
                        ReturnValue = myReader(ReturnField).ToString
                        Find_myMultiQuery = True
                        myReader.GetValues(WholeRow)
                        AllFieldValues = WholeRow
                        'myReader.Close()
                        'The following should only be executed if ONE value - the FIRST value found is required.
                        'Otherwise - store in a 2-dim array and return that in the array passed.
                        For Fieldindex = 0 To numFields - 1
                            Return_2Dim_Array(ReturnRowIdx, Fieldindex) = myReader.GetString(Fieldindex)
                        Next
                        If Len(ReturnField) > 0 Then
                            ReturnValue = myReader(ReturnField).ToString
                            LastReturnRow = numRows
                            Find_myMultiQuery = True
                        End If
                        If ReturnSingleValue = True Then
                            Exit While
                        End If
                    Else
                        If UCase(myReader(SearchField).ToString) = UCase(SearchText) Then
                            myReader.GetValues(WholeRow)
                            For Fieldindex = 0 To numFields - 1
                                Return_2Dim_Array(ReturnRowIdx, Fieldindex) = myReader.GetString(Fieldindex)
                            Next
                            If Len(ReturnField) > 0 Then
                                ReturnValue = myReader(ReturnField).ToString
                                LastReturnRow = numRows
                                Find_myMultiQuery = True
                            End If
                            AllFieldValues = WholeRow
                            Find_myMultiQuery = True
                            'myReader.Close()
                            If ReturnSingleValue = True Then
                                Exit While
                            End If
                        End If
                    End If
                    numRows = numRows + 1
                    ReturnRowIdx = ReturnRowIdx + 1
                End While

            Else
                'No Rows
                Find_myMultiQuery = False
                numRows = 0
            End If
            If (Not myReader Is Nothing) Then
                myReader.Close()
            End If
            con.Close()

        Catch ex As Exception
            Message = "Exception Error in Find_mymultiQuery : " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "Find_mymultiQuery()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

    End Function

    Function InsertUpdateMyRecord(ByVal UPDATE As Boolean, ByVal connString As String, DBTable As String,
                                  FieldNames As String, FieldValues As String, ByRef MyMessage As String,
                                  Optional ByVal Criteria As String = "",
                                  Optional ByVal ExcludeFields As String = "",
                                  Optional ByVal ElementStartIDX As Integer = 0,
                                  Optional ByVal EncaseFields As Boolean = True,
                                  Optional ByVal RemoveBadChars As Boolean = True,
                                  Optional ByVal IncludeComma As Boolean = True,
                                  Optional ByVal IncludeQuotes As Boolean = False) As Boolean

        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim strSQL As String
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Message As String = ""

        'Dim sql = "INSERT INTO Contacts (FName, LName, Age, " &
        '  "[Address Line 1], [Address Line 2], City, State, Zip, " &
        ' "[Home Phone], [Work Phone], Email, Sex) " &
        '"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        NumBlankFields = 0
        NumFields = 0
        NumRows = 0
        strSQL = ""
        DataSource = ""
        'connString = ""
        ExcludedFields = ""
        InsertUpdateMyRecord = False
        If Len(DBTable) = 0 Then
            'MsgBox("Error in Insert Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            MyMessage = "Error in Insert Record: Database Table not provided"
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "InsertUpdateMyRecord()",
                                        frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername & ", DBTable: " & DBTable & ", Port: ")
            InsertUpdateMyRecord = False
            Exit Function
        End If

        If Len(FieldNames) = 0 Then
            FieldNames = GetMyFields(DBTable, connString, Message)
            If Len(Message) > 0 Then
                MyMessage = "Problem getting Fieldnames"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "InsertUpdateMyRecord()",
                                            frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                            ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                Exit Function
            End If
        End If
        'NumBlankFields = CheckBlankValues(FieldNames, FieldValues, ExcludedFields) 'any blank values passed in fieldvalues will put corresponding fieldnames into ExcludedFields output.
        If Len(ExcludeFields) > 0 Then
            If Len(ExcludedFields) = 0 Then
                ExcludedFields = ExcludeFields
            Else
                ExcludedFields = ExcludedFields & "," & ExcludeFields
            End If
        End If
        If UPDATE = True Then
            strSQL = PrepareMyUpdate(connString, DBTable, FieldNames, FieldValues, MyMessage, Criteria, ExcludedFields,
                                     ElementStartIDX, EncaseFields, RemoveBadChars, IncludeComma, IncludeQuotes)
        Else
            strSQL = PrepareMyInsert(connString, DBTable, FieldNames, FieldValues, MyMessage, ExcludedFields,
                                     ElementStartIDX, EncaseFields, RemoveBadChars, IncludeComma, IncludeQuotes)
        End If
        If Len(strSQL) = 0 Then
            MyMessage = "Error in InsertUpdateMyRecord : Query passed is blank"
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "InsertUpdateMyRecord()",
                                        frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            InsertUpdateMyRecord = False
            Exit Function
        End If
        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(strSQL, con)

            'Using con
            'Using cmd
            con.Open()
            'cmd.Parameters.AddWithValue("@p1", "Value For FName")
            'cmd.Parameters.AddWithValue("@p2", "Value For LName")
            'cmd.Parameters.AddWithValue("@p3", Convert.ToInt32("Value For Age"))
            '....and so On For the other parameters .....
            cmd.ExecuteNonQuery()
            'End Using
            'End Using
            con.Close()


        Catch ex As Exception
            MyMessage = "Exception Error in InsertUpdateRecord Record: " & ex.ToString
            InsertUpdateMyRecord = False
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "InsertUpdateMyRecord()",
                                        frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function
        End Try
        InsertUpdateMyRecord = True 'needs a test to make sure the record HAS been inserted. Use the OleDbDataReader.RecordsAffected property

    End Function

    Function DeleteMyRecord(ByVal DBTable As String, ByVal connString As String, ByVal Criteria As String, ByRef MyMessage As String) As Boolean
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim strSQL As String
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim EqualPos As Integer

        NumBlankFields = 0
        NumFields = 0
        NumRows = 0
        strSQL = ""
        DataSource = ""
        'connString = ""
        ExcludedFields = ""
        DeleteMyRecord = False
        If Len(DBTable) = 0 Then
            'MsgBox("Error in Delete Record: Database Table not provided", vbOK, "Error in Asset Register " & frmMainGIForm.myVersion)
            MyMessage = "Error in Delete Record: Database Table not provided"
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "DeleteMyRecord()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            DeleteMyRecord = False
            Exit Function
        End If
        strSQL = "DELETE FROM " & DBTable

        EqualPos = InStr(Criteria, "=")
        If Len(Criteria) > 0 And EqualPos > 0 Then
            If EqualPos < Len(Criteria) Then
                strSQL = strSQL & " WHERE " & Criteria
            End If
        End If
        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(strSQL, con)

            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()

        Catch ex As Exception
            'MsgBox("Exception Error in InsertUpdateRecord Record: " & ex.ToString, vbOK, "Exception Error in AssetRegister v" & frmMainGIForm.myVersion)
            MyMessage = "Exception Error in DeleteMyRecord: " & ex.Message.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "DeleteMyRecord()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function
        End Try
        DeleteMyRecord = True 'needs a test to make sure the record HAS been Deleted. Use the OleDbDataReader.RecordsAffected property

    End Function

    Function PrepareMyInsert(ByVal connString As String, ByVal DBTable As String, ByRef Fieldnames As String, ByVal Fieldvalues As String,
                             ByRef myMessage As String,
                             Optional ByRef ExcludeFields As String = "",
                             Optional ByVal ElementStartIDX As Integer = 0,
                             Optional ByVal EncaseFields As Boolean = True,
                             Optional ByVal RemoveBadChars As Boolean = True,
                             Optional ByVal IncludeComma As Boolean = True,
                             Optional ByVal IncludeQuotes As Boolean = False) As String
        Dim FieldNameArray As String()
        Dim IgnoreFieldsArray As String()
        Dim ValueArray As String()
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim NumFields As Integer
        Dim NewFieldsArray As String()

        FinalCMD = ""
        ReDim FieldNameArray(1)
        ReDim NewFieldsArray(1)
        ReDim IgnoreFieldsArray(1)
        Try
            If Len(ExcludeFields) > 0 Then
                IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
            End If
            If Len(DBTable) = 0 Then
                'MsgBox("Error in PrepareInsert: No Database Table specified", vbOK, "Error in Asset Register")
                myMessage = "Error in PrepareInsert: No Database Table specified"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                PrepareMyInsert = ""
                Exit Function
            End If
            If Len(Fieldnames) = 0 Then
                Fieldnames = GetMyFields(DBTable, connString, myMessage)
            End If
            FieldNameArray = strToStringArray(Fieldnames, ",")
            If Len(Fieldvalues) > 0 Then
                ValueArray = strToStringArray(Fieldvalues, ",", ElementStartIDX, EncaseFields, RemoveBadChars, IncludeComma, "_", IncludeQuotes)
            Else
                'MsgBox("Error in PrepareInsert: No values specified", vbOK, "Error in Asset Register")
                myMessage = "Error in PrepareInsert: No values specified"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                PrepareMyInsert = ""
                Exit Function
            End If
            Application.DoEvents()

            Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray)
            If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then

                If UBound(FieldNameArray) < UBound(ValueArray) Then 'Unlikely - as this means too many values passed
                    'MsgBox("Error in PrepareInsert: Number of Fields passed are LESS THAN Number of Values Passed.", vbOK, "MISS-MATCH Error in Asset Register")
                    'PrepareInsert = ""
                    'Exit Function
                    myMessage = myMessage & vbCrLf & "Error in PrepareInsert: Number of VALUES passed EXCEED Number of Fields in database."
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    'frmMainGIForm.logger.
                End If
                If UBound(FieldNameArray) > UBound(ValueArray) Then
                    myMessage = "Error in PrepareInsert: Number of Fields passed are GREATER THAN Number of Values Passed."
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    'Not necessarily an Error but the Number of VALUES passed is LESS THAN Number of Fields in Database.
                    'So use WHILE IDX < UBOUND(ValueArray) 
                    'rather than IDX < Ubound(FieldNameArray) 
                    'PrepareInsert = ""
                    'Exit Function
                End If
            End If
            IDX = 0
            FinalCMD = "INSERT INTO " & DBTable & " (" & Fieldnames & ") VALUES (" & Fieldvalues & ")"
            'MsgBox("FinalCMD = " & FinalCMD)
        Catch ex As Exception
            'MsgBox("Error Exception in PrepareInsert: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            myMessage = myMessage & vbCrLf & "Error Exception in PrepareInsert: " & ex.ToString & ", FinalCMd= " & FinalCMD
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        PrepareMyInsert = FinalCMD


    End Function

    Function PrepareMyUpdate(ByVal connString As String, ByVal TableName As String,
                           ByRef Fieldnames As String, ByVal Fieldvalues As String, ByRef myMessage As String,
                           Optional ByVal criteria As String = "",
                           Optional ByRef ExcludeFields As String = "",
                           Optional ByVal ElementStartIDX As Integer = 0,
                           Optional ByVal EncaseFields As Boolean = True,
                           Optional ByVal RemoveBadChars As Boolean = True,
                           Optional ByVal IncludeComma As Boolean = True,
                           Optional ByVal IncludeQuotes As Boolean = False) As String
        Dim FieldNameArray As String()
        Dim FinalFields As String()
        Dim typeFieldNames As String
        Dim IgnoreFieldsArray As String()
        Dim ValueArray As String()
        Dim strFieldTypes As String
        Dim FieldTypes() As String
        Dim myDic_FieldTypes As Object
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim fldValue As String
        Dim NumFields As Integer
        Dim UpdateCmd As String = "UPDATE"
        Dim FinalFieldValues As String = ""
        Dim dtDateTime As DateTime
        Dim FieldType As String

        FinalCMD = ""
        ReDim FieldNameArray(1)
        ReDim FinalFields(1)
        ReDim IgnoreFieldsArray(1)
        ReDim ValueArray(1)
        Try
            'If Len(IgnoreFields) > 0 Then
            'IgnoreFieldsArray = strToStringArray(IgnoreFields, ",")
            'End If
            myDic_FieldTypes = CreateObject("Scripting.Dictionary")
            myDic_FieldTypes.comparemode = vbTextCompare
            typeFieldNames = GetMyFields(TableName, connString, myMessage, "", strFieldTypes, 0, myDic_FieldTypes)
            If Len(TableName) = 0 Then
                'MsgBox("Error in PrepareUpdate: No Database Table specified", vbOK, "Error in Asset Register")
                myMessage = "Error in PrepareMyUpdate: No Database Table specified"
                PrepareMyUpdate = ""
                Exit Function
            End If
            If Len(Fieldnames) = 0 Then
                'NumFields = GetNumFields(connString, "SELECT * FROM " & TableName, DBName, Fieldnames)
                Fieldnames = GetMyFields(TableName, connString, myMessage, "", strFieldTypes)
            End If

            FieldNameArray = strToStringArray(Fieldnames, ",")


            If Len(Fieldvalues) > 0 Then
                ValueArray = strToStringArray(Fieldvalues, ",", ElementStartIDX, False, RemoveBadChars, IncludeComma, "_", IncludeQuotes, 34, 39)
            Else
                'MsgBox("Error in PrepareInsert: No values specified", vbOK, "Error in Asset Register")
                myMessage = "Error in PrepareInsert: No values specified"
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyInsert()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                PrepareMyUpdate = ""
                Exit Function
            End If

            If Len(ExcludeFields) > 0 Then
                IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
            End If
            Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray) 'rebuilds whole list without the extracted fields
            FieldNameArray = strToStringArray(Fieldnames, ",")
            IDX = 0
            UpdateCmd = "UPDATE " & TableName & " SET "
            Fieldnames = ""
            For IDX = 0 To UBound(FieldNameArray)
                'If String.IsNullOrEmpty(ValueArray(IDX)) Then
                'If Len(ValueArray(IDX)) = 0 Or IsNothing(ValueArray(IDX)) Then
                'ExcludeFields = ExcludeFields & "," & FieldNameArray(IDX)
                'End If
                'Else
                If ValueArray(IDX) Is Nothing Then
                    Continue For
                End If
                Application.DoEvents()

                If InStr(ValueArray(IDX), "/") > 0 Then
                    ValueArray(IDX) = Replace(ValueArray(IDX), "/", "-")

                    If InStr(ValueArray(IDX), ":") > 0 Then
                        'The Date DOES contain a TIME also
                    Else
                        'ValueArray(IDX) = ValueArray(IDX) & " 01:00:00"

                    End If

                End If

                'FieldType = Dic_FieldTypes(TableName & "_" & FieldNameArray(IDX))

                FieldType = myDic_FieldTypes(FieldNameArray(IDX))

                If IsDate(ValueArray(IDX)) And FieldType = "DATETIME" Then
                    dtDateTime = CDate(ValueArray(IDX))
                    ValueArray(IDX) = dtDateTime.ToString("yyyy-MM-dd HH:mm:ss")

                End If
                If IDX = 0 Then
                    FinalFieldValues = ValueArray(IDX)
                    Fieldnames = FieldNameArray(IDX)
                Else
                    FinalFieldValues = FinalFieldValues & "," & ValueArray(IDX)
                    Fieldnames = Fieldnames & "," & FieldNameArray(IDX)
                End If
            Next
            'IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
            'Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray) 'rebuilds whole list without the extracted fields
            FieldNameArray = strToStringArray(Fieldnames, ",")
            'ReDim Preserve FieldNameArray(IDX - 1)
            'The following block needs to go first - element mis-match as the ID field is still included in the above test.
            If Len(FinalFieldValues) > 0 Then
                ValueArray = strToStringArray(FinalFieldValues, ",", ElementStartIDX, EncaseFields, RemoveBadChars, IncludeComma, "_", IncludeQuotes)
                'ValueArray = strToStringArray(FinalFieldValues, ",")
            End If

            If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
                If UBound(FieldNameArray) < UBound(ValueArray) Then
                    'MsgBox("Error in PrepareUpdate: Number of Fields Passed are LESS than Number of VALUES passed.", vbOK, "Miss-Match Error in Asset Register " & frmMainGIForm.myVersion)
                    myMessage = myMessage & vbCrLf & "Error in PrepareMyUpdate: Number of Fields Passed are LESS than Number of VALUES passed."
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyUpdate()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    PrepareMyUpdate = ""
                    Exit Function
                End If
                If UBound(FieldNameArray) > UBound(ValueArray) Then
                    'MsgBox("Error in PrepareUpdate: Number of Fields Passed are GREATER than Number of VALUES passed.", vbOK, "Miss-Match Error in Asset Register " & frmMainGIForm.myVersion)
                    myMessage = myMessage & vbCrLf & "Error in PrepareMyUpdate: Number of Fields Passed are GREATER than Number of VALUES passed."
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyUpdate()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    PrepareMyUpdate = ""
                    Exit Function
                End If
            End If
            IDX = 0
            UpdateCmd = "UPDATE " & TableName & " SET "
            For IDX = 0 To UBound(FieldNameArray) - 1
                fldName = FieldNameArray(IDX)
                fldValue = ValueArray(IDX)
                If Len(fldValue) > 0 And fldValue IsNot Nothing Then

                    If IDX = 0 Then
                        If InStr(fldValue, ":") > 0 Then
                            'FinalCMD = fldName & " = " & Chr(34) & fldValue & Chr(34)
                            FinalCMD = fldName & " = " & fldValue
                        Else
                            FinalCMD = fldName & " = " & fldValue
                        End If
                    Else
                        If InStr(fldValue, ":") > 0 Then
                            'FinalCMD = FinalCMD & "," & fldName & " = " & Chr(34) & fldValue & Chr(34)
                            FinalCMD = FinalCMD & "," & fldName & " = " & fldValue
                        Else
                            FinalCMD = FinalCMD & "," & fldName & " = " & fldValue
                        End If
                    End If
                Else
                    'No value
                    'myMessage = "fldValue(" & CStr(IDX) & ") : " & fldValue(IDX) & " LENGTH= " & CStr(Len(fldValue(IDX)))
                    myMessage = "Item = 0 chars or less: IDX=" & CStr(IDX)
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyUpdate()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                End If
                Application.DoEvents()
            Next
            FinalCMD = UpdateCmd & FinalCMD
            If Len(criteria) > 0 Then
                FinalCMD = FinalCMD & " WHERE " & criteria
            End If

        Catch ex As Exception
            'MsgBox("Error Exception in PrepareUpdate: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            myMessage = myMessage & vbCrLf & "Error Exception in PrepareMyUpdate: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "PrepareMyUpdate()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        PrepareMyUpdate = FinalCMD


    End Function

    Function CreateDBUser(ByVal connString As String, ByVal Hostname As String, ByVal Username As String, ByVal Password As String, Optional ByRef MyMessage As String = "") As Boolean
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim strSQL As String
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Message As String = ""

        CreateDBUser = False
        strSQL = "CREATE USER " & Username & "@" & Hostname & " IDENTIFIED BY " & Chr(39) & Password & Chr(39) & ";"

        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(strSQL, con)

            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()


        Catch ex As Exception
            MyMessage = "Exception Error in CreateDBUser: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "CreateDBUser()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        If Len(MyMessage) = 0 Then
            CreateDBUser = True
        End If

    End Function

    Function myPermissions(ByVal DATABASE As String, ByVal connString As String, ByVal Username As String, ByVal Password As String,
                           ByVal Permissions As String, ByVal Hostname As String,
                           Optional ByRef MyMessage As String = "", Optional ByVal DBTAble As String = "*", Optional ByVal AllowGrant As Boolean = False) As Boolean
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim DataSource As String
        Dim strSQL As String = ""
        Dim ExcludedFields As String
        Dim NumRows As Integer
        Dim NumFields As Integer
        Dim NumBlankFields As Integer
        Dim Message As String = ""
        Dim delimList As String = ""
        Dim AllowAddUser As Boolean = False

        myPermissions = False
        If Len(Permissions) = 0 Then
            'Create default permissions = SELECT, UPDATE, DELETE , INSERT
            'if AllowGrant is TRUE then add "GRANT" to the collection.
            'ASSUME NORMAL
            AllowAddUser = False
            delimList = "INSERT,DELETE,UPDATE,SELECT"
        Else
            If UCase(Permissions) = "NORMAL" Then
                delimList = "INSERT,DELETE,UPDATE,SELECT"
                AllowAddUser = False
            ElseIf UCase(Permissions) = "ADMIN" Then
                delimList = "INSERT,DELETE,UPDATE,SELECT"
                AllowAddUser = AllowGrant
            ElseIf UCase(Permissions) = "SUPER" Then
                delimList = "ALL"
                AllowAddUser = AllowGrant
            Else
                delimList = "INSERT,DELETE,UPDATE,SELECT"
                AllowAddUser = False
                MsgBox("Unrecognised Permission: Setting to NORMAL")
            End If
        End If


        strSQL = strSQL & " GRANT " & delimList & " ON " & DATABASE & "." & DBTAble & " TO " & Chr(39) & Username & Chr(39) & "@" & Hostname
        If AllowAddUser = True Then
            strSQL = strSQL & " WITH GRANT OPTION"
        End If
        strSQL = strSQL & ";"
        MsgBox("Apply Permissions: " & delimList & " " & vbCrLf & strSQL)
        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(strSQL, con)

            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()


        Catch ex As Exception
            MyMessage = "Exception Error in CreateDBUser: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, MyMessage, "myPermissions()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try


        If Len(MyMessage) = 0 Then
            myPermissions = True
        End If

    End Function

    Function GetMYValuebyFieldname(ByVal SearchObjects As Object(), ByVal FieldnamesArr_IN As String(), ByVal ReturnField As String,
                                    Optional ByRef Message As String = "") As String
        Dim fieldidx As Int32
        Dim ReturnValue As String
        Dim NumFields As Integer

        'Need to cycle through the Fieldnames and find the correct index.
        'THEN apply it to the SearchObjects to get the value.
        'NumFields = GetNumFields("", "SELECT * FROM " & dbTable, "AssetRegister.accdb", Fieldnames)
        ReturnValue = ""
        GetMYValuebyFieldname = ""

        If IsNothing(SearchObjects) Then
            Exit Function
        End If
        If IsNothing(FieldnamesArr_IN) Then
            Exit Function
        End If
        NumFields = UBound(FieldnamesArr_IN)
        Try
            For fieldidx = 0 To NumFields
                If UCase(ReturnField) = UCase(FieldnamesArr_IN(fieldidx)) Then
                    ReturnValue = SearchObjects(fieldidx).ToString
                    Return ReturnValue
                    Exit For
                End If
            Next
        Catch ex As Exception
            Message = "Exception Error in GetMyValuebyFieldname: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetMYValuebyFieldname()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        GetMYValuebyFieldname = ReturnValue

    End Function

    Function GetMyMultiValuebyFieldname(ByVal SearchObjects As Object(,), ByVal Fieldnames As String(), ByVal ReturnField As String,
                                    Optional ByRef Message As String = "", Optional ByVal RowIndex As Long = 0) As String
        Dim fieldidx As Int32
        Dim ReturnValue As String
        Dim NumFields As Integer

        'Need to cycle through the Fieldnames and find the correct index.
        'THEN apply it to the SearchObjects to get the value.
        'NumFields = GetNumFields("", "SELECT * FROM " & dbTable, "AssetRegister.accdb", Fieldnames)
        ReturnValue = ""
        GetMyMultiValuebyFieldname = ""

        If IsNothing(SearchObjects) Then
            Exit Function
        End If
        NumFields = UBound(Fieldnames)
        Try
            For fieldidx = 0 To NumFields
                If UCase(ReturnField) = UCase(Fieldnames(fieldidx)) Then
                    ReturnValue = SearchObjects(RowIndex, fieldidx).ToString
                    Return ReturnValue
                    Exit For
                End If
            Next
        Catch ex As Exception
            Message = "Exception Error in GetValuebyFieldname: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetMyMultiValuebyFieldname()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        GetMyMultiValuebyFieldname = ReturnValue
    End Function

    Function CheckFieldExists(ByVal SearchField As String, ByVal Fieldlist As String, Optional ByRef Messages As String = "") As Boolean
        Dim fieldidx As Int32
        Dim ReturnValue As Boolean
        Dim NumFields As Integer
        Dim tempArray As String()

        'Need to cycle through the Fieldnames and find the correct index.
        'THEN apply it to the SearchObjects to get the value.
        'NumFields = GetNumFields("", "SELECT * FROM " & dbTable, "AssetRegister.accdb", Fieldnames)
        ReturnValue = False
        CheckFieldExists = False

        If Len(SearchField) = 0 Then
            Exit Function
        End If
        ReDim tempArray(1)
        tempArray = strToStringArray(Fieldlist, ",")
        NumFields = UBound(tempArray)
        Try
            For fieldidx = 0 To NumFields
                If UCase(SearchField) = UCase(tempArray(fieldidx)) Then
                    ReturnValue = True
                    Return ReturnValue
                    Exit For
                End If
            Next
        Catch ex As Exception
            Messages = "Exception Error in GetValuebyFieldname: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "CheckFieldExists()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        CheckFieldExists = ReturnValue
        tempArray = Nothing
    End Function

    Function GetMyTotalRows(ByVal connString As String, ByVal DBTable As String, ByRef Message As String) As Long
        Dim con As MySqlConnection
        Dim cmd As MySqlCommand
        Dim myReader As MySqlDataReader
        Dim da As MySqlDataAdapter
        Dim ds As DataSet
        Dim SQLStr As String = ""
        Dim NumRows As Long = 0
        Dim strNumRows As String = ""

        GetMyTotalRows = 0
        SQLStr = "SELECT COUNT(*) AS rows FROM " & DBTable
        Try
            con = New MySqlConnection(connString)
            cmd = New MySqlCommand(SQLStr, con)
            Application.DoEvents()
            con.Open()
            myReader = cmd.ExecuteReader()
            da = New MySqlDataAdapter(cmd)
            ds = New DataSet()
            If myReader.HasRows Then
                myReader.Read()
                strNumRows = myReader("rows").ToString
                NumRows = CLng(strNumRows)
            End If
            'da.Fill(ds, "srcTable") 'MYSQL DATE CONVERSION PROBLEM HERE !
            'If IsDBNull(ds.Tables("srcTable")) Then
            'MsgBox("Problem: Empty Rows", vbOK, "Problem in Asset Register v" & frmMainGIForm.myVersion)
            'Message = "Problem: Empty Rows"
            'NumRows = 0
            'Else
            'NumRows = ds.Tables("srcTable").Rows.Count
            'ds.Tables("srcTable").Columns.Add.Caption = "Row Num"
            'For RowIDX = 1 To NumRows
            'ds.Tables("srcTable").item(rowidx).value = CStr(RowIDX)
            'Next
            'End If
            'DSource = ds.Tables("srcTable")
            con.Close()
        Catch ex As Exception
            'MsgBox("Exception Error: " & ex.ToString, vbOK, "Exception Error in Asset Register")
            Message = Message & vbCrLf & "Exception Error in PopulateMyDataSource: " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Message, "GetMyTotalRows()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try
        GetMyTotalRows = NumRows

    End Function

    Function ReadInPairedDevices_From_CSV_Into_Table(ByVal CSVFilename As String, ByVal DBTable As String, Optional ByRef ErrorMessages As String = "") As Long
        Dim strAssetBCode1 As String = ""
        Dim strAssetBCode2 As String = ""
        Dim strAssetName1 As String = ""
        Dim strAssetGroup1 As String = ""
        Dim strAssetName2 As String = ""
        Dim strAssetGroup2 As String = ""
        Dim strPairedNumber As String = ""
        Dim TotalItems As Long = 0
        Dim TotalFields As Long = 0
        Dim AssetArray As String(,)
        Dim Messages As String = ""
        Dim RowIdx As Long = 0
        Dim ColIDX As Long = 0
        Dim SearchField As String = ""
        Dim SearchText As String = ""
        Dim ReturnField As String = ""
        Dim Criteria As String = ""
        Dim FieldValues As String = ""
        Dim strDateCreated As String = ""
        Dim strDateModified As String = ""
        Dim strErrorMessage As String = ""
        Dim impasse As String = ""
        Dim ExcludedFields As String = ""
        Dim Percentage As Double = 0.0


        ReadInPairedDevices_From_CSV_Into_Table = 0
        'ReDim AssetArray(1, 1)
        TotalItems = CSVFileToArray(AssetArray, CSVFilename, TotalFields)
        While RowIdx < UBound(AssetArray, 2)
            strAssetGroup1 = "PODS"
            strAssetGroup2 = "SCANNERS"
            strAssetBCode1 = GetFieldValue_From_Fieldname(AssetArray, RowIdx, "Pod EP number")
            strAssetBCode2 = GetFieldValue_From_Fieldname(AssetArray, RowIdx, "Scanner EP number")
            strPairedNumber = GetFieldValue_From_Fieldname(AssetArray, RowIdx, "Paired Number")
            strErrorMessage = GetFieldValue_From_Fieldname(AssetArray, RowIdx, "Any Errors?")
            If IsNumeric(strPairedNumber) Then
                If CInt(strPairedNumber) < 10 Then
                    strAssetName1 = "VP" & "0" & strPairedNumber
                    strAssetName2 = "SC" & "0" & strPairedNumber
                Else
                    strAssetName1 = "VP" & strPairedNumber
                    strAssetName2 = "SC" & strPairedNumber
                End If
                SearchField = "AssetBarcode"
                SearchText = strAssetBCode1
                ReturnField = "DateCreated"
                Criteria = ""
                ExcludedFields = "Comments"
                SaveAssetToTable("tblPairedAssets", SearchField, SearchText, ReturnField, Criteria, strAssetGroup1, strAssetName1,
                                 strAssetBCode1, strPairedNumber, ExcludedFields, ErrorMessages, RowIdx)
                SearchField = "AssetBarcode"
                SearchText = strAssetBCode2
                ReturnField = "DateCreated"
                Criteria = ""
                ExcludedFields = "Comments"
                SaveAssetToTable("tblPairedAssets", SearchField, SearchText, ReturnField, Criteria, strAssetGroup2, strAssetName2,
                                 strAssetBCode2, strPairedNumber, ExcludedFields, ErrorMessages, RowIdx)
            Else
                Messages = "Row " & CStr(RowIdx) & " Contains a non-numeric symbol in the Paired Number Column" & vbCrLf
                ErrorMessages = ErrorMessages & Messages
                frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, Messages, "ReadInPairedDevices_From_CSV_Into_Table()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                impasse = impasse & " " & strErrorMessage
            End If
            'Now we have pulled out the info - put it into the chosen database table:
            'Check if Asset has already been inserted into the table:



            Percentage = (RowIdx / UBound(AssetArray)) * 100

            Application.DoEvents()

            RowIdx = RowIdx + 1
        End While

    End Function

    Function SaveAssetToTable(ByVal DBTable As String, ByVal SearchField As String, ByVal SearchText As String, ByVal ReturnField As String, ByVal Criteria As String,
                              ByVal strAssetGroup As String, ByVal strAssetName As String, ByVal strAssetBarcode As String, ByVal strPairedNumber As String,
                               ByVal ExcludedFields As String, Optional ByRef ErrMessage As String = "", Optional ByVal RowNumber As Long = 0) As Boolean
        Dim strDateCreated As String = "1970-01-01 01:00:00"
        Dim strDateModified As String = "1970-01-01 01:00:00"
        Dim FieldValues As String = ""
        Dim AssetExists As Boolean = False
        Dim InsertOK As Boolean = False
        Dim ReturnValue As String = ""
        Dim strExcludedFields As String = ""
        Dim strErrMessage As String = ""

        SaveAssetToTable = False
        ErrMessage = ""
        FieldValues = Chr(34) & strAssetGroup & Chr(34) _
            & "," & Chr(34) & strAssetName & Chr(34) _
            & "," & Chr(34) & strAssetBarcode & Chr(34) _
            & "," & Chr(34) & strPairedNumber & Chr(34)
        '& "," & Chr(34) & strDateCreated & Chr(34) _
        '& "," & Chr(34) & strDateModified & Chr(34)

        If Len(DBTable) = 0 Then
            ErrMessage = "Blank Db Table passed."
            Exit Function
        End If
        If Len(ExcludedFields) = 0 Then
            strExcludedFields = "ID"
        Else
            strExcludedFields = "ID," & ExcludedFields
        End If
        If UCase(frmMainGIForm.DatabaseType) = "ACCDB" Then
            AssetExists = Find_me(frmMainGIForm.DBName, "", DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", "")
            If Not AssetExists Then
                strDateCreated = CStr(Now())
                strDateModified = CStr(Now())
                FieldValues = FieldValues & "," & Chr(34) & strDateCreated & Chr(34) _
                    & "," & Chr(34) & strDateModified & Chr(34)
                InsertOK = InsertUpdateRecord(False, frmMainGIForm.DBName, "", DBTable, "", FieldValues, Criteria, strExcludedFields)
                If Not InsertOK Then
                    ErrMessage = ErrMessage & ",Error in row: " & CStr(RowNumber)
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "SaveAssetToTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
            Else
                strDateModified = CStr(Now())
                strExcludedFields = strExcludedFields & "," & "DateCreated"
                FieldValues = FieldValues & "," & Chr(34) & strDateModified & Chr(34)
                InsertOK = InsertUpdateRecord(True, frmMainGIForm.DBName, "", DBTable, "", FieldValues, Criteria, strExcludedFields)
                If Not InsertOK Then
                    ErrMessage = ErrMessage & ",Error in row: " & CStr(RowNumber)
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "SaveAssetToTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
            End If
        Else
            If Module_DanG_MySQL_Tools.mylogged_in Then
                AssetExists = Find_myQuery(frmMainGIForm.myConnString, DBTable, SearchField, SearchText, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", False, strErrMessage)
                If Len(strErrMessage) > 0 Then
                    ErrMessage = ErrMessage & " " & strErrMessage & "; Row= " & CStr(RowNumber)
                    frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "SaveAssetToTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                    Exit Function
                End If
                If Not AssetExists Then
                    strDateCreated = CStr(Now())
                    strDateModified = CStr(Now())
                    FieldValues = FieldValues & "," & Chr(34) & strDateCreated & Chr(34) _
                        & "," & Chr(34) & strDateModified & Chr(34)
                    InsertOK = InsertUpdateMyRecord(False, frmMainGIForm.myConnString, DBTable, "", FieldValues, ErrMessage, Criteria, strExcludedFields)
                    If Not InsertOK Then
                        ErrMessage = ErrMessage & ",Error in row: " & CStr(RowNumber)
                        frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "SaveAssetToTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                        Exit Function
                    End If
                Else
                    strDateModified = CStr(Now())
                    strExcludedFields = strExcludedFields & "," & "DateCreated"
                    FieldValues = FieldValues & "," & Chr(34) & strDateModified & Chr(34)
                    InsertOK = InsertUpdateMyRecord(True, frmMainGIForm.myConnString, DBTable, "", FieldValues, ErrMessage, Criteria, strExcludedFields)
                    If Not InsertOK Then
                        ErrMessage = ErrMessage & ",Error in row: " & CStr(RowNumber)
                        frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "SaveAssetToTable()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                        Exit Function
                    End If
                End If
            Else
                'get user to log in:
                Module_DanG_MySQL_Tools.Check_myLoggedIN()

            End If

        End If

    End Function

    Function GetDetails(ByVal conString As String, ByVal SearchTable As String,
                            ByVal SearchField As String, ByVal SearchValue As String, ByVal ReturnField As String,
                            ByRef ReturnValue As String, ByRef Message As String, Optional ByVal Criteria As String = "") As Boolean
        Dim FoundDetail As Boolean = False

        GetDetails = False
        FoundDetail = Find_myQuery(frmMainGIForm.myConnString, SearchTable, SearchField, SearchValue, "STRING", ReturnField, ReturnValue, Nothing, Nothing, Criteria, "", False, Message)
        If FoundDetail Then
            GetDetails = True
        End If

    End Function

    Function CreateSearchSQL(ByVal DBTable As String, ByVal DisplayFields As String, DateField As String,
                     ByVal DateFrom As String, ByVal DateTo As String, Optional ByVal SortField As String = "",
                     Optional ByVal Reversed As Boolean = False, Optional ByVal Criteria As String = "", Optional ByRef MyMessage As String = "") As String
        Dim dtDateTo As DateTime = Nothing
        Dim dtDateFrom As DateTime = Nothing
        Dim strDateTo As String = ""
        Dim strDateFrom As String = ""
        Dim Comments As String = ""
        Dim UpdatedOK As Boolean = False
        Dim FieldNames As String = ""
        Dim FieldValues As String = ""
        Dim RegisterTableName As String = ""
        Dim DBaseName As String = "AssetRegister.accdb"
        Dim mySearchField As String = ""
        Dim SearchText As String = ""
        Dim TotalExcluded As Integer = 0
        Dim NumFields As Integer = 0
        Dim Provider As String = ""
        Dim FoundID As Boolean = False
        Dim QryStr As String = ""
        Dim MainSortField As String = ""
        Dim MaxGridCols As Integer = 0
        Dim strID As String = ""
        Dim SearchCriteria As String = ""
        Dim AllFields As String = ""

        CreateSearchSQL = ""
        strDateTo = DateTo
        strDateFrom = DateFrom
        If Len(DBTable) = 0 Then
            MyMessage = "SearchFields: No Table Specified"
            Exit Function
        End If
        If Len(SortField) > 0 Then
            MainSortField = SortField
        End If
        If Len(DisplayFields) > 0 Then
            QryStr = "SELECT " & DisplayFields & " FROM " & DBTable
        Else
            QryStr = "SELECT * FROM " & DBTable
        End If

        If Len(DateFrom) > 0 And Len(DateTo) = 0 Then
            QryStr = QryStr & " WHERE DATE(" & DateField & ") = " & Chr(39) & strDateFrom & Chr(39)
            If Len(Criteria) > 0 Then
                QryStr = QryStr & " AND " & Criteria
            End If
        End If
        If Len(DateFrom) > 0 And Len(DateTo) > 0 Then
            QryStr = QryStr & " WHERE DATE(" & DateField & ") "
            QryStr = QryStr & " BETWEEN " & Chr(39) & strDateFrom & Chr(39)
            QryStr = QryStr & " AND " & Chr(39) & strDateTo & Chr(39)
            If Len(Criteria) > 0 Then
                QryStr = QryStr & " AND " & Criteria
            End If
        End If
        If Len(DateFrom) = 0 And Len(DateTo) = 0 And Len(Criteria) > 0 Then
            QryStr = QryStr & " WHERE " & Criteria
        End If

        If Len(SortField) > 0 Then
            QryStr = QryStr & " ORDER BY " & MainSortField
            If Reversed = False Then
                QryStr = QryStr & " ASC"
            Else
                QryStr = QryStr & " DESC"
            End If
        End If

        CreateSearchSQL = QryStr

    End Function

    Function ReturnLastID(DBTable As String, connString As String, ByRef ErrMessage As String) As String
        Dim QryStr As String
        Dim cmd As MySqlCommand
        Dim con As MySqlConnection
        Dim WholeRow As Object()
        Dim Fieldnames As String()
        Dim IDX As Integer
        Dim numFields As Integer
        Dim dr1 As MySqlDataReader
        Dim myReader As MySqlDataReader
        Dim SearchFieldType As Type = Nothing
        Dim ReturnValue As String = ""

        ReturnLastID = ""
        If Len(DBTable) = 0 Then
            ErrMessage = "Error in ReturnLastID: No Table name passed"
            Exit Function
        End If
        If Len(connString) = 0 Then
            ErrMessage = "Error in ReturnLastID: No Connection String Passed "
            Exit Function
        End If
        Try
            con = New MySqlConnection(connString)
            con.Open()
            QryStr = "SELECT ID FROM " & DBTable & " ORDER BY ID desc"

            cmd = New MySqlCommand(QryStr, con)
            myReader = cmd.ExecuteReader()

            If myReader.HasRows Then
                'myReader.GetValues(WholeRow)
                'AllFieldValues = WholeRow
                ReturnLastID = myReader("ID").ToString
                While myReader.Read()
                    'DescriptionText.Text = dr("Description").ToString
                    'CostText.Text = dr("Cost").ToString
                    'PriceText.Text = dr("Price").ToString

                    'WHAT IF THE SEARCH TYPE is not STRING here ??????
                    'myReader.GetValues(WholeRow)
                    'AllFieldValues = WholeRow

                End While
            Else
                'No Rows
                ErrMessage = "Exception Error in ReturnLastID() : No Rows Returned"
            End If
            con.Close()
            myReader.Close()

        Catch ex As Exception
            ErrMessage = "Exception Error in ReturnLastID() : " & ex.ToString
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, ErrMessage, "ReturnLastID()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
        End Try

    End Function

    Function IsLocked(ConnString As String, TableName As String, LockStatusField As String, LockValue As String,
                      EditField As String, EditValue As Object, Optional EditFieldType As String = "STRING") As Boolean
        Dim LoadedOK As Boolean = False
        Dim SearchField As String
        Dim SearchText As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim RecID As String
        Dim SearchCriteria As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ErrMessage As String = ""

        'Search Lock Field in specified table to see if LOCK or similar value exists AND search the actual field and value
        'that will be reserved for the lock - eg user wants to edit Reference abc123 in field DeliveryReference
        'LockStatusField = LockStatus in TableName and LockValue = LOCKED.
        'While EditField = DeliveryReference and EditValue = abc123
        SearchField = LockStatusField
        SearchText = LockValue
        FieldType = "STRING"
        ReturnField = "ID"
        SortFields = ""
        RecID = ""
        SearchCriteria = ""
        If UCase(EditFieldType) = "STRING" Then
            SearchCriteria = EditField & " = " & Chr(34) & EditValue & Chr(34)
        ElseIf UCase(EditFieldType) = "INTEGER" Then
            SearchCriteria = EditField & " = " & EditValue
        Else
            '
        End If

        IsLocked = False

        LoadedOK = Find_myQuery(ConnString, TableName, SearchField, SearchText, FieldType, ReturnField, RecID,
                                         Nothing, Nothing, SearchCriteria, SortFields, Reversed, ErrMessage)
        If Len(RecID) > 0 Then
            IsLocked = True
        End If
    End Function

    Function InsertUpdateRecords_Using_Parameters(ByVal connString As String, ByVal UpdateRecords As Boolean, ByVal REcordID As String,
                    ByVal DBTable As String, ByVal Fieldnames As String, ByVal FieldValues As String,
                    Optional ByVal UpdateCriteria As String = "", Optional ByVal ExcludeFields As String = "", Optional ByRef ErrMessages As String = "",
                    Optional EncaseFields As Boolean = False, Optional FieldValueDelim As String = ",") As Boolean

        Dim cn As Object
        Dim rs As Object
        Dim cmd As MySqlCommand
        Dim con As MySqlConnection
        Dim myConn As String
        Dim provider As String
        'Dim cmd As oledbcommand
        'Dim da As oledbdataadapter
        Dim strSQL As String
        Dim strDeliveryRef As String
        Dim dtDeliveryDate As Date
        Dim strDeliveryDate As String
        Dim ExcludedFields As String
        Dim FieldTypeArr() As String
        Dim FieldType As String
        Dim ParamInfo As String
        Dim ParamLength As Long
        Dim Dic_FieldTypes As Object
        Dim Dic_ParamInfo As Object
        Dim ParamInfoArr() As String
        Dim FieldIDX As Integer
        Dim ParamArr() As Object
        Dim Param As Object
        Dim ParamKey As String
        Dim ParamName As String
        Dim ParamValue As Object
        Dim FieldLength As Integer
        Dim FieldValue As String
        Dim FieldTypes As String
        Dim ErrMessage As String
        Dim cmdResult As Integer
        Dim myMessage As String
        Dim Fieldname As String
        Dim DontSave As Boolean = False
        Dim myGUID As Guid

        InsertUpdateRecords_Using_Parameters = False

        'myConn = "G:\MIS\Goodsin Timesheet Data\GoodsInTimesheetRecords.accdb"

        If Len(DBTable) > 0 Then
            'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
            'strSQL = strSQL & " VALUES (""" & Fieldvalues & """);"
        Else
            MsgBox("DATABASE TABLE NOT SPECIFIED")
            'cn.Close
            'cn = Nothing
            Exit Function
        End If

        'Set cn = ADODB.Connection
        Dic_ParamInfo = CreateObject("Scripting.Dictionary")
        Dic_ParamInfo.comparemode = vbTextCompare
        Dic_ParamInfo.RemoveAll
        Dic_FieldTypes = CreateObject("Scripting.Dictionary")
        Dic_FieldTypes.comparemode = vbTextCompare
        Dic_FieldTypes.RemoveAll
        'Set TargetRange = ThisWorkbook.Sheets("Prefs").Range("A10")
        con = New MySqlConnection(connString)


        'Set cn = CreateObject("ADODB.Connection")

        ExcludedFields = ExcludeFields

        'Set Dic = CreateObject("Scripting.Dictionary")
        ReDim FieldTypeArr(1)
        ReDim ParamArr(1)
        ErrMessage = ""
        FieldTypes = ""
        Call GetMyFields(DBTable, connString, ErrMessage, "", FieldTypes, 0, Dic_FieldTypes)

        'SHOULD BE FieldTypes() as OBJECT - 03-SEP-2018
        'GETTING AUTHENTICATION ERROR when the above function is run. seems ok now.

        'FieldValueDelim is passed into this procedure - it is semi-colon by default.

        'NO WHERE CLAUSE !

        If UpdateRecords = True Then
            strSQL = PrepareUpdate_With_Parameters(REcordID, connString, DBTable, Fieldnames, FieldValues, Dic_FieldTypes, Dic_ParamInfo, UpdateCriteria, ErrMessage, ExcludedFields, False, ",", FieldValueDelim)
        Else
            strSQL = PrepareInsert_With_Parameters(REcordID, connString, DBTable, Fieldnames, FieldValues, Dic_FieldTypes, Dic_ParamInfo, ErrMessage, ExcludedFields, False, ",", FieldValueDelim)
        End If
        'strSQL = "UPDATE " & DBTable & " SET [DeliveryDate] = ?"
        'strDeliveryRef = "TEST-DAN-YES"
        'Fieldnames = "DeliveryDate,DeliveryReference"
        'dtDeliveryDate = CDate("21/05/2018")
        'strDeliveryDate = Format(dtDeliveryDate, "dd/MM/yyyy")
        'FieldValues = Chr(34) & strDeliveryDate & Chr(34) & "," & Chr(34) & strDeliveryRef & Chr(34)
        'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
        'strSQL = strSQL & " VALUES (" & FieldValues & ")"
        'MsgBox ("SQL= " & strSQL)
        'cmd.Parameters.Clear
        FieldIDX = 1
        Try
            con.Open()
            cmd = New MySqlCommand(strSQL, con)
            cmd.Connection = con
            cmd.CommandText = strSQL
            cmd.Prepare()
            For Each ParamKey In Dic_ParamInfo
                'ParamKey = "@" & fieldname
                Param = Nothing
                If Len(ParamKey) = 0 Then
                    Continue For
                End If
                'where does dic_fieldtypes come in ?????????????
                ParamName = ParamKey
                ParamInfo = Dic_ParamInfo(ParamKey)
                Fieldname = Mid(ParamKey, 2, Len(ParamKey))
                If InStr(ParamInfo, "_") > 0 Then
                    ParamInfoArr = Split(ParamInfo, "_")
                    ParamLength = Len(ParamInfoArr(2))
                    FieldType = ParamInfoArr(0)
                    ParamValue = ParamInfoArr(2)

                    If IsNumeric(ParamInfoArr(1)) Then
                        FieldLength = CLng(ParamInfoArr(1))
                        If ParamInfoArr(0) = "VARCHAR" Then
                            If FieldLength = 0 Then
                                FieldLength = 1
                            End If
                        End If
                    Else
                        'bugger
                        FieldLength = CInt(ParamLength)
                    End If
                    If Len(ParamInfoArr(2)) > 0 Then
                        'cmd.Parameters.Add("@Baz", SqlDbType.VarChar, 50).Value = Baz
                        'Return cmd.ExecuteScalar().ToString()
                        'End Using
                        If UCase(FieldType) = "VARCHAR" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.VarChar, FieldLength).Value = ParamValue
                        ElseIf UCase(FieldType) = "INT" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.Int32, FieldLength).Value = CInt(ParamValue)
                        ElseIf UCase(FieldType) = "DOUBLE" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.Double, FieldLength).Value = CDbl(ParamValue)
                        ElseIf UCase(FieldType) = "DATE" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.Date, FieldLength).Value = CDate(ParamValue)
                        ElseIf UCase(FieldType) = "DATETIME" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.DateTime, FieldLength).Value = CDate(ParamValue)
                        ElseIf UCase(FieldType) = "TEXT" Then
                            cmd.Parameters.Add(ParamName, MySqlDbType.Text, FieldLength).Value = ParamValue
                        ElseIf UCase(FieldType) = "GUID" Then
                            myGUID = DirectCast(ParamValue, System.Guid)
                            cmd.Parameters.Add(ParamName, MySqlDbType.Guid, FieldLength).Value = myGUID
                        Else
                            myMessage = "Field Type Unknown: " & FieldType
                            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "InsertUpdateRecords_using_Parameters()",
                                        frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
                            DontSave = True
                            Continue For
                        End If

                    End If
                    'TREAT DATES as STRINGS before Insertion?:
                End If
                FieldIDX = FieldIDX + 1
            Next
            If DontSave = False Then

                cmdResult = cmd.ExecuteNonQuery
            End If
        Catch ex As Exception
            myMessage = "Exception Error in InsertUpdateRecords_Using_Parameters Record: " & ex.ToString
            InsertUpdateRecords_Using_Parameters = False
            frmMainGIForm.logger.LogError("GI_Error_v1_1.log", Application.StartupPath, myMessage, "InsertUpdateRecords_using_Parameters()",
                                        frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() &
                                        ", OPERATOR Logged in:" & frmMainGIForm.myUsername)
            Exit Function

        End Try
        con.Close()
        cmd = Nothing
        con = Nothing
        Dic_FieldTypes = Nothing
        Dic_ParamInfo = Nothing
        'Set ParamArr = Nothing
        InsertUpdateRecords_Using_Parameters = True


    End Function

    Function PrepareParameters(ByVal REcordID As String, ByRef InsertSQL As String, ByRef UpdateSQL As String, ByRef Dic_ParameterInfo As Object,
    ByVal DBTable As String, ByRef ParamFieldnames As String, ByRef ParamValues As String, dic_FieldTypes As Object,
    ByVal Dic_FieldsAndValues As Object, Optional Encase_Fields As Boolean = False, Optional UpdateCriteria As String = "", Optional ForAccess As Boolean = False) As Long
        'Dim FieldArr() As String
        Dim ValueArr As Object
        Dim Fieldname As String
        Dim fldValue As Object
        Dim ParamKey As String
        Dim ParamItem As String
        Dim TheFieldType As String
        Dim FieldTypesArr() As String
        Dim varKey As String
        Dim fldLength As Long
        Dim ParamFldLength As Long
        Dim DontSave As Boolean
        Dim IncludeComma As Boolean
        Dim IncludeQuotes As Boolean
        Dim IncludeSpaces As Boolean
        Dim IncludeSingleQuote As Boolean
        Dim DeliveryComments As String
        Dim FieldCount As Long
        Dim AllFieldnames As String
        Dim AllFieldTypes As String
        Dim dtDateTime As DateTime

        'THIS FUNCTION ONLY SAVES 1 RECORD.

        PrepareParameters = 0
        If Dic_FieldsAndValues.Count = 0 Then
            Exit Function
        End If
        DontSave = False
        'ReDim FieldArr(1)
        ReDim FieldTypesArr(1)

        fldLength = 0
        TheFieldType = ""
        FieldCount = 0

        DeliveryComments = "ISSUES: "
        ParamFldLength = 0
        ParamFieldnames = ""
        ParamValues = ""

        'Dic_FieldTypes = New Scripting.Dictionary
        Dic_ParameterInfo = CreateObject("Scripting.Dictionary")
        Dic_ParameterInfo.comparemode = vbTextCompare
        'Dic_FieldTypes(fieldname) = FieldType()
        InsertSQL = ""
        UpdateSQL = "UPDATE " & DBTable & " SET "
        For Each varKey In Dic_FieldsAndValues
            FieldCount = FieldCount + 1
            If varKey Is Nothing Or Len(varKey) = 0 Then
                Continue For
            End If
            'VarKey = UCase(VarKey)

            Fieldname = UCase(varKey)
            fldValue = Dic_FieldsAndValues(varKey)
            TheFieldType = dic_FieldTypes(Fieldname)
            DontSave = False 'for EACH field.
            fldLength = Len(fldValue)
            ParamFldLength = 0
            If TheFieldType = "VARCHAR" Then '202
                'SHORT TEXT
                ParamFldLength = Len(fldValue)
                If fldLength > 0 Then
                    If Encase_Fields Then
                        fldValue = Chr(34) & fldValue & Chr(34)
                    End If
                    IncludeComma = True
                    IncludeQuotes = True
                    IncludeSpaces = False
                    fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                Else
                    'BLANK / EMPTY field value:
                    'dontsave = True ??????????????????????????????? what if user wants to blank out a field ??
                    fldValue = " "
                    ParamFldLength = 1
                    DeliveryComments = DeliveryComments & " Field: " & Fieldname & " =EMPTY,"
                End If
            ElseIf TheFieldType = "INT" Then '3
                'INTEGER or LONG - NUMERIC.
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = True
                If fldLength > 0 Then
                    fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                    If IsNumeric(fldValue) Then
                        'OK SAVE PARAMETER
                        fldValue = CInt(fldValue)
                    Else
                        'TEXT appears in this numeric field:
                        DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": " & CStr(fldValue) & ","
                        fldValue = 0

                    End If
                Else
                    DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": EMPTY,"
                    fldValue = 0
                End If
            ElseIf TheFieldType = "DOUBLE" Then '5
                'DOUBLE - will have PRECISION and NUMBER of DECIMAL PLACES
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = True
                If fldLength > 0 Then
                    fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                    If IsNumeric(fldValue) Then
                        'OK SAVE PARAMETER
                        fldValue = CDbl(fldValue)
                    Else
                        'TEXT appears in this numeric field:
                        DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": " & CStr(fldValue) & ","
                        fldValue = 0

                    End If
                Else
                    DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": EMPTY,"
                    fldValue = 0 'Could put 0.00f instead ????
                End If

                'Convert fldValue (OBJECT TYPE) to CURRENCY TYPE
            ElseIf TheFieldType = "CURRENCY" Then '6
                'CURRENCY
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = True
                If fldLength > 0 Then
                    fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                    If IsNumeric(fldValue) Then
                        'OK SAVE PARAMETER
                        fldValue = CDbl(fldValue)
                    Else
                        'TEXT appears in this numeric field:
                        DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": " & CStr(fldValue) & ","
                        fldValue = 0

                    End If
                Else
                    DeliveryComments = DeliveryComments & " Field: " & Fieldname & ": EMPTY,"
                    fldValue = 0 'Could put 0.00f instead ????
                End If

            ElseIf TheFieldType = "DATETIME" Then '7
                'DATEs
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = False
                IncludeSpaces = False
                IncludeSingleQuote = False
                If fldLength < 8 Then
                    fldValue = "01/01/1970"
                    DeliveryComments = DeliveryComments & " Field:" & Fieldname & " Invalid Date,"
                Else
                    fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces, IncludeSingleQuote)
                    If IsDate(fldValue) Then
                        dtDateTime = CDate(fldValue)
                        fldValue = dtDateTime
                    Else
                        DeliveryComments = DeliveryComments & " Field:" & Fieldname & " Invalid Date,"
                    End If
                End If

            ElseIf TheFieldType = "" Then
                'Field Type is Unrecognised or some error here ? = 0
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = True
                If fldLength < 2 Then
                    DontSave = True
                Else
                    DontSave = True
                End If
                DeliveryComments = DeliveryComments & " Field:" & Fieldname & ": Type0= " & fldValue
            Else
                'UNKNOWN FIELD TYPE:
                ParamFldLength = 0
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = True
                If fldLength < 2 Then
                    DontSave = True
                Else
                    DontSave = True
                End If
                DeliveryComments = DeliveryComments & " Field:" & Fieldname & ": Unknown Type= " & fldValue

            End If


            If DontSave = False Then
                'Create Parameter for this field:
                ParamKey = "@" & Fieldname
                If FieldCount < 2 Then
                    ParamFieldnames = Fieldname
                    ParamValues = ParamKey
                    If ForAccess Then
                        UpdateSQL = UpdateSQL & "[" & Fieldname & "]" & " = " & ParamKey
                    Else
                        UpdateSQL = UpdateSQL & Fieldname & " = " & ParamKey
                    End If
                Else
                    ParamFieldnames = ParamFieldnames & "," & Fieldname
                    ParamValues = ParamValues & "," & ParamKey
                    If ForAccess Then
                        UpdateSQL = UpdateSQL & ",[" & Fieldname & "]" & " = " & ParamKey
                    Else
                        UpdateSQL = UpdateSQL & "," & Fieldname & " = " & ParamKey
                    End If
                End If
                ParamItem = TheFieldType & "_" & CStr(ParamFldLength) & "_" & fldValue & "_" & DeliveryComments
                If Not Dic_ParameterInfo.Exists(ParamKey) Then
                    Dic_ParameterInfo(ParamKey) = ParamItem
                End If

            Else
                'DONT CREATE PARAMETER - INVALID value:
                'BUT still save / update the DeliveryComments:
                DeliveryComments = DeliveryComments & " Field:" & Fieldname & ": Not Saved= " & fldValue
            End If
        Next
        If Len(REcordID) > 0 Then
            If CLng(REcordID) > 0 Then
                UpdateSQL = UpdateSQL & " WHERE ID = " & CLng(REcordID)
            End If
        Else
            If Len(UpdateCriteria) > 0 Then
                UpdateSQL = UpdateSQL & " WHERE " & UpdateCriteria
            End If
        End If
        InsertSQL = "INSERT INTO " & DBTable & " (" & ParamFieldnames & ")" & " VALUES " & "(" & ParamValues & ")"
        PrepareParameters = Dic_ParameterInfo.Count


    End Function

    Function PrepareUpdate_With_Parameters(ByRef REcordID As String, ByVal ConnString As String, ByVal DBTable As String, ByRef Fieldnames As String,
                                           ByVal FieldValues As String, ByRef dic_FieldTypes As Object, ByRef Dic_ParameterInfo As Object, ByVal UpdateCriteria As String,
                                           Optional ErrMessage As String = "",
                                           Optional ByRef ExcludeFields As String = "", Optional Encase_Fields As Boolean = False,
                                           Optional FieldDelim As String = ",", Optional ValueDelim As String = ";",
                                           Optional TypeStartIDX As Long = 0,
                                           Optional RemoveBadChars As Boolean = False,
                                           Optional IncludeCommaInBadChars As Boolean = False,
                                           Optional IncludeSpeechMarksInBadChars As Boolean = False,
                                           Optional REplaceWith As String = "_") As String
        'PrepareUpdate_With_Parameters(DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, UpdateCriteria, ExcludedFields, False, FieldValueDelim, FieldValueDelim)
        Dim FieldNameArray() As String
        Dim IgnoreFieldsArray() As String
        Dim ValueArray() As String
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim fldValue As String
        Dim NumFields As Integer
        Dim UpdateCmd As String
        Dim IncludeComma As Boolean
        Dim Operation As String
        Dim ParamKey As String
        Dim ParamItem As Object
        Dim FieldType As Integer
        Dim FieldTypes As String
        Dim FieldTypesArr() As String
        Dim varKey As String
        Dim Dic_FieldsAndValues As Object
        Dim InsertSQL As String
        Dim UpdateSQL As String
        Dim ParamFieldnames As String
        Dim ParamValues As String
        Dim TotalParameters As Long
        Dim AllFieldnames As String

        FinalCMD = ""
        ReDim FieldNameArray(1)
        ReDim IgnoreFieldsArray(1)
        ReDim ValueArray(1)

        Dic_ParameterInfo = CreateObject("Scripting.Dictionary")
        Dic_ParameterInfo.RemoveAll
        Dic_ParameterInfo.CompareMode = vbTextCompare
        Dic_FieldsAndValues = CreateObject("Scripting.Dictionary")
        Dic_FieldsAndValues.RemoveAll
        Dic_FieldsAndValues.CompareMode = vbTextCompare

        FieldTypes = ""
        AllFieldnames = GetMyFields(DBTable, ConnString, ErrMessage, ExcludeFields, FieldTypes, TypeStartIDX)
        FieldTypesArr = strToStringArray(FieldTypes, ",", TypeStartIDX, Encase_Fields, RemoveBadChars, IncludeCommaInBadChars, REplaceWith, IncludeSpeechMarksInBadChars)


        PrepareUpdate_With_Parameters = ""
        If Len(ExcludeFields) > 0 Then
            IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
        End If
        If Len(DBTable) = 0 Then
            MsgBox("Error in PrepareUpdate_With_Parameters: No Database Table specified")
            PrepareUpdate_With_Parameters = ""
            Exit Function
        End If
        If Len(Fieldnames) = 0 Then
            'NumFields = GetNumFields(connString, "SELECT * FROM " & TableName, DBName, Fieldnames)
            Fieldnames = AllFieldnames
        End If

        FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True, False, "_", False)
        If Len(FieldValues) > 0 Then
            ValueArray = strToStringArray(FieldValues, ValueDelim, 0, False, True, IncludeComma, "_", IncludeSpeechMarksInBadChars)
        Else
            MsgBox("Error in PrepareUpdate_With_Parameters: No values specified")
            PrepareUpdate_With_Parameters = ""
            Exit Function
        End If
        'BUT check that the values are removed too if they corresponded with those fields removed ????
        Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray, 0, 0) 'rebuilds whole list without the extracted fields

        If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
            If UBound(FieldNameArray) < UBound(ValueArray) Then
                MsgBox("Error in PrepareUpdate_With_Parameters: Number of Fields Passed are LESS than Number of VALUES passed.")
                PrepareUpdate_With_Parameters = ""
                Exit Function
            End If
            If UBound(FieldNameArray) > UBound(ValueArray) Then
                MsgBox("Error in PrepareUpdate_With_Parameters: Number of Fields Passed are GREATER than Number of VALUES passed.")
                PrepareUpdate_With_Parameters = ""
                Exit Function
            End If
        End If
        'Prepare and insert parameter INFO for the script dictionary


        'IDX = 1
        'UpdateCmd = "UPDATE " & DBTable & " SET "

        Dic_FieldsAndValues.RemoveAll
        For IDX = 0 To UBound(FieldNameArray)
            fldName = UCase(FieldNameArray(IDX))

            If Not Dic_FieldsAndValues.Exists(fldName) Then
                Dic_FieldsAndValues(fldName) = ValueArray(IDX)
            End If
        Next
        'Prepare and insert parameter INFO for the script dictionary

        TotalParameters = PrepareParameters(REcordID, InsertSQL, UpdateSQL, Dic_ParameterInfo, DBTable, ParamFieldnames, ParamValues, Dic_FieldTypes, Dic_FieldsAndValues, False)

        FinalCMD = UpdateSQL

        If Len(UpdateCriteria) > 0 Then
            FinalCMD = FinalCMD & " WHERE " & UpdateCriteria
        End If

        PrepareUpdate_With_Parameters = FinalCMD


    End Function

    Function PrepareInsert_With_Parameters(ByVal REcordID As String, ByVal connString As String, ByVal DBTable As String, ByRef Fieldnames As String,
                                           ByVal FieldValues As String, dic_FieldTypes As Object, ByRef Dic_ParameterInfo As Object,
                                           Optional ErrMessage As String = "",
                                           Optional ByRef ExcludeFields As String = "",
                                           Optional Encase_Fields As Boolean = False,
                                           Optional FieldDelim As String = ",",
                                           Optional ValueDelim As String = ",",
                                           Optional TypeStartIDX As Long = 0,
                                           Optional RemoveBadChars As Boolean = False,
                                           Optional IncludeCommaInBadChars As Boolean = False,
                                           Optional IncludeSpeechMarksInBadChars As Boolean = False,
                                           Optional REplaceWith As String = "_") As String
        'PrepareInsert_With_Parameters(DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, ExcludedFields, False, FieldValueDelim, FieldValueDelim)
        Dim FieldNameArray() As String
        Dim IgnoreFieldsArray() As String
        Dim ValueArray() As String
        Dim FinalCMD As String
        Dim IDX As Integer
        Dim fldName As String
        Dim fldValue As String
        Dim fldValues As String
        Dim NumFields As Integer
        Dim UpdateCmd As String
        Dim IncludeComma As Boolean
        Dim Operation As String
        Dim ParamKey As String
        Dim ParamItem As String
        Dim FieldType As Integer
        Dim FieldTypes As String
        Dim FieldTypesArr() As String
        Dim varKey As String
        Dim fldLength As Long
        Dim TotalParameters As Long
        Dim InsertSQL As String
        Dim UpdateSQL As String
        Dim ParamFieldnames As String
        Dim ParamValues As String
        Dim Dic_FieldsAndValues As Object
        Dim AllFieldnames As String

        FinalCMD = ""
        fldValues = ""
        ReDim FieldNameArray(1)
        ReDim IgnoreFieldsArray(1)
        ReDim ValueArray(1)


        Dic_ParameterInfo = CreateObject("Scripting.Dictionary")
        Dic_ParameterInfo.CompareMode = vbTextCompare
        Dic_FieldsAndValues = New Scripting.Dictionary
        Dic_FieldsAndValues.RemoveAll
        Dic_FieldsAndValues.CompareMode = vbTextCompare
        'Dic_ParameterInfo.RemoveAll

        FieldTypes = ""
        AllFieldnames = GetMyFields(DBTable, connString, ErrMessage, ExcludeFields, FieldTypes, TypeStartIDX)
        FieldTypesArr = strToStringArray(FieldTypes, ",", TypeStartIDX, Encase_Fields, RemoveBadChars, IncludeCommaInBadChars, REplaceWith, IncludeSpeechMarksInBadChars)

        PrepareInsert_With_Parameters = ""

        If Len(ExcludeFields) > 0 Then
            IgnoreFieldsArray = strToStringArray(ExcludeFields, FieldDelim, 0, False, True, False, "_", False)
        End If
        If Len(DBTable) = 0 Then
            MsgBox("Error in PrepareInsert_With_Parameters: No Database Table specified")
            PrepareInsert_With_Parameters = ""
            Exit Function
        End If

        If Len(Fieldnames) = 0 Then
            Fieldnames = AllFieldnames
        End If
        'SO what if Dic_ParameterInfo.count > 0 then ????????????
        'This means that its been passed back in - after evaluation from the parameters section - before the EXECUTE SQL.
        ' - means that number of fields do not match number of parameter @ values passed.
        ' so we need to resolve this. maybe instead of @P1 .... we need to use the actual fieldname - @DeliveryDate etc.
        ' This way we can remove the @ symbol and know which fields have been passed and which ARE MISSING !
        If Dic_ParameterInfo.Count > 0 Then
            'OK so we can get all of the fields from the table to start with.
            'Then we need to eliminate the fields that are NOT in the Dictionary:
            ' - this will form the fields that are left to put into the final SQL query.

        End If
        '
        'FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True) 'EACH fieldname passed must have square brackets around it.
        FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True, False, "_", False)
        If Len(FieldValues) > 0 Then
            ValueArray = strToStringArray(FieldValues, ValueDelim, 0, False, True, IncludeComma, "", IncludeSpeechMarksInBadChars)
        Else
            MsgBox("Error in PrepareInsert_With_Parameters: No values specified")
            PrepareInsert_With_Parameters = ""
            Exit Function
        End If
        Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, FieldDelim, FieldNameArray, 0, 0)
        If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
            If UBound(FieldNameArray) < UBound(ValueArray) Then
                MsgBox("Error in PrepareInsert_With_Parameters: Number of Fields Passed are LESS than Number of VALUES passed.")
                PrepareInsert_With_Parameters = ""
                Exit Function
            End If
            If UBound(FieldNameArray) > UBound(ValueArray) Then
                MsgBox("Error in PrepareInsert_With_Parameters: Number of Fields Passed are GREATER than Number of VALUES passed.")
                PrepareInsert_With_Parameters = ""
                Exit Function
            End If
        End If
        Dic_FieldsAndValues.RemoveAll
        For IDX = 0 To UBound(FieldNameArray)
            fldName = UCase(FieldNameArray(IDX))
            If Not Dic_FieldsAndValues.Exists(fldName) Then
                Dic_FieldsAndValues(fldName) = ValueArray(IDX)
            End If
        Next
        'Prepare and insert parameter INFO for the script dictionary

        TotalParameters = PrepareParameters(REcordID, InsertSQL, UpdateSQL, Dic_ParameterInfo, DBTable, ParamFieldnames, ParamValues, dic_FieldTypes, Dic_FieldsAndValues, False)

        FinalCMD = InsertSQL
        PrepareInsert_With_Parameters = FinalCMD

    End Function

    Sub Create_Dic_FieldTypes(ByRef Dic_FieldTypes As Object, TableName As String, Fieldnames As String, FieldTypes As String)
        Dim idx As Long
        Dim Fieldtype As String
        Dim Fieldname As String
        Dim FieldArr() As String
        Dim TypeArr() As String

        FieldArr = strToStringArray(Fieldnames, ",", 0, False)
        TypeArr = strToStringArray(FieldTypes, ",", 0, False)
        For idx = 0 To UBound(FieldArr)
            Fieldname = FieldArr(idx)
            Fieldtype = TypeArr(idx)
            If Fieldname Is Nothing Then
                Continue For
            End If
            If Not Dic_FieldTypes.exists(TableName & "_" & Fieldname) Then
                Dic_FieldTypes(Fieldname) = Fieldtype
            End If
        Next

    End Sub

End Module
