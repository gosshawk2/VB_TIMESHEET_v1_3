Public Class frmMainGIForm
    Private _SelectedImportDate As String
    Private _SelectedDeliveryDate As String
    Private _SelectedFirstImportDate As String
    Private _SelectedLastImportDate As String
    Public myVersion As String = "v1.3"
    Public logger = New clsErrorLog
    Public DatabaseType As String = "MYSQL" 'could also be MYSQL or ORACLE
    Public DBName As String = "AssetRegister.accdb" 'could also be just AssetRegister for Mysql databases.
    Public DBProvider As String = "Provider=Microsoft.ACE.OLEDB.12.0"
    Public myDocsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Public DataDir As String = "Data Source=|MyDocumentsDirectory|\" ' change to the documents directory. will work on any machine.
    Public dbsource As String = "Data Source=" & myDocsPath & "\"
    Public PublicDir1 As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
    Public ProgramData As String = Application.CommonAppDataPath 'C:\ProgramData\ECP\Asset_RegisterVB_v2_52\2.0.5.2\
    Public myConnString As String = ""
    Public myLOCALConnString As String = ""
    Public MyLocalMessage As String = ""
    Public myComputerName As String = ""
    Public myPort As String = "3306"
    Public myUsername As String = ""
    Public myUsernameSuffix As String = ""
    Public myUserID As String = ""
    Public myPassword As String = ""
    Public myServer As String = "localhost"
    Public myAccessRights As String = ""
    Public myDatabase As String = "assetregister"
    Public myEmpNo As String = ""
    Public myFirstname As String = ""
    Public myLastname As String = ""
    Public myScreenWidth As Integer = 0
    Public myScreenHeight As Integer = 0
    Public myLoggedIn As Boolean = False
    Public UpdatedOK As Boolean = False
    Public DebugMode As Boolean = False
    Public SYSTEMUsername As String
    Public SYSTEMPassword As String
    Public Initialised As Boolean
    Public EXITProcess As Boolean
    Public mySessionID As String
    Public ListSelectedItem As String
    Public WorkbookPath As String
    'Public TotalOperatives As Long
    Public zoomcontrolvalue As Integer
    Public TotalOpHours As Double
    Public TotalFLMHours As Double
    Public ErrList() As String
    'Public Dic_Controls As New Scripting.Dictionary
    Public Shared Dic_ShortCount As New Scripting.Dictionary
    Public Shared Dic_HighestShortTAGID As New Scripting.Dictionary
    Public Shared Dic_ExtraCount As New Scripting.Dictionary
    Public Shared Dic_HighestExtraTAGID As New Scripting.Dictionary
    Public Shared Dic_HighestOpTAGID As New Scripting.Dictionary
    Public Shared Dic_TotalOpHours As New Scripting.Dictionary
    Public Shared Dic_TotalFLMHours As New Scripting.Dictionary
    Public OpTimeTAGID As Long
    Public Shared Dic_HighestOpBtnTAGID As New Scripting.Dictionary
    Public Dic_TotalOperatives As New Scripting.Dictionary
    Public MonthlyDateSelected As DateTime
    Public DEBUG_MODE As Boolean = False
    Public MainWorkBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
    Public SelectedSheet As String
    Public SelectedWorkbookFile As String
    Public btnUpdateEmployeesColor As String
    Public btnImportDataColor As String
    Public myName As String
    'Public dic_Totals As Object
    Public OP_ID As Long
    Public ControlPanelIdx As String
    Public SelectedDeliveryRef As String
    Public SelectedASN As String
    Public ControlPanelFormName As String = "GI_TIMESHEET"


    Public Property GetImportDate() As String
        Get
            Return _SelectedImportDate
        End Get
        Set(value As String)
            _SelectedImportDate = value
        End Set
    End Property

    '_SelectedDeliveryDate

    Public Property GetDeliveryDate() As String
        Get
            Return _SelectedDeliveryDate
        End Get
        Set(value As String)
            _SelectedDeliveryDate = value
        End Set
    End Property

    Public Property GetFirstImportDate() As String
        Get
            Return _SelectedFirstImportDate
        End Get
        Set(value As String)
            _SelectedFirstImportDate = value
        End Set
    End Property

    Public Property GetLastImportDate() As String
        Get
            Return _SelectedLastImportDate
        End Get
        Set(value As String)
            _SelectedLastImportDate = value
        End Set
    End Property

    Private Sub DisplayAvailability(ByVal available As Boolean)
        If available Then
            Me.txtConnected.BackColor = Color.Lime
        Else
            Me.txtConnected.BackColor = Color.Red
        End If
    End Sub

    Private Sub MyComputerNetwork_NetworkAvailabilityChanged(
    ByVal sender As Object,
    ByVal e As Devices.NetworkAvailableEventArgs)

        DisplayAvailability(e.IsNetworkAvailable)
    End Sub

    Private Sub Handle_NetworkAvailabilityChanged()
        AddHandler My.Computer.Network.NetworkAvailabilityChanged,
       AddressOf MyComputerNetwork_NetworkAvailabilityChanged
        DisplayAvailability(My.Computer.Network.IsAvailable)
    End Sub

    Private Sub ShowMainControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowMainControlToolStripMenuItem.Click
        'Show Main Goods In Control Form:
        'MsgBox("Show Main Control Form")
        'CreateGIControlPanel()
        'frmGI_Userform.Show()
        Dim cf As New frmGI_RP_Userform

        ControlPanelIdx = Get_GUID()
        FormID = ControlPanelIdx
        CPFormName = "GI_TIMESHEET_"

        cf.MdiParent = Me
        cf.StartPosition = FormStartPosition.Manual
        cf.Text = "Goods In Control Panel_" & Now().ToString("dd/MM/yyyy HH:mm:ss")
        cf.Name = ControlPanelFormName & ControlPanelIdx
        cf.Tag = ControlPanelIdx
        cf.Show()
        'cf.ShowDialog() - gives error
        Clear_Entry_Controls(ControlPanelFormName & ControlPanelIdx, 1, 40)
        Clear_Entry_Controls(ControlPanelFormName & ControlPanelIdx, 801, 807)



    End Sub

    Public Sub getallforms(ByVal sender As Object)
        Dim Forms As New List(Of Form)()
        Dim formType As Type = Type.GetType("System.Windows.Forms.Form")
        For Each t As Type In sender.GetType().Assembly.GetTypes()
            If UCase(t.BaseType.ToString) = "SYSTEM.WINDOWS.FORMS.FORM" Then
                MsgBox(t.Name)
            End If
        Next
    End Sub

    Sub getAssemblyForms()
        Dim myAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

        Dim types As Type() = myAssembly.GetTypes()
        For Each t As Type In types
            If UCase(t.BaseType.ToString) = "SYSTEM.WINDOWS.FORMS.FORM" Then
                MessageBox.Show(t.Name)
            End If
        Next
    End Sub


    Sub ShowForms(FormName As String, Optional Dropdowns() As Object = Nothing, Optional OwnerForm As Form = Nothing)

        If UCase(FormName) = UCase("frmSelectSheet") Then
            Dim newform As New frmSelectSheet

            If OwnerForm Is Nothing Then
                newform.MdiParent = Me
            Else
                newform.MdiParent = OwnerForm
            End If
            newform.StartPosition = FormStartPosition.CenterParent
            newform.Text = "SELECT WORKSHEET"
            newform.Name = "SelectWorksheet"
            'NewForm.DropdownItems.AddRange(ExcelSheets) 'Getting Object not referenced error here.
            If Dropdowns IsNot Nothing Then
                For Each SheetName In Dropdowns
                    If SheetName IsNot Nothing Then
                        newform.DropdownItems.Add(SheetName)
                    End If
                Next
            End If
            newform.Visible = True
            newform.txtWorkbook.Text = ""
            newform.Show()

        End If

        If UCase(FormName) = UCase(ControlPanelFormName & ControlPanelIdx) Then
            Dim newform As New frmGI_RP_Userform

            If OwnerForm Is Nothing Then
                newform.MdiParent = Me
            Else
                newform.MdiParent = OwnerForm
            End If

        End If

    End Sub

    Sub InsertValueIntoForm(FormName As String, ControlName As String, value As String,
                            Optional ByRef comArray As Object = Nothing,
                            Optional NewFormTitle As String = "")
        Dim ctrl As Object
        Dim ctrls() As Control
        Dim comboCtrl As ComboBox
        Dim comboArr() As String
        Dim IDX As Long
        Dim FoundForm As Boolean = False
        Dim ErrMessage As String

        ReDim ctrls(1)
        For Each frm As Form In Application.OpenForms
            If UCase(frm.Name) = UCase(FormName) Then
                FoundForm = True
                If Len(NewFormTitle) > 0 Then
                    'frm.Text = NewFormTitle
                    Application.OpenForms.Item(FormName).Text = NewFormTitle
                End If
                ctrls = frm.Controls.Find(ControlName, True)
                Exit For
            End If
        Next
        If ctrls IsNot Nothing And UBound(ctrls) > -1 Then
            If ctrls(0) Is Nothing Then
                If FoundForm = False Then
                    ErrMessage = "Error: Cannot Find Form Passed"
                    logger.LogError("GI_Error_" & myVersion & ".log", Application.StartupPath, ErrMessage, "InsertValueIntoForm()",
                                    GetComputerName() & "," & GetIPv4Address() & "," & GetIPv6Address() & ", OPERATOR Logged in:" & myUsername,
                                    myUsernameSuffix)
                    Exit Sub
                Else
                    ErrMessage = "Error: Cannot Find CONTROL Passed"
                    logger.LogError("GI_Error_" & myVersion & ".log", Application.StartupPath, ErrMessage, "InsertValueIntoForm()",
                                    GetComputerName() & "," & GetIPv4Address() & "," & GetIPv6Address() & ", OPERATOR Logged in:" & myUsername,
                                    myUsernameSuffix)
                    Exit Sub
                End If
            End If
            ctrl = ctrls(0)
            If InStr(ctrl.Name, "com") > 0 Then
                'the control is a combobox. Needs to be filled:
                If Not comArray Is Nothing Then
                    comboCtrl = CType(ctrl, System.Windows.Forms.ComboBox)
                    'comboCtrl = CType(ctrl, comarray)
                    comboCtrl.Items.Clear()

                    For IDX = 0 To UBound(comArray)
                        If comArray(IDX) IsNot Nothing Then
                            comboCtrl.Items.Add(comArray(IDX))
                        End If
                    Next

                Else

                End If
                ctrl.text = value
            ElseIf InStr(ctrl.name, "btn") > 0 Then

            Else
                ctrl.Text = value
            End If
            Application.DoEvents()
        Else
            MsgBox("Could not find Control: " & ControlName)
        End If
    End Sub

    Function FindFrameControls(FormName As String, ControlName As String, Optional ByVal TAGNumber As String = "") As Control
        Dim ctrl As Control
        Dim ctrls() As Control
        Dim comboCtrl As ComboBox
        Dim comboArr() As String
        Dim IDX As Long

        ReDim ctrls(1)
        ctrl = Nothing
        For Each frm As Form In Application.OpenForms
            If UCase(frm.Name) = UCase(FormName) Then
                If Len(TAGNumber) > 0 Then
                    ControlName = FindControls(FormName, "", TAGNumber).Name
                End If
                If Len(ControlName) > 0 Then
                    ctrls = frm.Controls.Find(ControlName, True)
                    Exit For
                End If
            End If
        Next

        If ctrls IsNot Nothing And UBound(ctrls) > -1 Then
            ctrl = ctrls(0)
            Return ctrl
        End If
        Return ctrl
    End Function

    Function FindControls(Formname As String, ControlName As String, TagNumber As String, Optional ByRef childCtrl As Control = Nothing) As Control
        Dim ctrl As Control
        Dim FinalCtrl As Control
        Dim ctrls() As Control
        Dim comboCtrl As ComboBox
        Dim comboArr() As String
        Dim IDX As Long
        Dim FoundForm As Form = Nothing

        ReDim ctrls(1)
        FinalCtrl = Nothing
        'ctrl = Nothing
        For Each frm As Form In Application.OpenForms
            If UCase(frm.Name) = UCase(Formname) Then
                FoundForm = frm
            End If
        Next
        If FoundForm IsNot Nothing Then
            For Each ctrl In FoundForm.Controls
                If ctrl.HasChildren Then
                    If Len(ControlName) > 0 Then
                        childCtrl = FindControl_Recursive(ctrl, ControlName, "")
                    Else
                        childCtrl = FindControl_Recursive(ctrl, "", TagNumber)
                    End If
                    FinalCtrl = childCtrl
                    'Return childCtrl
                    'Exit For
                Else

                    If TypeOf (ctrl) Is ComboBox Or TypeOf (ctrl) Is TextBox Then
                        If UCase(ctrl.Name) = UCase(ControlName) Then
                            FinalCtrl = ctrl
                            Exit For
                        End If
                        If UCase(ctrl.Tag) = UCase(TagNumber) Then
                            'Return ctrl
                            FinalCtrl = ctrl
                            Exit For
                        End If

                    End If

                End If
                If TypeOf (childCtrl) Is ComboBox Or TypeOf (childCtrl) Is TextBox Then
                    FinalCtrl = childCtrl
                    Exit For
                End If
            Next
        End If

        FindControls = FinalCtrl


    End Function

    Public Function FindControl_Recursive(ByVal parent As Control, ByVal ControlName As String, Optional ByVal ControlTAG As String = "") As Control
        Dim ControlIDX As Integer
        Dim tmpctrl As Control
        Dim tmpctrl2 As Control
        For ControlIDX = 0 To parent.Controls.Count - 1
            tmpctrl = parent.Controls(ControlIDX)
            If Len(ControlName) > 0 Then
                If UCase(tmpctrl.Name) = UCase(ControlName) Then
                    Return parent.Controls(ControlIDX)
                ElseIf tmpctrl.Controls.Count > 0 Then
                    tmpctrl2 = FindControl_Recursive(tmpctrl, ControlName, "")
                    If Not IsNothing(tmpctrl2) Then
                        Return tmpctrl2
                    End If
                End If
            End If
            If Len(ControlTAG) > 0 Then
                If UCase(tmpctrl.Tag) = UCase(ControlTAG) Then
                    Return parent.Controls(ControlIDX)
                ElseIf tmpctrl.Controls.Count > 0 Then
                    tmpctrl2 = FindControl_Recursive(tmpctrl, "", ControlTAG)
                    If Not IsNothing(tmpctrl2) Then
                        Return tmpctrl2
                    End If
                End If
            End If
        Next
        ' Not found
        Return Nothing
    End Function


    Private Sub frmMainGIForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '************************************************* LOAD EVENT PROCEDURE *******************************************

        'MsgBox("Welcome to the Goods In Timesheet")

        Initialize("MYSQL")
    End Sub

    Function Get_GUID() As String
        Dim strGUID As String

        strGUID = System.Guid.NewGuid.ToString
        Get_GUID = strGUID
    End Function

    Function GetComputerName() As String
        Dim ComputerName As String

        ComputerName = My.Computer.Name

        GetComputerName = ComputerName
    End Function

    Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim myIP4HostEntry As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each IPAddr As System.Net.IPAddress In myIP4HostEntry.AddressList
            If IPAddr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = IPAddr.ToString()
            End If
        Next

    End Function

    Function GetIPv6Address() As String
        GetIPv6Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim myIP6HostEntry As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each IPAddr As System.Net.IPAddress In myIP6HostEntry.AddressList
            'address = hostEntry.AddressList().Where(Function(a) a.AddressFamily = Sockets.AddressFamily.InterNetworkV6).First().ToString()
            If IPAddr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                GetIPv6Address = IPAddr.ToString()
            End If
        Next

    End Function

    Function GetScreenResolution_Actual(ByVal ScreenNumber As Integer, ByRef ScreenWidth As Integer, ByRef ScreenHeight As Integer,
                                            Optional ByRef Message As String = "") As Integer
        Dim NumberOfScreens As Integer

        NumberOfScreens = Screen.AllScreens.Count
        If ScreenNumber < 2 Then
            ScreenWidth = Screen.PrimaryScreen.Bounds.Width
            ScreenHeight = Screen.PrimaryScreen.Bounds.Height
        Else
            If ScreenNumber <= NumberOfScreens Then
                ScreenWidth = Screen.AllScreens(ScreenNumber).Bounds.Width
                ScreenHeight = Screen.AllScreens(ScreenNumber).Bounds.Height
            Else
                Message = "Error: Passed Parameter Exceeds Number of Screens Available"
                logger.LogError("GI_Error_" & myVersion & ".log", Application.StartupPath, Message, "GetScreenResolution_Actual()",
                                GetComputerName() & "," & GetIPv4Address() & "," & GetIPv6Address() & ", OPERATOR Logged in:" & myUsername,
                                myUsernameSuffix)
            End If
        End If
        GetScreenResolution_Actual = NumberOfScreens

    End Function

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        'call session close routine:

        Application.Exit()
    End Sub

    Sub CreateGIControlPanel()
        Dim newWindow As New Form 'Name of my Default Form's Class assuming Form1.
        Dim newControlWidth As Integer
        Dim newControlHeight As Integer
        Dim FoundControl As Control

        'Replicate the Goods In Timesheet control panel:
        'Will need the control type and the control name and the control tag - with the size and location of each control:
        'read in from the Access DB - tblMainUserformControls perhaps - from the EXCEL version.

        AddFormControls(newWindow, "TEXTBOX", "txtTestBox", "hey", "TAG1", 100, 100, 200, 200)
        'newWindow.Width = findcontrol(newWindow.Controls.name,)
        FoundControl = FindControl(newWindow, "txtTestBox", "")
        newControlWidth = FoundControl.Width
        newControlHeight = FoundControl.Height
        MsgBox("Width: " & CStr(newControlWidth) & ", " & CStr(newControlHeight))
        newWindow.Show()
    End Sub

    Function Get_MDIChildForm(TheFormName As String) As Form
        Dim frmName As String

        Get_MDIChildForm = Nothing
        For Each frm As Form In Application.OpenForms
            If UCase(frm.Name) = UCase(TheFormName) Then
                Get_MDIChildForm = frm
                Exit For
            End If
        Next


    End Function

    Public Function FindMyChild(ByVal ReferenceNumber As String, Optional ByVal ASNNo As String = "", Optional GoToChild As Boolean = True,
                                Optional ByRef Form_UniqueID As String = "",
                                Optional ByRef ChildFormIDX As Integer = 0,
                                Optional ByRef ChildName As String = "",
                                Optional ByVal FormControl As String = "",
                                Optional ByVal BackColor As String = "GREEN") As Boolean
        'Dim MyChild As Form
        Dim SearchField As String
        Dim SearchValue As String
        Dim FoundCtrl As Control
        Dim idx As Integer
        Dim ChildTAG As String
        Dim Title As String
        Dim FormField As String
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ErrMessage As String
        Dim SearchCriteria As String
        Dim FieldType As String
        Dim SearchText As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim FoundASN As Boolean = False
        Dim DBInfoTable As String = "tblDeliveryInfo"

        FindMyChild = False
        DBInfoTable = "tblDeliveryInfo"
        If Application.OpenForms.Count = 0 Then Exit Function
        For Each MyChild As Form In Application.OpenForms
            If Len(ReferenceNumber) > 0 Then
                SearchValue = ReferenceNumber
                FormField = "comDeliveryRef"
                FoundCtrl = FindControl(MyChild, FormField)
            Else
                SearchValue = ""
                FormField = "comASNNo"
                'FIND THE DELIVERY REFERENCE !
                AllValues = Nothing
                AllFields = Nothing
                SortFields = "DeliveryDate"
                Reversed = True
                ErrMessage = ""
                SearchCriteria = ""
                SearchField = "ASN_Number"
                SearchText = ASNNo
                FieldType = "String"
                ReturnField = "ID"
                ReturnValue = ""

                FoundASN = Find_myQuery(myConnString, DBInfoTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                If FoundASN Then
                    SearchValue = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryReference")
                    FoundCtrl = FindControl(MyChild, FormField)
                End If

            End If
            Form_UniqueID = MyChild.Tag
            For idx = 0 To MyChild.MdiChildren.Count - 1
                ChildName = MyChild.MdiChildren(idx).Name
                ChildTAG = MyChild.MdiChildren(idx).Tag
                Title = MyChild.MdiChildren(idx).Text

                If Len(FormControl) = 0 Then
                    FoundCtrl = FindControl(MyChild.MdiChildren(idx), "comDeliveryRef")

                    If FoundCtrl IsNot Nothing Then
                        If InStr(FoundCtrl.Text, SearchValue) > 0 Then
                            If GoToChild Then
                                Application.OpenForms.Item(ChildName).Activate()
                            End If
                            ChildFormIDX = idx
                            ControlPanelIdx = ChildTAG
                            Application.DoEvents()
                            FindMyChild = True
                            Return True
                            Exit Function
                        End If
                    End If
                Else
                    FoundCtrl = FindControl(MyChild.MdiChildren(idx), FormControl)

                    If FoundCtrl IsNot Nothing Then
                        'ControlBACKCOLOUR = System.Drawing.ColorTranslator.ToWin32(FoundCtrl.BackColor)
                        If InStr(Title, SearchValue) > 0 Then
                            If UCase(BackColor) = "GREEN" Then
                                FoundCtrl.BackColor = Color.LimeGreen
                            Else
                                FoundCtrl.BackColor = Color.Red
                            End If
                            ChildFormIDX = idx
                            ControlPanelIdx = ChildTAG
                            Application.DoEvents()
                            FindMyChild = True
                            Return True
                            Exit Function
                        End If
                    End If
                End If
            Next
        Next
    End Function

    Sub Initialize(TheDatabaseType As String)
        Dim NumRows As Long = 0
        Dim conString As String = "" 'not used here.
        Dim Messages As String = ""
        Dim cf As New frmMysqlLogin

        If UCase(DatabaseType) = "ACCDB" Then
            'PopulateDataSource(dgvRegister, DBName, DBProvider, "SELECT " & DisplayRegistryFields & " From tblRegister ORDER BY DateOut DESC", Messages, NumRows)
            Startup(DatabaseType)
        Else
            For xxx = 1 To 3000
                'wait
            Next
            Application.DoEvents()

            If myLoggedIn Then
                Initialised = True
                Startup(DatabaseType)

            Else
                'Now its an MDI app - so create MID child form:
                'cf.MdiParent = Me
                'cf.StartPosition = FormStartPosition.CenterParent
                'cf.Text = "Goods In LOGIN" & CStr(MdiChildren.Count)
                'cf.Name = "GI_LOGIN"
                'cf.Show)
                frmMysqlLogin.ShowDialog()

                frmMysqlLogin.txtUsername.Focus()
                Initialised = True
                Startup(DatabaseType)
            End If
        End If
        zoomcontrolvalue = 50

    End Sub

    Sub Startup(TheDatabaseType As String)
        Dim NumRows As Long = 0
        Dim NumOperators As Long = 0
        Dim conString As String = ""
        Dim Messages As String = ""
        Dim MyOpsMessages As String = ""
        Dim OpMessages As String = ""
        Dim ColWidths As String = ""
        Dim UserId As String = ""
        Dim Username As String = ""
        Dim ComputerName As String = ""
        Dim IPv4Addr As String = ""
        Dim IPv6Addr As String = ""
        Dim EmpNo As String = ""
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim Location As String = ""
        Dim Comments As String = "" 'The Computer Workstation number - usually.
        Dim OnlineMsg As String = ""
        Dim strWorkstationID As String = ""
        Dim OpsOnlineCriteria As String = ""
        Dim StartName As String = "ReferenceProgress"
        'Dim cf As New frmGI_RP_Userform
        Dim cf As New frmReferenceProgress

        Me.WindowState = FormWindowState.Maximized
        Me.ShowIcon = True

        Me.ShowInTaskbar = True
        If TheDatabaseType = "MYSQL" Then
            If myLoggedIn Then
                KeyPreview = True
                Me.DatabaseType = "MYSQL"
                UserId = myUserID
                Username = myUsername

                If Get_NumericPartOfString(Username) > 0 Then
                    myUsernameSuffix = CStr(Get_NumericPartOfString(Username)) 'eg dgoss_2 , dgoss_3 and dgoss_4
                Else
                    myUsernameSuffix = ""
                End If

                ComputerName = GetComputerName()
                IPv4Addr = GetIPv4Address()
                IPv6Addr = GetIPv6Address()
                mySessionID = Get_GUID()
                Location = ""
                EmpNo = myEmpNo
                Firstname = myFirstname
                Lastname = myLastname
                GetDetails(myConnString, "tblOperatorComputers", "ComputerName", ComputerName, "Location", Location, OnlineMsg)
                strWorkstationID = ""
                GetDetails(myConnString, "tblOperatorComputers", "ComputerName", ComputerName, "WorkstationName", strWorkstationID, OnlineMsg)
                'Find the EmpNo logged in now.
                'Remove from the online Operators table first - before adding a new one.
                'so set the criteria: No - use the UserID - only reliable bit of data - as same Operator could log into a diffent machine.
                OpsOnlineCriteria = "ComputerName = " & Chr(39) & ComputerName & Chr(39)
                DeleteMyRecord("tblOperatorsOnline", myConnString, OpsOnlineCriteria, Messages)
                saveSession(myConnString, UserId, Username, ComputerName, IPv4Addr, IPv6Addr, mySessionID, myAccessRights,
                                Messages, EmpNo, Firstname, Lastname, Comments, Location, strWorkstationID, myVersion, MyOpsMessages)
                If myServer! = "localhost" Then
                    'WHAT DOES THIS DO ?
                    'saveSession(myLOCALConnString, UserId, Username, ComputerName, IPv4Addr, IPv6Addr, mySessionID, myAccessRights,
                    'Messages, EmpNo, Firstname, Lastname, Comments, Location, strWorkstationID, MyOpsMessages)
                End If
                If Len(Messages) > 0 Or Len(MyLocalMessage) > 0 Then
                    Me.txtMessages.Text = "Error from Save Session: " & Messages & vbCrLf & " " & MyLocalMessage
                End If

                'Need to adjust the column widths:
                ColWidths = ""
                If Initialised Then
                    'If logged in remotely - does this need to be done twice ?
                    'If logged in locally - myConnString should contain the local connection.
                    'Operators need a criteria - only view those that are logged in now !
                    'CHECK THIS OUT - GIVING ERROR ?????
                    'PopulateMyDataSource(dgvRegister.DataSource, myConnString, "SELECT " & DisplayRegistryFields & " From tblRegister ORDER BY DateOut DESC", NumRows, Messages)
                    'Select Case OPS.username,ops.firstname,ops.lastname,ops.empno,ops.isloggedin,ses.computername,ses.ipv4addr,ses.logindatetime,ses.logoffdatetime
                    'From tblOperators ops
                    'Left Join tblSessions ses on ops.sessionid = ses.sessionid
                    '-- Where ops.username = ses.username
                    '-- And ops.isloggedin = 1
                    'WHERE ops.isloggedin = 1
                    'PopulateMyDataSource(dgvOperators.DataSource, myConnString, DisplayOperatorFields, NumOperators, OpMessages)
                    Application.DoEvents()
                    Beep()
                    Me.myName = Me.myFirstname & " " & Me.myLastname
                    ControlsManager.Initialise()
                    Me.txtMessages.Text = "WELCOME " & Me.myName & " TO THE Goods In TIMESHEET " & myVersion

                    'PopulateTreeView(25000,30000)
                    'CheckAssetTree()
                    Me.CreateNewUsersToolStripMenuItem.Visible = False
                    Me.ShowAltControlToolStripMenuItem.Visible = False
                    Me.SendMessageToolStripMenuItem.Visible = False
                    If UCase(myAccessRights) = "ADMIN" Then
                        Me.CreateNewUsersToolStripMenuItem.Visible = True
                        'Me.AddUserToolStripMenuItem.Enabled = True
                        'Me.AddUserToDB.Visible = True
                        'Me.rbIgnoreNotCheckedOut.Visible = False
                    End If
                    If UCase(myAccessRights) = "SUPER" Then
                        Me.CreateNewUsersToolStripMenuItem.Visible = True
                        Me.ShowAltControlToolStripMenuItem.Visible = True
                        Me.SendMessageToolStripMenuItem.Visible = True
                        'Me.AddUserToolStripMenuItem.Enabled = True
                        'Me.AddUserToDB.Visible = True
                        'Me.rbIgnoreNotCheckedOut.Visible = False

                    End If

                    Me.IsMdiContainer = True

                    ControlPanelFormName = "GI_TIMESHEET_"
                    CPFormName = "GI_TIMESHEET_"
                    ControlPanelIdx = Get_GUID()
                    FormID = ControlPanelIdx

                    cf.MdiParent = Me
                    cf.StartPosition = FormStartPosition.CenterScreen
                    cf.WindowState = FormWindowState.Maximized
                    'cf.Text = "Goods In Control Panel " & Now().ToString("dd/MM/yyyy HH:mm:ss")
                    cf.Text = "Reference Progress"
                    cf.Name = CPFormName & FormID
                    cf.Tag = FormID
                    'Clear_Entry_Controls(CPFormName & FormID, 1, 40)
                    'Clear_Entry_Controls(CPFormName & FormID, 801, 807)
                    cf.Show()
                    'cf.ShowDialog() - gives error

                    '==================================================

                End If
            Else
                MsgBox("Need to LOGIN")
                Application.Exit()
            End If
        End If
        If DEBUG_MODE = False Then
            Timer1.Interval = 1000
            Timer1.Start()
            Timer2.Interval = 5000
            Timer2.Start()

            Me.txtMainClock.Visible = True
        End If
        KeyPreview = True
        Me.TopMost = False
        'Me.ClearBoxes()
        'Me.txtAssetEntry.Focus()
        Beep()
        If Len(Messages) > 0 Then
            'Me.txtStatusMessage.AppendText(CStr(Now() & Messages))
            logger.LogError("GI_ERRORS_" & myVersion & ".log", Application.StartupPath, Messages, "Startup()",
                            GetComputerName() & "," & GetIPv4Address() & "," & GetIPv6Address() & ", OPERATOR LOGGED OUT:" & myUsername,
                            myUsernameSuffix)
            Exit Sub
        End If
        'Me.txtTotalRegisterEntries.Text = CStr(NumRows)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'for MAIN CLOCK:
        Me.txtMainClock.Text = CStr(Now())
        If dic_AnyChanges(ControlPanelIdx) = "YES" Then
            'Me.txtMessages.Text = "CHANGES have Occured"
        End If
        If dic_AnyChanges(ControlPanelIdx) = "NO" Then
            'Me.txtMessages.Text = "NO CHANGES"
        End If
    End Sub

    Private Sub TimesheetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimesheetToolStripMenuItem.Click
        'VIEW TIMESHEET in GRID:

        frmViewExcelSheet.Show()

    End Sub

    Private Sub frmMainGIForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        '*************************************************** SHOWN EVENT PROCEDURE *********************************************

        'Me.WindowState = FormWindowState.Maximized
        Me.StartPosition = FormStartPosition.CenterScreen

        Me.pbarMain.Style = ProgressBarStyle.Blocks
        Me.pbarMain.Value = 0
    End Sub

    Private Sub TimesheetToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TimesheetToolStripMenuItem1.Click
        'For all items in TIMESHEET MENU - executed first:

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim TimesheetForm As Form
        Dim frmName As String
        Dim FindControlName As String

        TimesheetForm = Get_MDIChildForm(frmGI_RP_Userform.Name)

        If Not IsNothing(TimesheetForm) Then
            frmName = TimesheetForm.Name
            If UCase(frmName) = UCase(ControlPanelFormName & ControlPanelIdx) Then
                Me.txtMessages.Text = frmName & " is Open"
            Else
                Me.txtMessages.Text = ""
            End If
        End If


    End Sub

    Private Sub ImportErrorListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportErrorListToolStripMenuItem.Click
        'Dim cf As New frmScrollList

        'cf.MdiParent = Me
        'cf.StartPosition = FormStartPosition.CenterParent
        'cf.Text = "Error List" & CStr(MdiChildren.Count)
        'cf.Name = "Error List"
        'cf.Show()
    End Sub

    Private Sub DatabaseTablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatabaseTablesToolStripMenuItem.Click
        'frmViewDataInGrid
        Dim cf As New frmViewDataInGrid

        cf.MdiParent = Me
        cf.StartPosition = FormStartPosition.CenterParent
        cf.Text = "VIEW Database Information" & CStr(MdiChildren.Count)
        cf.Name = "ViewDB"
        cf.Show()
    End Sub

    Private Sub ShowAltControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowAltControlToolStripMenuItem.Click
        Dim altGrid As New frmGI_ControlPanel_Using_Grid

        altGrid.MdiParent = Me
        altGrid.StartPosition = FormStartPosition.CenterParent
        altGrid.Text = "Goods In Control Panel v2_" & CStr(MdiChildren.Count)
        altGrid.Name = "GI_TIMESHEET2"
        altGrid.Show()

    End Sub

    Private Sub frmMainGIForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim criteria As String
        Dim MyMessage As String = ""
        Dim AreYouSure As Integer
        Dim OpsOnlineCriteria As String
        Dim Messages As String = ""

        'MsgBox("CLOSE MAIN FORM")
        'Needs to check if any changes have been made since LAST SAVE ??? check AnyChanges(ControlPanelIDX) global flag.
        'WELL - in this case will have to be a DICTIONARY variable where the key will be ControlPanelIDX and value is YES/NO
        'So set to TRUE if ANY character typed into either TEXT BOX or COMBO BOX (or start or end button clicked) - see clsControls.
        'set to FALSE after a SUCCESSFULL save.
        criteria = "SessionID = " & Chr(39) & mySessionID & Chr(39)

        If Check_myLoggedIN() Then
            AreYouSure = MsgBox("Are You Sure ?", vbYesNo, "QUIT APPLICATION")
            If AreYouSure = vbYes Then
                If UCase(DatabaseType) = "MYSQL" Then
                    UpdateMySession(myConnString, "", "", "", "", "", mySessionID, "", CStr(Now()), "NO", "", criteria, MyMessage)
                    OpsOnlineCriteria = "ComputerName = " & Chr(39) & GetComputerName() & Chr(39)
                    DeleteMyRecord("tblOperatorsOnline", myConnString, OpsOnlineCriteria, Messages)
                    If Len(Messages) > 0 Then
                        MsgBox("Error during removal of online username: " & Messages)
                    End If
                End If
            Else
                'NO - but user has now closed the main window !
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub OnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimerOnMenuItem.Click
        'Turn on Timer
        MsgBox("Turn ON TIMER")
        Timer1.Interval = 1000
        Timer1.Start()
        Timer2.Interval = 5000
        Timer2.Start()
    End Sub

    Private Sub TimerOffMenuItem_Click(sender As Object, e As EventArgs) Handles TimerOffMenuItem.Click
        'Turn Off Timer
        MsgBox("Turn Off Timer")
        Timer1.Interval = 1000
        Timer1.Stop()
        Timer2.Interval = 5000
        Timer2.Stop()
    End Sub

    Private Sub CreateNewUsersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewUsersToolStripMenuItem.Click
        'frmCreateUsers
        Dim cf As New frmCreateUsers
        Dim CreateUsersForm As Form
        Dim frmName As String
        Dim SearchName As String = "CreateUsers"


        CreateUsersForm = Get_MDIChildForm(SearchName)

        If Not IsNothing(CreateUsersForm) Then
            frmName = CreateUsersForm.Name
            If UCase(frmName) = UCase(SearchName) Then
                CreateUsersForm.Show()
            End If
        Else
            CreateUsersForm = New frmCreateUsers
            CreateUsersForm.Name = SearchName
            CreateUsersForm.Text = "View/Edit Users"
            CreateUsersForm.MdiParent = Me
            CreateUsersForm.StartPosition = FormStartPosition.Manual
            CreateUsersForm.Show()
        End If

    End Sub

    Private Sub ReferenceProgressToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReferenceProgressToolStripMenuItem.Click
        Dim cf As New frmReferenceProgress
        Dim RefProgForm As Form
        Dim frmName As String
        Dim SearchName As String = "ReferenceProgress"

        'Test if child window not already open ?
        'cf.MdiParent = Me

        'cf.StartPosition = FormStartPosition.CenterParent
        'cf.Text = "REFERENCE PROGRESS " & CStr(MdiChildren.Count)
        'cf.Name = "ReferenceProgress"
        'cf.Show()

        RefProgForm = Get_MDIChildForm(SearchName)

        If Not IsNothing(RefProgForm) Then
            frmName = RefProgForm.Name
            If UCase(frmName) = UCase(SearchName) Then
                'ALREADY OPEN:
                'Generic Form or Child Forms of the Generic ?
                Application.OpenForms.Item(SearchName).Activate()

            End If
        Else
            RefProgForm = New frmReferenceProgress
            RefProgForm.Name = SearchName
            RefProgForm.Text = "Reference Progress"
            RefProgForm.MdiParent = Me
            RefProgForm.StartPosition = FormStartPosition.Manual
            RefProgForm.Show()
        End If

    End Sub

    Private Sub ViewOnlineUsersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewOnlineUsersToolStripMenuItem.Click
        Dim cf As New frmOnlineUsers
        Dim OnlineForm As Form
        Dim frmName As String
        Dim SearchName As String = "OnlineUsers"
        'cf.MdiParent = Me
        'cf.StartPosition = FormStartPosition.Manual

        'cf.Text = "Online Users" & CStr(MdiChildren.Count)
        'cf.Name = "OnlineUsers"
        'Test if window already open ?
        OnlineForm = Get_MDIChildForm(SearchName)

        If Not IsNothing(OnlineForm) Then
            frmName = OnlineForm.Name
            If UCase(frmName) = UCase(SearchName) Then
                Application.OpenForms.Item(SearchName).Activate()
            End If
        Else
            OnlineForm = New frmOnlineUsers
            OnlineForm.Name = SearchName
            OnlineForm.Text = "Online Users:"
            OnlineForm.MdiParent = Me
            OnlineForm.StartPosition = FormStartPosition.Manual
            OnlineForm.Show()
        End If


    End Sub

    Private Sub SendMessageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SendMessageToolStripMenuItem.Click
        MsgBox("SEND MESSAGES COMMING SOON ... ")
    End Sub

    Private Sub InboxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InboxToolStripMenuItem.Click
        'INBOX
        MsgBox("PERSONAL INBOX COMMING SOON ... ")
    End Sub

    Private Sub MainGITimesheet_MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MainGITimesheet_MenuStrip1.ItemClicked

    End Sub

    Private Sub EnterArrivalTimeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnterArrivalTimeToolStripMenuItem.Click
        'Show INBOUND OFFICE FORM:
        Dim cf As New frmInboundOfficeEntry
        Dim InboundForm As Form
        Dim frmName As String
        Dim SearchName As String = "Inbound_Office"
        'cf.MdiParent = Me
        'cf.StartPosition = FormStartPosition.Manual

        'cf.Text = "Online Users" & CStr(MdiChildren.Count)
        'cf.Name = "OnlineUsers"
        'Test if window already open ?
        InboundForm = Get_MDIChildForm(SearchName)

        If Not IsNothing(InboundForm) Then
            frmName = InboundForm.Name
            If UCase(frmName) = UCase(SearchName) Then
                Application.OpenForms.Item(SearchName).Activate()
            End If
        Else
            InboundForm = New frmInboundOfficeEntry
            InboundForm.Name = SearchName
            InboundForm.Text = "INBOUND OFFICE:"
            InboundForm.MdiParent = Me
            InboundForm.StartPosition = FormStartPosition.Manual
            InboundForm.Show()
        End If

    End Sub

    Private Sub InboundToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InboundToolStripMenuItem.Click

    End Sub

    Private Sub PreferencesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PreferencesToolStripMenuItem.Click
        'ALLOW USER TO CHOOSE CERTAIN PREFERENCES
        'SUCH AS START FORM: ENTRY FORM or REFERENCE PROGRESS ?
        'COLOURS OF PANELS etc.
        'Show INBOUND OFFICE FORM:
        Dim cf As New frmPreferences
        Dim PrefForm As Form
        Dim frmName As String
        Dim SearchName As String = "Preference_Form"
        'cf.MdiParent = Me
        'cf.StartPosition = FormStartPosition.Manual

        'cf.Text = "Online Users" & CStr(MdiChildren.Count)
        'cf.Name = "OnlineUsers"
        'Test if window already open ?
        PrefForm = Get_MDIChildForm(SearchName)

        If Not IsNothing(PrefForm) Then
            frmName = PrefForm.Name
            If UCase(frmName) = UCase(SearchName) Then
                Application.OpenForms.Item(SearchName).Activate()
            End If
        Else
            PrefForm = New frmPreferences
            PrefForm.Name = SearchName
            PrefForm.Text = "CHANGE PREFERENCES:"
            PrefForm.MdiParent = Me
            PrefForm.StartPosition = FormStartPosition.Manual
            PrefForm.Show()
        End If

    End Sub
End Class
