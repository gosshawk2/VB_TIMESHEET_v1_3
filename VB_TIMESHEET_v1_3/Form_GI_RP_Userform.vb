Public Class frmGI_RP_Userform
    Public strDeliveryDate As String
    Public strDeliveryRef As String
    Public strASNNo As String
    Private _pbar_controlValue As Integer
    Private _DeliveryRef As String
    Private _ASNNo As String

    Public Property ProgressBarValue As Integer

        Get
            '_pbar_controlValue = Me.pbar.Value
            Return _pbar_controlValue
        End Get
        Set(value As Integer)
            _pbar_controlValue = value
            'Me.pbar.Value = value
        End Set

    End Property

    Public Property DeliveryRef As String

        Get
            Return _DeliveryRef
        End Get
        Set(value As String)
            _DeliveryRef = value
        End Set

    End Property

    Public Property ASNNo As String

        Get
            Return _ASNNo
        End Get
        Set(value As String)
            _ASNNo = value
        End Set

    End Property

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)
        'Menu Strip Clicked
        'MsgBox("MENU STRIP CLICKED")
    End Sub

    Private Sub btnImportData_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ZoomControl_Scroll(sender As Object, e As ScrollEventArgs) Handles ZoomControl.Scroll
        Dim OldValue As Integer
        Dim NewValue As Integer
        Dim DiffUp As Integer
        Dim DiffDown As Integer

        OldValue = frmMainGIForm.zoomcontrolvalue
        NewValue = ZoomControl.Value
        If NewValue > OldValue Then
            'increase:
            DiffUp = NewValue - OldValue
        Else
            'reduce:
            DiffDown = OldValue - NewValue

        End If

    End Sub

    Private Sub rbDeliveryRef_CheckedChanged(sender As Object, e As EventArgs)
        Me.comDeliveryRef.Visible = True
        Me.comASNNo.Visible = False
        Me.lblSelectDeliveryRefASN.Text = "Select Delivery Ref."
    End Sub

    Private Sub rbASNNo_CheckedChanged(sender As Object, e As EventArgs)
        Me.comDeliveryRef.Visible = False
        Me.comASNNo.Visible = True
        Me.lblSelectDeliveryRefASN.Text = "Select ASN No."
    End Sub

    Private Sub rbDeliveryRef_CheckedChanged_1(sender As Object, e As EventArgs)
        MsgBox("Whats This ? rb_DeliveryRef_CheckedChanged_1")
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub imgCalendar_SelectDeliveryDate_Click(sender As Object, e As EventArgs) Handles imgCalendar_SelectDeliveryDate.Click
        'SELECT DELIVERY DATE:
        'Bring up calendar control to allow user to select date:
        Dim Monthlycal = New frmMonthly

        For Each frm As Form In Application.OpenForms
            If frm Is frmMonthly Then
                frm.Focus()
                Return
            End If
        Next
        Monthlycal.MdiParent = frmMainGIForm 'frmGI_RP_Userform is NOT an MDI Container.
        Monthlycal.StartPosition = FormStartPosition.CenterParent
        Monthlycal.FormTitle = "Select Delivery Date"
        Monthlycal.Name = "MonthlyCalView"
        Monthlycal.Show()

        'frmMonthly.Show()

    End Sub

    Private Sub btnAddOperative_Click(sender As Object, e As EventArgs)
        ' ********* ADD OPERATIVE: *****************
        Dim TotalRows As Long
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim NewRow As Long

        If IsDate(Me.txtDeliveryDate.Text) Then
            strDeliveryDate = CDate(Me.txtDeliveryDate.Text).ToString("dd/MM/yyyy")
            strDeliveryRef = Me.txtDeliveryRef.Text
            ASNNo = Me.txtASNNum.Text
            ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If ExtractTotals IsNot Nothing Then
                'TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
                TotalRows = ExtractTotals.Total_Operatives
            End If
            NewRow = TotalRows + 1
            TotalRows = InsertOperatives(True, strDeliveryDate, strDeliveryRef, ASNNo, Me.Frame_Operatives, NewRow)
            TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)

            Me.txtTotalOps.Text = CStr(TotalRows)
        Else
            'MsgBox("Something wrong with the Delivery Date ?")
            frmMainGIForm.txtMessages.Text = "Something wrong with the Delivery Date ?"
        End If
    End Sub

    Sub InsertOperative(Optional ByVal TimeStartTAG As Long = 400, Optional ByVal LowerTag As Long = 43,
                        Optional ByVal ControlsPerRow As Long = 6, Optional ByVal StartFieldIndex As Long = 5)
        Dim OpID As Long
        Dim UpperTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim TotalRows As Long
        Dim StartBtnTAG As Long = 0
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim strASNNo As String
        Dim HighestOpTag As Long = 0
        Dim TotalControlsInFrame As Long = 0

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:

        strDeliveryDate = Me.txtDeliveryDate.Text
        strDeliveryRef = Me.txtDeliveryRef.Text
        'strDeliveryDate = Me.txtSelectDeliveryDate.Text
        'strDeliveryRef = Me.txtDeliveryRef.Text

        If Len(strDeliveryDate) = 0 Then
            MsgBox("Delivery Date not SET")
            Exit Sub

        End If

        If Len(strDeliveryRef) = 0 Then
            MsgBox("Delivery Reference Not Set")
            Exit Sub
        End If
        UpperTAG = 0
        TotalRows = 0
        StartBtnTAG = 0
        HighestOpTag = 0
        MsgBox("IN INSERT OPERATIVE")
        If Not IsNothing(frmMainGIForm.Dic_TotalOperatives) Then
            If frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) < 1 Then
                frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) = 1
            End If
            TotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1
        End If
        If Not IsNothing(frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef)) Then
            StartBtnTAG = frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef)
        End If
        If Not IsNothing(frmMainGIForm.Dic_HighestOpTAGID(strDeliveryDate & "_" & strDeliveryRef)) Then
            HighestOpTag = frmMainGIForm.Dic_HighestOpTAGID(strDeliveryDate & "_" & strDeliveryRef)
        End If
        strASNNo = Me.txtASNnum.Text
        TotalControlsInFrame = (TotalRows * ControlsPerRow)


        If Len(ErrMessage) > 0 Then
            MsgBox("Error while Getting FieldNames")
            Exit Sub
        End If

        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTag
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTag
        End If

        OpID = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef)
        AddNewOperatives(frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef), OpID, Frame_Operatives, HighestOpTag, StartBtnTAG,
                         strDeliveryDate, strDeliveryRef, strASNNo, LowerTag, UpperTAG, TimeStartTAG, OpFieldnames, TotalRows, StartFieldIndex)
        Me.txtTotalOps.Text = CStr(frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1)
        frmMainGIForm.Dic_HighestOpTAGID(strDeliveryDate & "_" & strDeliveryRef) = HighestOpTag 'Should be 48 after first row entry (48 = comments TAG)
        frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef) = StartBtnTAG
    End Sub

    Sub InsertShorts()
        Dim ShortID As Long = 0
        Dim LowerTAG As Long
        Dim UpperTAG As Long
        Dim TimeStartTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim TotalRows As Long
        Dim StartFieldIndex As Long
        Dim StartBtnTAG As Long = 0
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim strASNNo As String
        Dim HighestShortTag As Long = 0
        Dim ControlsPerRow As Long = 0
        Dim TotalControlsInFrame As Long = 0

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:

        strDeliveryDate = Me.txtSelectDeliveryDate.Text
        strDeliveryRef = Me.txtDeliveryRef.Text

        If Len(strDeliveryDate) = 0 Then
            MsgBox("Delivery Date not SET")
            Exit Sub

        End If

        If Len(strDeliveryRef) = 0 Then
            MsgBox("Delivery Reference Not Set")
            Exit Sub
        End If

        'strDeliveryDate = Me.txtDeliveryDate.Text
        'strDeliveryRef = Me.txtDeliveryRef.Text
        LowerTAG = 1001
        UpperTAG = 0
        TotalRows = 0
        StartBtnTAG = 0
        HighestShortTag = 0
        TimeStartTAG = 1000
        If Not IsNothing(frmMainGIForm.Dic_ShortCount) Then
            If frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef) < 1 Then
                frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef) = 1
            End If
            TotalRows = frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef) - 1
        End If
        If Not IsNothing(frmMainGIForm.Dic_HighestShortTAGID(strDeliveryDate & "_" & strDeliveryRef)) Then
            StartBtnTAG = frmMainGIForm.Dic_HighestShortTAGID(strDeliveryDate & "_" & strDeliveryRef)
        End If
        If Not IsNothing(frmMainGIForm.Dic_HighestShortTAGID(strDeliveryDate & "_" & strDeliveryRef)) Then
            HighestShortTag = frmMainGIForm.Dic_HighestShortTAGID(strDeliveryDate & "_" & strDeliveryRef)
        End If
        strASNNo = Me.txtASNnum.Text
        ControlsPerRow = 3
        TotalControlsInFrame = (TotalRows * ControlsPerRow)

        If Len(ErrMessage) > 0 Then
            MsgBox("Error : " & ErrMessage)
            Exit Sub
        End If

        StartFieldIndex = 3
        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTAG
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTAG
        End If

        ShortID = frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef)
        AddNewShorts(frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef), Frame_Short_Parts, ShortID, HighestShortTag, StartBtnTAG,
                         strDeliveryDate, strDeliveryRef, strASNNo, LowerTAG, UpperTAG, TimeStartTAG, ShortAndExtraFieldnames, TotalRows, StartFieldIndex)
        Me.txtTotalShorts.Text = CStr(frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef) - 1)
        frmMainGIForm.Dic_HighestShortTAGID(strDeliveryDate & "_" & strDeliveryRef) = HighestShortTag 'Should be 48 after first row entry (48 = comments TAG)
        'frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef) = StartBtnTAG
    End Sub

    Private Sub frmGI_RP_Userform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '*********************************************** LOAD EVENT PROCEDURE - INITIALIZE BEFORE FORM APPEARS ****************************************

    End Sub

    Private Sub frmGI_RP_Userform_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        '*********************************************** SHOWN EVENT PROCEDURE - AFTER FORM APPEARS ***************************************************
        Dim ErrMessage As String = ""
        Dim TotalRows As Long
        Dim ReturnColour As String = ""
        Dim ComboArray() As String = Nothing
        Dim ListArray(,) As String = {{0}, {"NONE"}}
        Dim FLMFirstname As String = ""
        Dim FLMLastname As String = ""
        Dim ComboIDX As Long
        Dim Fullname As String = ""
        Dim tempTotals As New clsTotals
        Dim type1 As String
        Dim type2 As String
        Dim type3 As String
        Dim type4 As String
        Dim type5 As String
        Dim NumRows As Integer
        Dim strSQL As String
        Dim RefFromSender As String
        Dim ColonPos As Integer
        Dim comboResult As ComboBox
        Dim comboControl As Control

        Static Times_ThisProcExecuted As Integer
        'strDeliveryDate = Me.txtDeliveryDate.Text
        'strDeliveryRef = Me.txtDeliveryRef.Text

        If UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
            Me.btnUpdateEmployees.Enabled = True
        Else
            Me.btnUpdateEmployees.Enabled = False
        End If

        RefFromSender = sender.ToString
        ColonPos = InStr(RefFromSender, "REF:")
        If ColonPos > 0 Then
            strDeliveryRef = Mid(RefFromSender, ColonPos + 4, Len(RefFromSender))
        Else
            strDeliveryRef = ""
        End If
        Times_ThisProcExecuted = Times_ThisProcExecuted + 1
        Beep()
        Me.Frame_Operatives.Controls.Clear()
        Me.Frame_Short_Parts.Controls.Clear()
        Me.Frame_Extra_Parts.Controls.Clear()
        'dic_Controls(strDeliveryDate & "_" & strDeliveryRef) = 1

        ControlsManager.dic_Controls.CompareMode = vbTextCompare

        'Display the Users Online DataGridView:
        strSQL = "SELECT Firstname,Lastname,LogInDatetime,EmpNo,ComputerName,Location FROM tblOperatorsOnline"
        'Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)

        Timer_Check_Online.Enabled = True

        If Times_ThisProcExecuted = 1 Then
            'Call IsItCompleted()

        End If

        tempTotals.Total_Operatives = 0
        tempTotals.Total_ShortParts = 0
        tempTotals.Total_ExtraParts = 0
        tempTotals.HighestOpTAGID = 43
        tempTotals.HighestOpBtnTAGID = 10
        tempTotals.HighestShortTAGID = 1001
        tempTotals.HighestExtraTAGID = 2001
        tempTotals.TotalOpHours = 0
        tempTotals.TotalFLMHours = 0
        Me.txtCompletedComment.Text = ""

        'Populate the comDeliveryRef and comASN dropdowns with values from tblDeliveryInfo:
        If Len(frmMainGIForm.GetDeliveryDate) = 0 Then
            frmMainGIForm.GetDeliveryDate = Now().ToString("yyyy-MM-dd")
        End If
        Populate_DeliveryDateCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
        Populate_ASNCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
        'ReDim ComboArray(1)
        comboControl = Nothing
        comboResult = Populate_Name_Dropdown(comboControl)
        Me.comSearchName = DirectCast(comboResult, System.Windows.Forms.ComboBox)
        'Me.comSearchName.Items.AddRange(ComboArray)
        If Len(strDeliveryRef) > 0 Or Len(Me.comASNNo.Text) > 0 Then
            'Me.comDeliveryRef.Text = Me.txtDeliveryRef.Text
            'Me.comDeliveryRef.Focus()
            'CAll PopulateControls() here ??
            Call PopulateControls(strDeliveryDate, strDeliveryRef, Me.comASNNo.Text, "", Me.dgvOperatives)
            'SendKeys.Send("{ENTER}")
        Else
            If Len(frmMainGIForm.SelectedDeliveryRef) > 0 Then
                Me.comDeliveryRef.Text = frmMainGIForm.SelectedDeliveryRef
                'MsgBox("Form Property = " & DeliveryRef)
                'frmMainGIForm.SelectedDeliveryRef = ""
                'Populate the Ref/ASN dropdowns here:
            Else
                strDeliveryDate = Now().ToString("dd/MM/yyyy")
                Me.txtSelectDeliveryDate.Text = strDeliveryDate
                strDeliveryRef = ""
            End If
            frmMainGIForm.GetDeliveryDate = strDeliveryDate
            Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtSelectDeliveryDate", frmMainGIForm.GetDeliveryDate)
        End If
        strASNNo = Me.txtASNNum.Text

        '.comDeliveryRef.Sorted = True
        'Me.comASNNo.Sorted = True
        'Populate the comDeliveryRef and comASN dropdowns with values from tblDeliveryInfo:
        'Populate_DeliveryDateCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
        'Populate_ASNCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")

        If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
        End If

        If Len(strASNNo) > 0 Then
            Me.txtASNnum.Text = strASNNo
        End If
        'TotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef)
        'Me.txtTotalOps.Text = CStr(TotalRows)
        If Len(ErrMessage) > 0 Then
            MsgBox(ErrMessage)
        End If

        If Check_ImportDate_IsToday("tblDeliveryInfo", "DeliveryDate", frmMainGIForm.myConnString, ReturnColour, ErrMessage) Then
            'MsgBox("IMPORT DATA IS OK")
        End If
        'MsgBox("Colour = " & ReturnColour)
        If UCase(ReturnColour) = "GREEN" Then
            Me.btnImportData.BackColor = Color.Lime
        ElseIf UCase(ReturnColour) = "AMBER" Then
            Me.btnImportData.BackColor = Color.Orange
        Else
            Me.btnImportData.BackColor = Color.Red
        End If

        If Check_ImportDate_IsToday("tblEmployees", "LastUpdated", frmMainGIForm.myConnString, ReturnColour, ErrMessage) Then
            'MsgBox("EMPLOYEES OK")
        End If
        'MsgBox("Colour = " & ReturnColour)
        If UCase(ReturnColour) = "GREEN" Then
            Me.btnUpdateEmployees.BackColor = Color.Lime
        ElseIf UCase(ReturnColour) = "AMBER" Then
            Me.btnUpdateEmployees.BackColor = Color.Orange
        Else
            Me.btnUpdateEmployees.BackColor = Color.Red
        End If

        'INSERT FLM COntrols:

        'Me.Frame_FLMDetails.Controls.Clear()
        'Call InsertFLMButtons(Me.Frame_FLMDetails, strDeliveryDate, strDeliveryRef, strASNNo)


        KeyPreview = True


    End Sub


    Private Sub btnSaveAndContinue_Click(sender As Object, e As EventArgs) Handles btnSaveAndContinue.Click
        'SAVE HERE:
        Dim strDeliveryDAte As String
        Dim strDeliveryREF As String
        Dim strPopulateDeliveryDate As String

        strDeliveryDAte = Me.txtSelectDeliveryDate.Text
        strDeliveryREF = Me.txtDeliveryRef.Text
        strPopulateDeliveryDate = CDate(strDeliveryDAte).ToString("dd/MM/yyyy")

        If Len(strDeliveryDAte) = 0 Then
            MsgBox("Delivery Date not SET")
            Exit Sub

        End If

        If Len(strDeliveryREF) = 0 Then
            MsgBox("Delivery Reference Not Set")
            Exit Sub
        End If

        If SaveAllControls(strDeliveryDAte, strDeliveryREF) = True Then
            'MsgBox("OK UPDATED - NO ERRORS")
            frmMainGIForm.txtMessages.Text = "OK UPDATED"
            PopulateControls(strPopulateDeliveryDate, strDeliveryREF, "", "", Me.dgvOperatives)
        Else
            'MsgBox("DID NOT UPDATE")
            frmMainGIForm.txtMessages.Text = "DID NOT UPDATE"
        End If

    End Sub

    Private Sub btnImportData_Click_1(sender As Object, e As EventArgs) Handles btnImportData.Click
        'IMPORT DATA from Chettans spreadsheet into tblDaily and tblDeliveryInfo tables in MYSQL database:
        'MsgBox("IMPORT DATE")
        '1) USER chooses the spreadsheet.
        '2) Daily - whole spreadsheet gets totally loaded into memory - into the ExtractARRAY(Rows,Fields) - Daily - titles at row2 and data at row3.
        '3) Extract each field from the ExtractARray row by row. write each row to the MYSQL tables - tblDeliveryInfo and tblDaily
        '4) Repeat for each row in the array.
        '5) BUT WAIT - BEFORE EACH ROW IS WRITTEN TO THE TABLE - CHECK THAT THE ROW DOES NOT EXIST ALREADY. Check Delivery Date and Delivery Reference.
        '6) May need an extra field here to test to make sure each row IS individual ?????
        '***** 02-01-2019 - Some REFERENCES such as starting with A like A7179 have been re-used AND
        '***** have a different supplier against it but at least have a different P.O. / ASN number against each date.

        Dim DateSelected As String = ""
        Dim ReturnColour As String = ""
        Dim ErrMessage As String = ""
        Dim LastDateSelected As String = ""
        Dim dtDateSelected As DateTime
        Dim Answer As Integer

        ReDim frmMainGIForm.ErrList(1)

        dtDateSelected = CDate(Me.txtSelectDeliveryDate.Text)
        DateSelected = frmMainGIForm.GetImportDate
        If DateSelected Is Nothing Then
            DateSelected = Now().ToString("dd/MM/yyyy")
        End If
        LastDateSelected = frmMainGIForm.GetLastImportDate
        If Len(DateSelected) > 0 Then
            Answer = MsgBox("DATE SELECTED: " & DateSelected, vbOKCancel, "Date Selected")
            If Answer = vbCancel Then
                Exit Sub
            End If
        End If
        Call ImportData2(frmMainGIForm.ErrList, DateSelected, LastDateSelected, True)

        'MsgBox("COMPLETED IMPORTING DATA")
        frmMainGIForm.txtMessages.Text = "IMPORT COMPLETED"

        If Check_ImportDate_IsToday("tblDeliveryInfo", "DeliveryDate", frmMainGIForm.myConnString, ReturnColour, ErrMessage) Then
            'MsgBox("IMPORT DATA IS OK")
        End If
        'MsgBox("Colour = " & ReturnColour)
        If UCase(ReturnColour) = "GREEN" Then
            Me.btnImportData.BackColor = Color.Lime
        ElseIf UCase(ReturnColour) = "AMBER" Then
            Me.btnImportData.BackColor = Color.Orange
        Else
            Me.btnImportData.BackColor = Color.Red
        End If

    End Sub

    Private Sub btnUpdateEmployees_Click(sender As Object, e As EventArgs) Handles btnUpdateEmployees.Click
        Dim Messages As String = ""
        Dim ErrMessage As String = ""
        Dim ReturnColour As String = ""

        'Need to read in the XML file ???
        'instead.
        ReDim frmMainGIForm.ErrList(1)
        'Call UpdateEmployees(frmMainGIForm.ErrList, True)
        If UBound(frmMainGIForm.ErrList) > 0 Then
            'Messages = ": Total Errors After Import= " & CStr(UBound(frmMainGIForm.ErrList))
        End If
        'by calling the following procedure - will make ME the MAIN FORM as the parent and NOT this form - frmGI_RP_Userform.
        ' - unless we pass ME in as the OWNER.
        If UCase(frmMainGIForm.myAccessRights) = "SUPER" Then
            frmMainGIForm.ShowForms("frmSelectSheet")
        End If
        'Need to populate the Errors form.



    End Sub



    Private Sub pbar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnArrivedOnTime_Click(sender As Object, e As EventArgs) Handles btnArrivedOnTime.Click


        If Me.txtArrivedOnTime.Text = "YES" Then
            Me.txtArrivedOnTime.Text = "NO"
        Else
            Me.txtArrivedOnTime.Text = "YES"
        End If
    End Sub

    Private Sub btnIsItSafe_Click(sender As Object, e As EventArgs) Handles btnIsItSafe.Click
        If Me.txtIsItSafe.Text = "YES" Then
            Me.txtIsItSafe.Text = "NO"
        Else
            Me.txtIsItSafe.Text = "YES"
        End If
    End Sub

    Private Sub btnCompleted_Click(sender As Object, e As EventArgs) Handles btnCompleted.Click

        If UCase(Me.txtIsItCompleted.Text) = "YES" Then
            Me.txtIsItCompleted.Text = "NO"
        Else
            Me.txtIsItCompleted.Text = "YES"
        End If
    End Sub

    Function IsItCompleted() As Boolean
        Dim Completed As Boolean
        Dim FLMMessages As String = ""
        Dim OpMessages As String = ""
        Dim GridMessage As String = ""

        IsItCompleted = False
        strDeliveryDate = Me.txtDeliveryDate.Text
        strDeliveryRef = Me.txtDeliveryRef.Text
        Me.txtCompletedComment.Text = ""
        Completed = Test_For_Completion(strDeliveryDate, strDeliveryRef, FLMMessages, OpMessages, GridMessage)
        If Len(OpMessages) > 0 Then
            'MsgBox(OpMessages)
        End If
        If Len(FLMMessages) > 0 Then
            'MsgBox(FLMMessages)
        End If
        If Completed Then
            'GREAT - change the STATUS to COMPLETED
            'Completed = Change_STATUS(strDeliveryRef, "COMPLETED")
            Me.txtIsItCompleted.Text = "YES"
        Else
            Me.txtIsItCompleted.Text = "NO"
            Me.txtCompletedComment.AppendText(OpMessages)
            Me.txtCompletedComment.AppendText(vbCrLf & FLMMessages)
            Me.txtCompletedComment.AppendText(vbCrLf & GridMessage)
        End If
    End Function

    Private Sub btnPalletise_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnWrapStrap_Click(sender As Object, e As EventArgs) Handles btnWrapStrap.Click
        If Me.txtWrapStrap.Text = "YES" Then
            Me.txtWrapStrap.Text = "NO"
        Else
            Me.txtWrapStrap.Text = "YES"
        End If
    End Sub

    Private Sub btnCollar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnBaggingReq_Click(sender As Object, e As EventArgs) Handles btnLabelledCorrectly.Click
        If Me.txtLabelledCorrectly.Text = "YES" Then
            Me.txtLabelledCorrectly.Text = "NO"
        Else
            Me.txtLabelledCorrectly.Text = "YES"
        End If
    End Sub

    Private Sub rbDeliveryRef_CheckedChanged_2(sender As Object, e As EventArgs) Handles rbDeliveryRef.CheckedChanged
        'Just Clicked away from rbDeliveryRef onto ASNDeliveryRef
        'Me.comASNNo.Visible = False
        'Me.comDeliveryRef.Visible = True
        'MsgBox("rbDeliveryRef checked")
    End Sub

    Private Sub btnDeleteOperative_Click(sender As Object, e As EventArgs)
        Dim FieldIDX As Long
        Dim TotalRows As Long
        Dim LastHighestTag As Long
        Dim LowestTAG As Long
        Dim ErrMessage As String = ""
        Dim TotalControlsPerRow As Long
        Dim tempTotals As clsTotals
        Dim dtDateTime As DateTime

        strDeliveryDate = Me.txtDeliveryDate.Text
        dtDateTime = CDate(strDeliveryDate)
        strDeliveryDate = dtDateTime.ToString("dd/MM/yyyy")
        strDeliveryRef = Me.txtDeliveryRef.Text
        LastHighestTag = 0
        LowestTAG = 43
        TotalControlsPerRow = 6
        tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef) 'hits the database though.

        If TotalRows > 0 Then
            'Call DeleteControls(strDeliveryDate, strDeliveryRef, Me.Frame_Operatives, TotalControlsPerRow, LastHighestTag, TotalRows, LowestTAG, ErrMessage)
            Call DeleteOpsRow(strDeliveryDate, strDeliveryRef, Me.Frame_Operatives, TotalRows, TotalControlsPerRow)
            If Len(ErrMessage) > 0 Then
                MsgBox(ErrMessage)
                Exit Sub
            End If
            'If LastHighestTag > 0 Then
            TotalRows = TotalRows - 1
            tempTotals.Total_Operatives = TotalRows
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

            'End If
        Else
            MsgBox("NO ROWS TO DELETE")
        End If
        'TotalRows = dic_Totals(strDeliveryDate & "_" & strDeliveryRef) - 1
        Me.txtTotalOps.Text = CStr(TotalRows)
    End Sub

    Sub DeleteShorts()
        Dim FieldIDX As Long
        Dim TotalRows As Long
        Dim LastHighestTag As Long
        Dim LowestTAG As Long
        Dim ErrMessage As String = ""
        Dim TotalControlsPerRow As Long
        Dim tempTotals As clsTotals
        Dim dtDeliveryDate As DateTime

        strDeliveryDate = Me.txtDeliveryDate.Text
        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        strDeliveryRef = Me.txtDeliveryRef.Text
        LastHighestTag = 0
        LowestTAG = 1001
        TotalControlsPerRow = 3
        tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        TotalRows = tempTotals.Total_ShortParts


        If TotalRows > 1 Then
            'Call DeleteControls(strDeliveryDate, strDeliveryRef, Me.Frame_Operatives, TotalControlsPerRow, LastHighestTag, TotalRows, LowestTAG, ErrMessage)
            'Call DeleteOpsRow(strDeliveryDate, strDeliveryRef, Me.Frame_Short_Parts, TotalRows, TotalControlsPerRow)
            Call DeletePartsRow(strDeliveryDate, strDeliveryRef, "txtShortPartNo", Me.Frame_Short_Parts, TotalRows, TotalControlsPerRow)
            If Len(ErrMessage) > 0 Then
                MsgBox(ErrMessage)
                Exit Sub
            End If
            'If LastHighestTag > 0 Then
            TotalRows = TotalRows - 1
            tempTotals.Total_ShortParts = TotalRows
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals


            'End If
        Else
            MsgBox("NO ROWS TO DELETE")
        End If
        'TotalRows = frmMainGIForm.Dic_ShortCount(strDeliveryDate & "_" & strDeliveryRef) - 1
        Me.txtTotalShorts.Text = CStr(TotalRows)
    End Sub

    Sub DeleteExtras()
        Dim FieldIDX As Long
        Dim TotalRows As Long
        Dim LastHighestTag As Long
        Dim LowestTAG As Long
        Dim ErrMessage As String = ""
        Dim TotalControlsPerRow As Long
        Dim tempTotals As clsTotals
        Dim dtDeliveryDate As DateTime

        strDeliveryDate = Me.txtDeliveryDate.Text
        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        strDeliveryRef = Me.txtDeliveryRef.Text

        strDeliveryRef = Me.txtDeliveryRef.Text
        LastHighestTag = 0
        LowestTAG = 1001
        TotalControlsPerRow = 3

        tempTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        TotalRows = tempTotals.Total_ExtraParts


        If TotalRows > 1 Then
            'Call DeleteControls(strDeliveryDate, strDeliveryRef, Me.Frame_Operatives, TotalControlsPerRow, LastHighestTag, TotalRows, LowestTAG, ErrMessage)
            'Call DeleteOpsRow(strDeliveryDate, strDeliveryRef, Me.Frame_Short_Parts, TotalRows, TotalControlsPerRow)
            Call DeletePartsRow(strDeliveryDate, strDeliveryRef, "txtExtraPartNo", Me.Frame_Extra_Parts, TotalRows, TotalControlsPerRow)
            If Len(ErrMessage) > 0 Then
                MsgBox(ErrMessage)
                Exit Sub
            End If
            'If LastHighestTag > 0 Then
            TotalRows = TotalRows - 1
            tempTotals.Total_ExtraParts = TotalRows
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals

            'End If
        Else
            MsgBox("NO ROWS TO DELETE")
        End If
        'TotalRows = frmMainGIForm.Dic_ExtraCount(strDeliveryDate & "_" & strDeliveryRef) - 1
        Me.txtTotalExtras.Text = CStr(TotalRows)
    End Sub

    Private Sub lblTotalHours_Click(sender As Object, e As EventArgs) Handles lblTotalHours.Click
        Dim dblTotalHours As Double
        Dim TimesheetHrs() As String
        Dim TotalRows As Long
        Dim Totals As clsTotals

        If Len(Me.txtDeliveryDate.Text) = 0 Then
            MsgBox("Delivery Date is blank")
            Exit Sub
        Else
            strDeliveryDate = Me.txtDeliveryDate.Text
        End If
        If Len(Me.txtDeliveryRef.Text) = 0 Then
            MsgBox("Delivery Ref is blank")
            Exit Sub
        Else
            strDeliveryRef = Me.txtDeliveryRef.Text
        End If

        strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
        dblTotalHours = 0.0F
        dblTotalHours = Module_TIMESHEET.Get_TotalHours(Me.Frame_Operatives, strDeliveryDate, strDeliveryRef, TimesheetHrs)
        Totals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        TotalRows = Totals.Total_Operatives
        If UBound(TimesheetHrs) >= TotalRows + 1 Then
            Me.txtTotalHours.Text = CStr(TimesheetHrs(TotalRows + 1))
        End If
        If Save_TotalHours(strDeliveryDate, strDeliveryRef, TimesheetHrs) Then
            MsgBox("OK HRS SUCCESSFULLY SAVED")
        End If
    End Sub

    Private Sub imgCalendar_SelectImportDate_Click(sender As Object, e As EventArgs) Handles imgCalendar_SelectImportDate.Click
        'SELECT IMPORT DATE:
        'Bring up calendar control to allow user to select date:
        Dim Monthlycal = New frmMonthly

        For Each frm As Form In Me.MdiChildren
            If frm Is frmMonthly Then
                frm.Focus()
                Return
            End If
        Next

        Monthlycal.MdiParent = frmMainGIForm
        Monthlycal.StartPosition = FormStartPosition.CenterParent
        Monthlycal.FormTitle = "Select IMPORT Date"
        Monthlycal.Name = "MonthlyCalView"
        Monthlycal.Show()


        'frmMonthly.Show()
    End Sub

    Private Sub rbASNNo_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbASNNo.CheckedChanged
        'MsgBox("ASN CLICKED")
        Me.comASNNo.Visible = True
        Me.comDeliveryRef.Visible = False
    End Sub

    Private Sub btnAddShort_Click(sender As Object, e As EventArgs) Handles btnAddShort.Click
        'ADD SHORT PART AND QTY FIELDS:
        'Call InsertShorts()
        Dim TotalRows As Long
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim dtDeliveryDate As DateTime
        Dim NewRow As Long

        'ADD A NEW SHORT ROW - #,PartNO,Qty
        strDeliveryDate = Me.txtDeliveryDate.Text
        If Len(strDeliveryDate) = 0 Then
            MsgBox("DELIVERY DATE is blank")
            Exit Sub
        End If
        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        strDeliveryRef = Me.txtDeliveryRef.Text
        ASNNo = Me.txtASNnum.Text
        'If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        'End If
        If ExtractTotals IsNot Nothing Then
            TotalRows = ExtractTotals.Total_ShortParts
            NewRow = TotalRows + 1
            TotalRows = InsertShortParts(True, Me.txtDeliveryDate.Text, Me.txtDeliveryRef.Text, Me.txtASNnum.Text, Me.Frame_Short_Parts, NewRow)
            Me.txtTotalShorts.Text = CStr(TotalRows)
        Else
            frmMainGIForm.txtMessages.Text = "Could Not Get Totals"
        End If


        '************************************************************************************************************************

    End Sub

    Private Sub btnDeleteShort_Click(sender As Object, e As EventArgs)
        'DELETE SHORT FIELDS IN FRAME:
        Call DeleteShorts()

    End Sub

    Private Sub btnAddExtra_Click(sender As Object, e As EventArgs) Handles btnAddExtra.Click
        'ADD EXTRA FIELDS IN FRAme: #,PartNO,Qty
        Dim TotalRows As Long
        Dim tempTotals As clsTotals
        Dim dtDeliveryDate As DateTime
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim NewRow As Long

        'Call InsertExtraParts(Me.txtDeliveryDate.Text, Me.txtDeliveryRef.Text, Me.txtASNnum.Text, Me.Frame_Extra_Parts, TotalRows)
        Me.txtTotalExtras.Text = CStr(TotalRows)

        strDeliveryDate = Me.txtDeliveryDate.Text
        If Len(strDeliveryDate) = 0 Then
            MsgBox("DELIVERY DATE is blank")
            Exit Sub
        End If
        dtDeliveryDate = CDate(strDeliveryDate)
        strDeliveryDate = dtDeliveryDate.ToString("dd/MM/yyyy")
        strDeliveryRef = Me.txtDeliveryRef.Text
        ASNNo = Me.txtASNnum.Text
        'If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        'End If
        If ExtractTotals IsNot Nothing Then
            TotalRows = ExtractTotals.Total_ExtraParts
            NewRow = TotalRows + 1
            TotalRows = InsertExtraParts2(True, Me.txtDeliveryDate.Text, Me.txtDeliveryRef.Text, Me.txtASNnum.Text, Me.Frame_Extra_Parts, NewRow)

            Me.txtTotalExtras.Text = CStr(TotalRows)
        Else
            frmMainGIForm.txtMessages.Text = "Could Not Get Totals"
        End If


        '************************************************************************************************************************


    End Sub

    Private Sub btnDeleteExtra_Click(sender As Object, e As EventArgs)
        'DELETE FIELDS IN FRAME:
        Call DeleteExtras()


    End Sub

    Private Sub rbDeliveryRef_Click(sender As Object, e As EventArgs) Handles rbDeliveryRef.Click
        'MsgBox("Delivery Ref Clicked") - YES THIS WORKS:
        Me.comASNNo.Visible = False
        Me.comDeliveryRef.Visible = True


    End Sub

    Private Sub rbASNNo_Click(sender As Object, e As EventArgs) Handles rbASNNo.Click
        'MsgBox("ASN NO clicked")
        Me.comASNNo.Visible = True
        Me.comDeliveryRef.Visible = False
    End Sub

    Private Sub txtEmployeeNo_TextChanged(sender As Object, e As EventArgs)
        'Search for name :
        MsgBox("CHANGING ?????")
    End Sub

    Private Sub frmGI_RP_Userform_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'DETECT when the RETURN KEY has been pressed - on the txtEmployee.text control:
        Dim Entry As String

        'MsgBox("KEY has been pressed")
        If e.KeyCode = Keys.Enter Then
            'If Me.txtSearchEmployeeNo.ContainsFocus Then
            'MsgBox("KEY has been pressed")
            'Entry = Me.txtEmployeeNo.Text
            'SearchAndInsertName(Len(Entry))
            'End If
        End If
    End Sub



    Private Sub btnFLMStartTime_Click(sender As Object, e As EventArgs)
        'INSERT START FLM TIME:

    End Sub

    Private Sub btnFLMFinishTime_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub comDeliveryRef_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comDeliveryRef.SelectedIndexChanged
        'Selected Delivery REference:
        Dim SelectedRef As String

        strDeliveryDate = Me.txtSelectDeliveryDate.Text
        If Len(strDeliveryDate) = 0 Then
            'MsgBox("No Delivery Date")
            frmMainGIForm.txtMessages.Text = "No Delivery Date"
            Exit Sub
        End If
        SelectedRef = Me.comDeliveryRef.Text
        If Len(SelectedRef) = 0 Then
            'MsgBox("No Delivery Ref")
            frmMainGIForm.txtMessages.Text = "No Delivery Ref"
            Exit Sub
        End If
        Me.DeliveryRef = SelectedRef
        Beep()
        Beep()

        'Call PopulateControls(strDeliveryDate, SelectedRef, "", "", Me.dgvOperatives)
        'MsgBox("Finished Populating from Ref = " & SelectedRef)
        frmMainGIForm.txtMessages.Text = "OK. Finished Populating from Ref = " & SelectedRef & " (sELECTED INDEX CHANGE)"
    End Sub

    Private Sub comASNNo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comASNNo.SelectedIndexChanged
        'Selected ASN number:
        Dim SelectedASN As String

        strDeliveryDate = Me.txtSelectDeliveryDate.Text
        If Len(strDeliveryDate) = 0 Then
            'MsgBox("No Delivery Date")
            Exit Sub
        End If
        SelectedASN = Me.comASNNo.Text
        If Len(SelectedASN) = 0 Then
            'MsgBox("No ASN Number")
            Exit Sub
        End If
        Me.ASNNo = SelectedASN
        Beep()
        Beep()

        'Call PopulateControls(strDeliveryDate, "", SelectedASN, "", Me.dgvOperatives)
        'MsgBox("Finished Populating from ASN = " & SelectedASN)
        frmMainGIForm.txtMessages.Text = "OK. Finished Populating from ASN = " & SelectedASN

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call ResetControls(1, 30, True, "", False)
        Call ResetControls(40, 42, False, "", False)
        Call ResetControls(801, 813, False, "", False)
        Call ResetControls(441, 442, False, "", False)
        Call ResetControls(0, 0, False, "Frame_Operatives", False)
        Call ResetControls(0, 0, False, "Frame_Short_Parts", False)
        Call ResetControls(0, 0, False, "Frame_Extra_Parts", False)
        Me.comASNNo.Text = ""
        Me.comDeliveryRef.Text = ""

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer_Check_Online.Tick
        Dim strSQL As String
        Dim NumRows As Integer
        Dim ErrMessage As String

        'MsgBox("check Users online") 'WORKS !!!
        strSQL = "SELECT Firstname,Lastname,LogInDatetime,EmpNo,ComputerName,Location FROM tblOperatorsOnline"
        'Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)
        'Me.dgvOnlineUsers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        'Me.dgvOnlineUsers.Columns(0).Width = 50 ' not work ?
    End Sub

    Private Sub Timer_Check_Refs_Tick(sender As Object, e As EventArgs) Handles Timer_Check_Refs.Tick
        'Check References every 30 seconds:
        'Get reference from txtDeliveryRef.
        'Search in tblDeliveryInfo ?
        'UPDATE field LOCK_STATUS in tblDeliveryInfo accordingly.

    End Sub

    Private Sub frmGI_RP_Userform_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        'MsgBox("De-Activated")
        'YEP this is CALLEd wheN any other event to open another form is called:
        'frmMainGIForm.InsertValueIntoForm(ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtMessages", "De-Activated Control")
    End Sub

    Private Sub comDeliveryRef_Leave(sender As Object, e As EventArgs) Handles comDeliveryRef.Leave
        Dim ControlName As String
        Dim comboBox As Control
        Dim txtBOX As Control

        'If ctype(sender, ) Is TextBox Then

        'End If
        'frmMainGIForm.InsertValueIntoForm(ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtMessages", "Leaving DeliveryRef")
        'MsgBox("LEAVING: " & ControlName)
    End Sub

    Private Sub comDeliveryRef_KeyDown(sender As Object, e As KeyEventArgs) Handles comDeliveryRef.KeyDown
        'EXECUTE REFERENCE selected:
        'Selected Delivery REference:
        Dim SelectedRef As String

        MsgBox("IN comDeliveryRef_KeyDown")
        If e.KeyCode = Keys.Return Then

            strDeliveryDate = Me.txtSelectDeliveryDate.Text
            SelectedRef = Me.comDeliveryRef.Text
            If Len(SelectedRef) = 0 Then
                'MsgBox("No Delivery Ref")
                frmMainGIForm.txtMessages.Text = "No Delivery Ref"
                Exit Sub
            End If
            Me.DeliveryRef = SelectedRef
            Beep()
            Beep()

            'Call PopulateControls(strDeliveryDate, SelectedRef, "", "", Me.dgvOperatives)
            'MsgBox("Finished Populating from Ref = " & SelectedRef)
            frmMainGIForm.txtMessages.Text = "OK. Finished Populating from Ref = " & SelectedRef & " (kEY DOWN ACTIVATION)"
        End If
    End Sub

    Private Sub comASNNo_KeyDown(sender As Object, e As KeyEventArgs) Handles comASNNo.KeyDown
        'EXECUTE ASN selected:
        'Selected ASN Number:
        Dim SelectedASN As String

        MsgBox("comASNNo_KeyDown")
        If e.KeyCode = Keys.Return Then

            strDeliveryDate = Me.txtSelectDeliveryDate.Text
            SelectedASN = Me.comASNNo.Text
            If Len(SelectedASN) = 0 Then
                'MsgBox("No Delivery Ref")
                frmMainGIForm.txtMessages.Text = "No ASN Entered."
                Exit Sub
            End If
            Me.ASNNo = SelectedASN
            'Call PopulateControls(strDeliveryDate, "", SelectedASN, "", Me.dgvOperatives)
            'MsgBox("Finished Populating from Ref = " & SelectedRef)
            frmMainGIForm.txtMessages.Text = "OK. Finished Populating from ASN = " & SelectedASN
        End If
    End Sub

    Private Sub frmGI_RP_Userform_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        Dim UniqueID As String
        Dim childForm As Form

        Static NumActivated As Integer
        NumActivated = NumActivated + 1
        UniqueID = frmMainGIForm.ControlPanelIdx 'LAST Form activated.
        If frmMainGIForm.ActiveMdiChild IsNot Nothing Then
            childForm = frmMainGIForm.ActiveMdiChild
            frmMainGIForm.txtMessages.Text = "ACTIVATED FORM: " & childForm.Text
            UniqueID = childForm.Tag
            'If Index > 0 Then
            frmMainGIForm.ControlPanelIdx = UniqueID
            If Len(frmMainGIForm.SelectedDeliveryRef) > 0 Then
                frmMainGIForm.InsertValueIntoForm(childForm.Name, "txtDeliveryRef", frmMainGIForm.SelectedDeliveryRef)
                frmMainGIForm.InsertValueIntoForm(childForm.Name, "txtASNnum", frmMainGIForm.SelectedASN)
                'frmMainGIForm.InsertValueIntoForm(childForm.Name, "txtSelectDeliveryRef", frmMainGIForm.SelectedDeliveryRef)
                'frmMainGIForm.InsertValueIntoForm(childForm.Name, "comDeliveryRef", frmMainGIForm.SelectedDeliveryRef)
                frmMainGIForm.SelectedDeliveryRef = ""
            End If
            'Me.comDeliveryRef.Text = Me.txtDeliveryRef.Text
            'Me.comASNNo.Text = Me.txtASNnum.Text
            'End If
            'End If
        End If
    End Sub

    Private Sub txtSelectDeliveryDate_TextChanged(sender As Object, e As EventArgs) Handles txtSelectDeliveryDate.TextChanged

    End Sub

    Private Sub btnReferenceCompleted_Click(sender As Object, e As EventArgs) Handles btnReferenceCompleted.Click
        'IS THE REFERENCE COMPLETED ?
        'Turn button green if it is or remain red if not:
        Dim Completed As Boolean = False
        Dim FLMMessages As String
        Dim OpMessages As String
        Dim Answer As Integer
        Dim GridMessage As String

        strDeliveryDate = Me.txtDeliveryDate.Text
        strDeliveryRef = Me.txtDeliveryRef.Text
        Me.txtRefCompleteComments.Text = ""
        Call Check_For_Completion(strDeliveryDate, strDeliveryRef, True)

    End Sub

    Sub Set_Ref_Completed()
        'Might not actually be called.
        Dim Completed As Boolean = False
        Dim FLMMessages As String
        Dim OpMessages As String
        Dim Answer As Integer
        Dim GridMessage As String

        If UCase(Me.txtReferenceCompleted.Text) = "YES" Then
            Answer = MsgBox("Change Status to NO ?", vbYesNo, "Confirm")
            If Answer = vbYes Then
                Me.txtReferenceCompleted.Text = "NO"
            End If
        Else
            '1) Make sure ALL OP FINISH TIMES are in place.
            '2) Make sure FLM FINISH TIME is in place.
            '3) Issue an are you sure ? dialog to confirm.
            '4) Write to the tblDeliveryInfo - STATUS field - "COMPLETE". - use ChangeStatus()


            Completed = Test_For_Completion(strDeliveryDate, strDeliveryRef, FLMMessages, OpMessages, GridMessage)
            If Len(OpMessages) > 0 Then
                Me.txtRefCompleteComments.AppendText(OpMessages)
            End If
            If Len(FLMMessages) > 0 Then
                Me.txtRefCompleteComments.AppendText(vbCrLf & FLMMessages)
            End If
            If Completed Then
                'GREAT - change the STATUS to COMPLETED
                Me.txtReferenceCompleted.Text = "YES"
                Me.btnReferenceCompleted.BackColor = Color.Green
                Call Change_STATUS(strDeliveryRef, "COMPLETED")
                Call Change_LockSTATUS(strDeliveryRef, "CLEAR")
            Else
                Me.txtReferenceCompleted.Text = "NO"
                Me.btnReferenceCompleted.BackColor = Color.Red
                'Call Change_LockSTATUS(strDeliveryRef, "LOCKED")
            End If
        End If
    End Sub

    Private Sub btnAnyDamage_Click(sender As Object, e As EventArgs) Handles btnAnyDamage.Click
        'Any Damage
        If Me.txtDamage.Text = "YES" Then
            Me.txtDamage.Text = "NO"
        Else
            Me.txtDamage.Text = "YES"
        End If
    End Sub

    Private Sub btnEuroPallet_Click(sender As Object, e As EventArgs) Handles btnEuroPallet.Click
        If Me.txtEuroPallet.Text = "YES" Then
            Me.txtEuroPallet.Text = "NO"
        Else
            Me.txtEuroPallet.Text = "YES"
        End If
    End Sub

    Private Sub btnMixedPallet_Click(sender As Object, e As EventArgs) Handles btnMixedPallet.Click
        If Me.txtMixedPallet.Text = "YES" Then
            Me.txtMixedPallet.Text = "NO"
        Else
            Me.txtMixedPallet.Text = "YES"
        End If
    End Sub

    Private Sub btnBaggingRequired_Click(sender As Object, e As EventArgs) Handles btnBaggingRequired.Click
        If Me.txtBaggingRequired.Text = "YES" Then
            Me.txtBaggingRequired.Text = "NO"
        Else
            Me.txtBaggingRequired.Text = "YES"
        End If
    End Sub

    Private Sub btnInsertOpRow_Click(sender As Object, e As EventArgs) Handles btnInsertOpRow.Click
        'Take all the details from the single row of boxes and dropdowns in the Frame_Operatives and put into the 
        ' appropriate columns in the new GRID row.
        'Should be contained in appropriate text boxes and dropdowns with TAG numbers from 43 onwards. 446 and 447 for start and end times
        ' for first row etc.
        Dim DBTotalOps As Long
        Dim TotalRows As Long
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim NewRow As Long


        DBTotalOps = Get_TotalRows("tblOperatives", strDeliveryRef)
        MsgBox(DBTotalOps)

        'Call ADD_ROW_TO_DATA_GRID_VIEW_OPERATIVES(Me.dgvOperatives, 1, 43, 9)

        '========================================== reference progress form is causing an error - due to the vertical veriable going out of range for the grid.

        If IsDate(Me.txtDeliveryDate.Text) Then
            strDeliveryDate = CDate(Me.txtDeliveryDate.Text).ToString("dd/MM/yyyy")
            strDeliveryRef = Me.txtDeliveryRef.Text
            ASNNo = Me.txtASNNum.Text
            ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If ExtractTotals IsNot Nothing Then
                'TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
                TotalRows = ExtractTotals.Total_Operatives
            End If
            If Len(Me.txtOpActivity.Text) = 0 Then
                MsgBox("Please select an activity")
                Exit Sub
            End If
            NewRow = TotalRows + 1
            TotalRows = InsertOperatives(True, strDeliveryDate, strDeliveryRef, ASNNo, Me.Frame_Operatives, NewRow)
            TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)

            Me.txtTotalOps.Text = CStr(TotalRows)
        Else
            'MsgBox("Something wrong with the Delivery Date ?")
            MsgBox("SELECT VALID DELIVERY DATE FIRST")
            frmMainGIForm.txtMessages.Text = "Something wrong with the Delivery Date ?"
        End If


    End Sub

    Private Sub btnGO_Click(sender As Object, e As EventArgs)
        'Start SEARCH function:
        'Insert into next row ? or into the row the user has selected ?
        'Need total number of operatives. Default is row 1
        Dim DBTotalOps As Long

        DBTotalOps = Get_TotalRows("tblOperatives", strDeliveryRef)
        'MsgBox(DBTotalOps)
    End Sub

    Private Sub txtEmployeeNo_TextChanged_1(sender As Object, e As EventArgs)
        'Search for Emp No just entered and fill in name in adjacent name box:
        'txtOpName

    End Sub

    Private Sub lblSearchLastname_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub comSearchName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comSearchName.SelectedIndexChanged
        'NAME SELECTED FROM DROPDOWN:
        Dim SelectedIDX As Integer
        Dim Entry As String

        SelectedIDX = comSearchName.SelectedIndex
        Entry = comSearchName.Items(SelectedIDX)
        MsgBox("Entry=" & Entry)
        'Me.txtSearchName.Text = comSearchName.Items(SelectedIDX)

    End Sub

    Private Sub btnStartTime_Click(sender As Object, e As EventArgs) Handles btnOp_StartTime.Click
        'Insert Start Time:
        Dim CurrentDateTime As DateTime
        Dim CurrentDate As String
        Dim CurrentTime As String

        CurrentDate = Now().ToString("yyyy-MM-dd")
        CurrentTime = Now().ToString("HH:mm:dd")
        'NEED TO STORE both parts in clsControls to be available for SAVE.


        'Me.txtStartTime.Text = Now()


    End Sub

    Private Sub dgvOperatives_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOperatives.CellContentClick
        'EXTRACT INFO from GRID ROW clicked and copy info to text boxes to EDIT:
        Dim strFirstCOL As String

        strFirstCOL = dgvOperatives.Item(0, Me.dgvOperatives.CurrentCell.RowIndex).Value.ToString
        'MsgBox("First Col: " & strFirstCOL)

    End Sub

    Private Sub frmGI_RP_Userform_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd

    End Sub

    Private Sub dgvOperatives_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOperatives.CellClick
        'TRY THIS !
        Dim strFirstCOL As String
        'nEED TO FIND THE OPERATIVES REC id GRID COLUMN. SET TO 2ND COLUMN WITH WIDTH OF 0:


        strFirstCOL = dgvOperatives.Item(0, Me.dgvOperatives.CurrentCell.RowIndex).Value.ToString


    End Sub

    Private Sub btnCreateTimesheet_Click(sender As Object, e As EventArgs) Handles btnCreateTimesheet.Click
        'Create TIMESHEET:
        'Either use a template and pass the values OR create from scratch:
        Call Populate_Timesheet()

    End Sub
End Class