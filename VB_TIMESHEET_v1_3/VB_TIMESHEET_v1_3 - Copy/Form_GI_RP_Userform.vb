Public Class frmGI_RP_Userform
    Public strDeliveryDate As String
    Public strDeliveryRef As String
    Public strASNNo As String
    Private _pbar_controlValue As Integer

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

    Private Sub btnAddOperative_Click(sender As Object, e As EventArgs) Handles btnAddOperative.Click
        ' ********* ADD OPERATIVE: *****************
        Dim TotalRows As Long
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim NewRow As Long

        If IsDate(Me.txtDeliveryDate.Text) Then
            strDeliveryDate = CDate(Me.txtDeliveryDate.Text).ToString("dd/MM/yyyy")
            strDeliveryRef = Me.txtDeliveryRef.Text
            ASNNo = Me.txtASNnum.Text
            ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
            If ExtractTotals IsNot Nothing Then
                'TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
                TotalRows = ExtractTotals.Total_Operatives
            End If
            NewRow = TotalRows + 1
            TotalRows = InsertOperatives(True, strDeliveryDate, strDeliveryRef, ASNNo, Me.Frame_Operatives, NewRow)
            'TotalRows = Get_TotalRows("tblOperatives", strDeliveryRef)
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
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTAG
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

        Static Times_ThisProcExecuted As Integer
        'strDeliveryDate = Me.txtDeliveryDate.Text
        'strDeliveryRef = Me.txtDeliveryRef.Text
        strASNNo = Me.txtASNnum.Text

        Times_ThisProcExecuted = Times_ThisProcExecuted + 1
        Beep()
        Me.Frame_Operatives.Controls.Clear()
        Me.Frame_Short_Parts.Controls.Clear()
        Me.Frame_Extra_Parts.Controls.Clear()
        'dic_Controls(strDeliveryDate & "_" & strDeliveryRef) = 1

        ControlsManager.dic_Controls.CompareMode = vbTextCompare

        'Display the Users Online DataGridView:
        strSQL = "SELECT Firstname,Lastname,LogInDatetime,EmpNo,ComputerName,Location FROM tblOperatorsOnline"
        Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)

        Timer_Check_Online.Enabled = True

        tempTotals.Total_Operatives = 0
        tempTotals.Total_ShortParts = 0
        tempTotals.Total_ExtraParts = 0
        tempTotals.HighestOpTAGID = 43
        tempTotals.HighestOpBtnTAGID = 10
        tempTotals.HighestShortTAGID = 1001
        tempTotals.HighestExtraTAGID = 2001
        tempTotals.TotalOpHours = 0
        tempTotals.TotalFLMHours = 0

        strDeliveryDate = Now().ToString("dd/MM/yyyy")
        Me.txtSelectDeliveryDate.Text = strDeliveryDate
        strDeliveryRef = ""

        frmMainGIForm.GetDeliveryDate = strDeliveryDate
        'Call frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtSelectDeliveryDate", frmMainGIForm.GetDeliveryDate)
        'Populate the comDeliveryRef and comASN dropdowns with values from tblDeliveryInfo:
        Populate_DeliveryDateCombo("GI_TIMESHEET", frmMainGIForm.GetDeliveryDate, "")
        Populate_ASNCombo("GI_TIMESHEET", frmMainGIForm.GetDeliveryDate, "")

        If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
        End If


        'TotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef)
        Me.txtTotalOps.Text = CStr(TotalRows)
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

        Me.Frame_FLMDetails.Controls.Clear()
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
            PopulateControls(strPopulateDeliveryDate, strDeliveryREF, "")
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
        Call ImportData(frmMainGIForm.ErrList, DateSelected, LastDateSelected)

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
        'Call UpdateEmployees(frmMainGIForm.ErrList)
        If UBound(frmMainGIForm.ErrList) > 0 Then
            'Messages = ": Total Errors After Import= " & CStr(UBound(frmMainGIForm.ErrList))
        End If
        'by calling the following procedure - will make ME the MAIN FORM as the parent and NOT this form - frmGI_RP_Userform.
        ' - unless we pass ME in as the OWNER.
        frmMainGIForm.ShowForms("frmSelectSheet")
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
        If Me.txtIsItCompleted.Text = "YES" Then
            Me.txtIsItCompleted.Text = "NO"
        Else
            Me.txtIsItCompleted.Text = "YES"
        End If
    End Sub

    Private Sub btnPalletise_Click(sender As Object, e As EventArgs) Handles btnPalletise.Click
        If Me.txtPalletise.Text = "YES" Then
            Me.txtPalletise.Text = "NO"
        Else
            Me.txtPalletise.Text = "YES"
        End If
    End Sub

    Private Sub btnWrapStrap_Click(sender As Object, e As EventArgs) Handles btnWrapStrap.Click
        If Me.txtWrapStrap.Text = "YES" Then
            Me.txtWrapStrap.Text = "NO"
        Else
            Me.txtWrapStrap.Text = "YES"
        End If
    End Sub

    Private Sub btnCollar_Click(sender As Object, e As EventArgs) Handles btnCollar.Click
        If Me.txtCollar.Text = "YES" Then
            Me.txtCollar.Text = "NO"
        Else
            Me.txtCollar.Text = "YES"
        End If
    End Sub

    Private Sub btnBaggingReq_Click(sender As Object, e As EventArgs) Handles btnBaggingReq.Click
        If Me.txtBaggingReq.Text = "YES" Then
            Me.txtBaggingReq.Text = "NO"
        Else
            Me.txtBaggingReq.Text = "YES"
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

    Private Sub txtEmployeeNo_TextChanged(sender As Object, e As EventArgs) Handles txtEmployeeNo.TextChanged
        'Search for name :
        MsgBox("CHANGING ?????")
    End Sub

    Private Sub frmGI_RP_Userform_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'DETECT when the RETURN KEY has been pressed - on the txtEmployee.text control:
        Dim Entry As String

        'MsgBox("KEY has been pressed")
        If e.KeyCode = Keys.Enter Then
            If Me.txtEmployeeNo.ContainsFocus Then
                MsgBox("KEY has been pressed")
                'Entry = Me.txtEmployeeNo.Text
                'SearchAndInsertName(Len(Entry))
            End If
        End If
    End Sub



    Private Sub btnFLMStartTime_Click(sender As Object, e As EventArgs) Handles btnFLMStartTime.Click
        'INSERT START FLM TIME:

    End Sub

    Private Sub btnFLMFinishTime_Click(sender As Object, e As EventArgs) Handles btnFLMFinishTime.Click

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

        Call PopulateControls(strDeliveryDate, SelectedRef, "")
        'MsgBox("Finished Populating from Ref = " & SelectedRef)
        frmMainGIForm.txtMessages.Text = "OK. Finished Populating from Ref = " & SelectedRef
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
        Call PopulateControls(strDeliveryDate, "", SelectedASN)
        'MsgBox("Finished Populating from ASN = " & SelectedASN)
        frmMainGIForm.txtMessages.Text = "OK. Finished Populating from ASN = " & SelectedASN

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call ResetControls(1, 30, True, "", False)
        Call ResetControls(40, 42, False, "", False)
        Call ResetControls(801, 807, False, "", False)
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
        Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)
        Me.dgvOnlineUsers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        'Me.dgvOnlineUsers.Columns(0).Width = 50 ' not work ?
    End Sub

    Private Sub Timer_Check_Refs_Tick(sender As Object, e As EventArgs) Handles Timer_Check_Refs.Tick
        'Check References every 30 seconds:
        'Get reference from txtDeliveryRef.
        'Search in tblDeliveryInfo ?
        'UPDATE field LOCK_STATUS in tblDeliveryInfo accordingly.

    End Sub

    Private Sub frmGI_RP_Userform_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        MsgBox("De-Activated")
    End Sub

    Private Sub comDeliveryRef_Leave(sender As Object, e As EventArgs) Handles comDeliveryRef.Leave
        Dim ControlName As String

        ControlName = e.ToString

        MsgBox("LEAVING: " & ControlName)
    End Sub
End Class