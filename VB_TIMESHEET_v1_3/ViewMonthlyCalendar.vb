Public Class frmMonthly
    Public dtDateSelected As DateTime
    Private _FormTitle As String
    Public Property FormTitle() As String
        Get
            Return _FormTitle
        End Get
        Set(value As String)
            _FormTitle = value
        End Set
    End Property
    Private Sub Button1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub btnCllse_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Dim PArentForm As Form
        Dim PArentCtrls() As Control
        Dim ctrl As Control
        Dim myTEXTBOX As TextBox

        frmMainGIForm.MonthlyDateSelected = dtDateSelected
        If UCase(Me.FormTitle) = UCase("Select Delivery Date") Then
            frmMainGIForm.GetDeliveryDate = Me.txtSelectedDate.Text
            Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtSelectDeliveryDate", frmMainGIForm.GetDeliveryDate)
            'Populate the comDeliveryRef and comASN dropdowns with values from tblDeliveryInfo:
            Populate_DeliveryDateCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
            Populate_ASNCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
        End If
        If UCase(Me.FormTitle) = UCase("Select IMPORT Date") Then
            frmMainGIForm.GetImportDate = Me.txtSelectedDate.Text
            If Me.pnlDateTo.Visible = True Then
                frmMainGIForm.GetLastImportDate = Me.txtLastSelectedDate.Text
            Else
                frmMainGIForm.GetLastImportDate = ""
            End If
        End If
        If UCase(Me.FormTitle) = UCase("alt:Select Delivery Date") Then
            frmMainGIForm.GetDeliveryDate = Me.txtSelectedDate.Text
            Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtSelectDeliveryDate", frmMainGIForm.GetDeliveryDate)
            'Populate the comDeliveryRef and comASN dropdowns with values from tblDeliveryInfo:
            Populate_DeliveryDateCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
            Populate_ASNCombo(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, frmMainGIForm.GetDeliveryDate, "")
        End If
        If UCase(Me.FormTitle) = UCase("alt:Select IMPORT Date") Then
            frmMainGIForm.GetImportDate = Me.txtSelectedDate.Text
            If Me.pnlDateTo.Visible = True Then
                frmMainGIForm.GetLastImportDate = Me.txtLastSelectedDate.Text
            Else
                frmMainGIForm.GetLastImportDate = ""
            End If
        End If
        'PArentForm = Me.Owner 'returns NOTHING
        'ReDim PArentCtrls(1)
        'PArentCtrls = PArentForm.Controls.Find("txtSelectDeliveryDate", True)
        'ctrl = PArentCtrls(0)
        'myTEXTBOX = CType(ctrl, System.Windows.Forms.TextBox)
        'ctrl.Text = Me.txtSelectedDate.Text
        Me.Close()
    End Sub

    Private Sub monthly_DateChanged(sender As Object, e As DateRangeEventArgs) Handles monthly.DateChanged

        txtSelectedDate.Text = Me.monthly.SelectionStart.ToString("dd/MM/yyyy")
        'txtLastSelectedDate.Text = Me.MonthlyDateTo.SelectionStart.ToString("dd-MM-yyyy")
    End Sub

    Private Sub img_Empty_Square_Click(sender As Object, e As EventArgs) Handles img_Empty_Square.Click
        'SINGLE MONTH VIEW - default

        txtLastSelectedDate.Visible = False
        monthly.SelectionStart = Today
        pnlDateTo.Visible = False
        Me.Width = 340
        Me.Height = 450
        Me.pnlCalendars.Width = 264
        Me.pnlCalendars.Height = 236
        Me.monthly.CalendarDimensions = New Size(1, 1)
        Me.btnClose.Left = 145
        Me.btnClose.Top = 315
    End Sub

    Private Sub img_Horiz_1x3_Click(sender As Object, e As EventArgs) Handles img_Horiz_1x3.Click
        'Calendar layout 1x3
        Dim dtToday As DateTime
        Dim dtStartDate As DateTime
        Dim dtEndDate As DateTime
        Dim CurrentMonth As Integer
        Dim YearMonthAgo As Integer
        Dim YearMonthAhead As Integer
        Dim PreviousMonth As Integer
        Dim StartYear As Integer
        Dim StartMonth As Integer
        Dim StartDay As Integer
        Dim EndYear As Integer
        Dim EndMonth As Integer
        Dim EndDay As Integer

        PreviousMonth = DatePart(DateInterval.Month, DateAdd(DateInterval.Month, -1, Today))
        YearMonthAgo = DatePart(DateInterval.Year, DateAdd(DateInterval.Month, -1, Today))
        YearMonthAhead = DatePart(DateInterval.Year, DateAdd(DateInterval.Month, 1, Today))
        CurrentMonth = DatePart(DateInterval.Month, DateAdd(DateInterval.Month, 0, Today))

        dtToday = Now()

        StartDay = dtToday.Day
        StartMonth = PreviousMonth
        StartYear = YearMonthAgo

        EndDay = dtToday.Day
        EndMonth = CurrentMonth + 1
        EndYear = YearMonthAhead

        dtStartDate = DateSerial(StartYear, StartMonth, StartDay)
        dtEndDate = DateSerial(StartYear, CurrentMonth + 1, StartDay)


        monthly.ShowTodayCircle = True
        monthly.SelectionStart = dtStartDate


        MonthlyDateTo.Visible = True
        txtLastSelectedDate.Visible = True
        MonthlyDateTo.SelectionStart = dtEndDate

        Me.Width = 880
        Me.Height = 352
        Me.pnlCalendars.Width = 860
        Me.pnlCalendars.Height = 236
        Me.monthly.CalendarDimensions = New Size(3, 1)

    End Sub

    Private Sub img_Vert_3x1_Click(sender As Object, e As EventArgs) Handles img_Vert_3x1.Click
        'Calendar layout 3x1
        Me.Width = 322
        Me.Height = 640
        Me.pnlCalendars.Width = 250
        Me.pnlCalendars.Height = 436
        Me.monthly.CalendarDimensions = New Size(1, 3)
    End Sub

    Private Sub img_4x4_Click(sender As Object, e As EventArgs) Handles img_4x4.Click
        'Calendar layout 2x2
        Me.pnlDateTo.Visible = True
        Me.txtLastSelectedDate.Visible = True
        Me.Width = 640
        Me.Height = 450
        Me.pnlCalendars.Width = 250
        Me.pnlCalendars.Height = 436
        Me.monthly.CalendarDimensions = New Size(1, 1)
        Me.btnClose.Left = 285
        Me.btnClose.Top = 315
    End Sub

    Private Sub frmMonthly_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Me.Text = Me.FormTitle
        Me.pnlDateTo.Visible = False

    End Sub

    Private Sub MonthlyDateTo_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthlyDateTo.DateChanged
        txtLastSelectedDate.Text = Me.MonthlyDateTo.SelectionStart.ToString("dd-MM-yyyy")
    End Sub
End Class