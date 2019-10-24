Option Explicit On

Public Class Form_myCalendar
    Public Enum calDayOfWeek
        sunday = 1
        Monday = 2
        Tuesday = 3
        Wednesday = 4
        Thursday = 5
        Friday = 6
        Saturday = 7
    End Enum
    Public Enum InternetControls
        vbUseSystemDayOfWeek = 0
        vbsunday = 1
        vbMonday = 2
        vbTuesday = 3
        vbWednesday = 4
        vbThursday = 5
        vbFriday = 6
        vbSaturday = 7
    End Enum

    Public Enum calFirstWeekOfYear
        FirstJan1 = 1
        FirstFourDays = 2
        FirstFullWeek = 3
    End Enum

    Private UserformEventsEnabled As Boolean = False
    Private DateOut As Date
    Private SelectedDateIn As Date

    Private OkayEnabled As Boolean = False
    Private TodayEnabled As Boolean = False
    Private MinDate As Date
    Private MaxDate As Date
    Private cmbYearMin As Long = 0
    Private cmbYearMax As Long = 0
    Private StartWeek As Integer = 1
    Private WeekOneOfYear As Integer = 1

    Private HoverControlColor As Long           'Original color of the date label that is currently being hovered over
    Private RatioToResize As Double             'Ratio to resize elements of userform. This is set by the DateFontSize argument in the GetDate function

    Private bgDateColor As Long                 'Color of date label backgrounds
    Private bgDateHoverColor As Long            'Color of date label backgrounds when hovering over
    Private bgDateSelectedColor As Long         'Color of selected date label background
    Private lblDateColor As Long                'Font color of date labels
    Private lblDatePrevMonthColor As Long       'Font color of trailing month date labels
    Private lblDateTodayColor As Long           'Font color of today's date
    Private lblDateSatColor As Long             'Font color of Saturday date labels
    Private lblDateSunColor As Long             'Font color of Sunday date labels

    ' Below is a list of all arguments, their data type, and their function:
    '   SelectedDate (Date) - This is the initial selected date on the calendar. Used to
    '       show the users last selection. If this value is set, the calendar will
    '       initialize to the month and year of the SelectedDate. If not, it will
    '       initialize to today's date (with no selection).
    '   FirstDayOfWeek (calDayOfWeek) - Sets which day to use as first day of the week.
    '   MinimumDate (Date) - Restricts the selection of any dates below this date.
    '   MaximumDate (Date) - Restricts the selection of any dates above this date.
    '   RangeOfYears (Long) - Sets the range of years to show in the year combobox in
    '       either direction from the initial SelectedDate. For example, if the
    '       SelectedDate is in 2014, and the RangeOfYears is set to 10 (the default value),
    '       the year combobox will show 10 years below 2014 to 10 years above 2014, so it
    '       will have a range of 2004-2024. Note that if this range falls outside the bounds
    '       set by the MinimumDate or MaximumDate, it will be overridden. Also, this
    '       range does NOT limit the years that a user can select. If the upper limit of
    '       the year combobox is 2024, and the user clicks the month spinner to surpass
    '       December 2024, it will keep right on going to 2025 and beyond (and those
    '       years will be added to the year combobox).
    '   DateFontSize (Long) - Controls the size of the CalendarForm. This value cannot
    '       be set below 9 (the default). To make the form bigger, set this value larger,
    '       and everything else in the userform will be resized to fit.
    '   TodayButton (Boolean) - Controls whether or not the Today button is visible.
    '   OkayButton (Boolean) - Controls whether or not the Okay button is visible. If the
    '       Okay button is enabled, when the user selects a date, it is highlighted, but
    '       is not returned until they click Okay. If the Okay button is disabled,
    '       clicking a date will automatically return that date and unload the form.
    '   ShowWeekNumbers (Boolean) - Controls the visibility of the week numbers.
    '   FirstWeekOfYear (calFirstWeekOfYear) - Sets the behavior of the week numbers. See
    '       the calFirstWeekOfYear Enum in the Global Variables section to see the possible
    '       values and their behavior.
    '   PositionTop (Long) - Sets the top position of the CalendarForm. If no value is
    '       assigned, the CalendarForm is set to position 1 - CenterOwner. Note that
    '       PositionTop and PositionLeft must BOTH be set in order to override the default
    '       center position.
    '   PositionLeft (Long) - Sets the left position of the CalendarForm. If no value is
    '       assigned, the CalendarForm is set to position 1 - CenterOwner. Note that
    '       PositionTop and PositionLeft must BOTH be set in order to override the default
    '       center position.
    '   BackgroundColor (Long) - Sets the background color of the CalendarForm.
    '   HeaderColor (Long) - Sets the background color of the header. The header is the
    '       month and year label at the top.
    '   HeaderFontColor (Long) - Sets the color of the header font.
    '   SubHeaderColor (Long) - Sets the background color of the subheader. The subheader
    '       is the day of week labels under the header (Su, Mo, Tu, etc).
    '   SubHeaderFontColor (Long) - Sets the color of the subheader font.
    '   DateColor (Long) - Sets the background color of the individual date labels.
    '   DateFontColor (Long) - Sets the font color of the individual date labels.
    '   SaturdayFontColor (Long) - Sets the font color of Saturday date labels.
    '   SundayFontColor (Long) - Sets the font color of Sunday date labels.
    '   DateBorder (Boolean) - Controls whether or not the date labels have borders.
    '   DateBorderColor (Long) - Sets the color of the date label borders. Note that the
    '       argument DateBorder must be set to True for this setting to take effect.
    '   DateSpecialEffect (fmSpecialEffect) - Sets a special effect for the date labels.
    '       This can be set to bump, etched, flat (default value), raised, or sunken.
    '       This can be used to make the date labels look like buttons if you desire.
    '       Note that this setting overrides any date border settings you have made.
    '   DateHoverColor (Long) - Sets the background color when hovering the mouse over
    '       a date label.
    '   DateSelectedColor (Long) - Sets the background color of the selected date.
    '   TrailingMonthFontColor (Long) - Sets the color of the date labels in trailing
    '       months. Trailing months are the date labels from last month at the top of the
    '       calendar and from next month at the bottom of the calendar.
    '   TodayFontColor (Long) - Sets the font color of today's date.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GetDate(Optional SelectedDate As Date = Nothing,
    Optional FirstDayOfWeek As calDayOfWeek = vbSunday,
    Optional MinimumDate As Date = Nothing,
    Optional MaximumDate As Date = Nothing,
    Optional RangeOfYears As Long = 10,
    Optional DateFontSize As Long = 9,
    Optional TodayButton As Boolean = False, Optional OkayButton As Boolean = False,
    Optional ShowWeekNumbers As Boolean = False, Optional FirstWeekOfYear As calFirstWeekOfYear = 1,
    Optional PositionTop As Long = -5, Optional PositionLeft As Long = -5,
    Optional BackgroundColor As Long = 16777215,
    Optional HeaderColor As Long = 15658734,
    Optional HeaderFontColor As Long = 0,
    Optional SubHeaderColor As Long = 16448250,
    Optional SubHeaderFontColor As Long = 8553090,
    Optional DateColor As Long = 16777215,
    Optional DateFontColor As Long = 0,
    Optional SaturdayFontColor As Long = 0,
    Optional SundayFontColor As Long = 0,
    Optional DateBorder As Boolean = False, Optional DateBorderColor As Long = 15658734,
    Optional DateSpecialEffect As VariantType = 0,
    Optional DateHoverColor As Long = 15658734,
    Optional DateSelectedColor As Long = 14277081,
    Optional TrailingMonthFontColor As Long = 12566463,
    Optional TodayFontColor As Long = 15773696) As Date

        'Set global variables
        DateFontSize = 9 'Font size cannot be below 9
        OkayEnabled = OkayButton
        TodayEnabled = TodayButton
        RatioToResize = DateFontSize / 9
        bgDateColor = DateColor
        lblDateColor = DateFontColor
        lblDateSatColor = SaturdayFontColor
        lblDateSunColor = SundayFontColor
        bgDateHoverColor = DateHoverColor
        bgDateSelectedColor = DateSelectedColor
        lblDatePrevMonthColor = TrailingMonthFontColor
        lblDateTodayColor = TodayFontColor
        StartWeek = FirstDayOfWeek
        WeekOneOfYear = FirstWeekOfYear

        'Initialize userform
        UserformEventsEnabled = False
        Call Initializeform(SelectedDate, MinimumDate, MaximumDate, RangeOfYears, PositionTop, PositionLeft,
        DateFontSize, ShowWeekNumbers, BackgroundColor, HeaderColor, HeaderFontColor, SubHeaderColor,
        SubHeaderFontColor, DateBorder, DateBorderColor, DateSpecialEffect)
        UserformEventsEnabled = True

        'Show userform, return selected date, and unload
        Me.Show()
        GetDate = DateOut

    End Function

    Private Sub Initializeform(SelectedDate As Date, MinimumDate As Date, MaximumDate As Date,
    RangeOfYears As Long,
    PositionTop As Long, PositionLeft As Long,
    SizeFont As Long, bWeekNumbers As Boolean,
    BackgroundColor As Long,
    HeaderColor As Long,
    HeaderFontColor As Long,
    SubHeaderColor As Long,
    SubHeaderFontColor As Long,
    DateBorder As Boolean, DateBorderColor As Long,
    DateSpecialEffect As VariantType)

        Dim TempDate As Date                        'Used to set selected date, if none has been provided
        Dim SelectedYear As Long                    'Year of selected date
        Dim SelectedMonth As Long                   'Month of selected date
        Dim SelectedDay As Long                     'Day of seledcted date (if applicable)
        Dim TempDayOfWeek As Long                   'Used to set day labels in subheader
        Dim BorderSpacing As Double                 'Padding between the outermost elements of userform and edge of userform
        Dim HeaderDefaultFontSize As Long           'Default font size of the header labels (month and year)
        Dim bgHeaderDefaultHeight As Double         'Default height of the background behind header labels
        Dim lblMonthYearDefaultHeight As Double     'Default height of the month and year header labels
        Dim scrlMonthDefaultHeight As Double        'Default height of the month scroll bar
        Dim bgDayLabelsDefaultHeight As Double      'Default height of the background behind the subheader day of week labels
        Dim bgDateDefaultHeight As Double           'Default height of the background behind each date label
        Dim bgDateDefaultWidth As Double            'Default width of the background behind each date label
        Dim lblDateDefaultHeight As Double          'Default height of each date label
        Dim cmdButtonDefaultHeight As Double        'Default height of Today and Okay command buttons
        Dim cmdButtonDefaultWidth As Double         'Default width of Today and Okay command buttons
        Dim cmdButtonsCombinedWidth As Double       'Combined width of Today and Okay buttons. Used to center on userform
        Dim cmdButtonsMaxHeight As Double           'Maximum height of command buttons and month scroll bar
        Dim cmdButtonsMaxWidth As Double            'Maximum width of command buttons
        Dim cmdButtonsMaxFontSize As Long           'Maximum font size of command buttons
        Dim bgControl As VariantType                'Stores current date label background in loop to initialize various settings
        Dim lblControl As VariantType               'Stores current date label in loop to initialize various settings
        Dim HeightOffset As Double                  'Difference between form height and inside height, to account for toolbar
        Dim i As Long                               'Used for loops
        Dim j As Long                               'Used for loops

        'Initialize default values
        BorderSpacing = 6 * RatioToResize
        HeaderDefaultFontSize = 11
        bgHeaderDefaultHeight = 30
        lblMonthYearDefaultHeight = 13.5
        scrlMonthDefaultHeight = 18
        bgDayLabelsDefaultHeight = 18
        bgDateDefaultHeight = 18
        bgDateDefaultWidth = 18
        lblDateDefaultHeight = 10.5
        cmdButtonDefaultHeight = 24
        cmdButtonDefaultWidth = 60
        cmdButtonsMaxHeight = 36
        cmdButtonsMaxWidth = 90
        cmdButtonsMaxFontSize = 14


        'Set MinDate and MaxDate. If no MinimumDate or MaximumDate are provided, set the
        'MinDate to 1/1/1900 and the MaxDate to 12/31/9999. If MaxDate is less than
        'MinDate, it will default to the MinDate.
        If Not IsDate(MinimumDate) Then
            MinDate = CDate("1/1/1900")
        Else
            MinDate = MinimumDate
        End If
        If Not IsDate(MaximumDate) Then
            MaxDate = CDate("12/31/9999")
        Else
            MaxDate = MaximumDate
        End If
        If MaxDate < MinDate Then MaxDate = MinDate

        'If today's date falls outside min/max, make sure Today button is disabled
        If Now < MinDate Or Now > MaxDate Then TodayEnabled = False

        'Initialize userform position. Initial value of top and left is -5. Check
        'this value to see if a different value has been passed. If not, position
        'to CenterOwner. Must set both top and left positions to override center position
        If PositionTop <> 1 And PositionLeft <> 1 Then
            Me.Top = PositionTop
            Me.Left = PositionLeft
        Else
            Me.Top = 8
            Me.Left = 8
        End If

        'Size header elements - header background, month scroll bar, scroll cover (which is just
        'a blank label which sits on top of the month scroll bar to make it look like two spin
        'buttons), month/year labels in header, and the month and year comboboxes
        With bgHeader
            .Height = bgHeaderDefaultHeight * RatioToResize
            'The header width depends on whether week numbers are visible or not
            If bWeekNumbers Then
                .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
            Else
                .Width = 7 * (bgDateDefaultWidth * RatioToResize)
            End If
            .Left = BorderSpacing
            .Top = BorderSpacing
        End With
        'Month scroll bar. I set a maximum height for the scroll bar, because as it gets
        'larger, the width of the scroll buttons never increases, so eventually it ends
        'up looking really tall and skinny and weird.
        With scrlMonth
            .Width = bgHeader.Width - (2 * BorderSpacing)
            .Left = bgHeader.Left + BorderSpacing
            .Height = scrlMonthDefaultHeight * RatioToResize
            If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
            .Top = bgHeader.Top + ((bgHeader.Height - .Height) / 2)
        End With
        'Cover over month scroll bar
        With bgScrollCover
            .Height = scrlMonth.Height
            .Width = scrlMonth.Width - 25 '25 is the width of the actual scroll buttons, which need to remain visible
            .Left = scrlMonth.Left + 12.5
            .Top = scrlMonth.Top
        End With
        'The .left position of the month and year labels in the header will be set
        'in the function SetMonthYear, as it changes based on the selected month/year.
        'So only the top needs to be positioned now
        With lblMonth
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            '.Font.Size = HeaderDefaultFontSize * RatioToResize
            .Width = HeaderDefaultFontSize * RatioToResize
            .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
        End With
        With lblYear
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            '.Font.Size = HeaderDefaultFontSize * RatioToResize
            .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
        End With
        cmbMonth.Top = lblMonth.Top + (lblMonth.Height - cmbMonth.Height)
        cmbYear.Top = lblYear.Top + (lblYear.Height - cmbYear.Height)

        'Size subheader elements - the subheader bacgkround (bgDayLabels), the day of
        'week labels themselves, and the week number subheader label, if applicable
        With bgDayLabels
            .Height = bgDayLabelsDefaultHeight * RatioToResize
            'The width depends on whether week numbers are visible or not
            If bWeekNumbers Then
                .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
            Else
                .Width = 7 * (bgDateDefaultWidth * RatioToResize)
            End If
            .Left = BorderSpacing
            .Top = bgHeader.Top + bgHeader.Height
        End With
        'Week number subheader label
        If Not bWeekNumbers Then
            lblWk.Visible = False
        Else
            With lblWk
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.Size = SizeFont
                .Width = bgDateDefaultWidth * RatioToResize
                .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
                .Left = BorderSpacing
            End With
        End If
        'Day of week subheader labels
        For i = 1 To 7
            With Me("lblDay" & CStr(i))
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.Size = SizeFont
                .Width = bgDateDefaultWidth * RatioToResize
                .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
                If i = 1 Then
                    'Left position of first label depends on whether week numbers are visible
                    If bWeekNumbers Then
                        .Left = lblWk.Left + lblWk.Width + BorderSpacing
                    Else
                        .Left = BorderSpacing
                    End If
                Else 'All other labels placed directly next to preceding label
                    .Left = Me("lblDay" & CStr(i - 1)).Left + Me("lblDay" & CStr(i - 1)).Width
                End If
            End With
        Next i

        'Size all date labels and backgrounds
        For i = 1 To 6 'Rows
            'First set position and visibility of week number label
            If Not bWeekNumbers Then
                Me("lblWeek" & CStr(i)).Visible = False
            Else
                With Me("lblWeek" & CStr(i))
                    .AutoSize = False
                    .Height = lblDateDefaultHeight * RatioToResize
                    .Font.Size = SizeFont
                    .Width = bgDateDefaultWidth * RatioToResize
                    .Left = BorderSpacing
                    If i = 1 Then
                        .Top = bgDayLabels.Top + bgDayLabels.Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                    Else
                        .Top = Me("bgDate" & CStr(i - 1) & "1").Top + Me("bgDate" & CStr(i - 1) & "1").Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                    End If
                End With
            End If

            'Now set position of each date label in current row
            For j = 1 To 7
            Set bgControl = Me("bgDate" & CStr(i) & CStr(j))
            Set lblControl = Me("lblDate" & CStr(i) & CStr(j))
            'The date label background is sized and placed first. Then the actual date label is simply
            'set to the same position and centered vertically.
            With bgControl
                    .Height = bgDateDefaultHeight * RatioToResize
                    .Width = bgDateDefaultWidth * RatioToResize
                    If j = 1 Then
                        'Left position of first label in row depends on whether week numbers are visible
                        If bWeekNumbers Then
                            .Left = Me("lblWeek" & CStr(i)).Left + Me("lblWeek" & CStr(i)).Width + BorderSpacing
                        Else
                            .Left = BorderSpacing
                        End If
                    Else 'All other labels placed directly next to preceding label in row
                        .Left = Me("bgDate" & CStr(i) & CStr(j - 1)).Left + Me("bgDate" & CStr(i) & CStr(j - 1)).Width
                    End If
                    If i = 1 Then
                        .Top = bgDayLabels.Top + bgDayLabels.Height
                    Else
                        .Top = Me("bgDate" & CStr(i - 1) & CStr(j)).Top + Me("bgDate" & CStr(i - 1) & CStr(j)).Height
                    End If
                End With
                'Size and position actual date label
                With lblControl
                    .AutoSize = False
                    .Height = lblDateDefaultHeight * RatioToResize
                    .Font.Size = SizeFont
                    .Width = bgControl.Width
                    .Left = bgControl.Left
                    .Top = bgControl.Top + ((bgControl.Height - .Height) / 2)
                End With
            Next j
        Next i

        'Set userform width. Height set later, since it depends on Today and Okay buttons
        frameCalendar.Width = bgDate67.Left + bgDate67.Width + BorderSpacing
        'Make sure userform is large enough to show entire calendar
        If Me.InsideWidth < (frameCalendar.Left + frameCalendar.Width) Then
            Me.Width = Me.Width + ((frameCalendar.Left + frameCalendar.Width) - Me.InsideWidth)
        End If

        'Set size and visibility of Okay button and date selection labels
        If Not OkayEnabled Then
            cmdOkay.Visible = False
            lblSelection.Visible = False
            lblSelectionDate.Visible = False
        Else
            'Okay button. I set a maximum and width, for the same reason as the month
            'scroll bar. Eventually, the gigantic buttons just start looking weird.
            With cmdOkay
                .Visible = True
                .Height = cmdButtonDefaultHeight * RatioToResize
                If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
                .Width = cmdButtonDefaultWidth * RatioToResize
                If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
                If SizeFont > cmdButtonsMaxFontSize Then
                    .Font.Size = cmdButtonsMaxFontSize
                Else
                    .Font.Size = SizeFont
                End If
                .Top = bgDate61.Top + bgDate61.Height + bgDayLabels.Height + BorderSpacing
            End With
            'The "Selection" label
            With lblSelection
                .Visible = True
                .AutoSize = False
                .Height = lblMonthYearDefaultHeight * RatioToResize
                .Width = frameCalendar.Width
                .Font.Size = HeaderDefaultFontSize * RatioToResize
                .AutoSize = True
                .Top = (bgDate61.Top + bgDate61.Height) + ((bgDayLabels.Height + BorderSpacing - .Height) / 2)
            End With
            'The actual selected date label
            With lblSelectionDate
                .Visible = True
                .AutoSize = False
                .Height = lblMonthYearDefaultHeight * RatioToResize
                .Width = frameCalendar.Width - lblSelection.Width
                .Font.Size = HeaderDefaultFontSize * RatioToResize
                .Top = lblSelection.Top
            End With
        End If

        'Set size and visibility of Today button. Make sure it is within max bounds.
        'Top is not set for Today button yet, because it depends on whether Okay button
        'is enabled. Therefore, it is set farther down.
        If Not TodayEnabled Then
            cmdToday.Visible = False
        Else
            With cmdToday
                .Visible = True
                .Height = cmdButtonDefaultHeight * RatioToResize
                If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
                .Width = cmdButtonDefaultWidth * RatioToResize
                If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
                If SizeFont > cmdButtonsMaxFontSize Then
                    .Font.Size = cmdButtonsMaxFontSize
                Else
                    .Font.Size = SizeFont
                End If
            End With
        End If

        'Position Okay and Today buttons, depending on which ones are enabled
        If OkayEnabled And TodayEnabled Then 'Both buttons enabled.
            cmdToday.Top = cmdOkay.Top
            cmdButtonsCombinedWidth = cmdToday.Width + cmdOkay.Width
            cmdToday.Left = ((frameCalendar.Width - cmdButtonsCombinedWidth) / 2) - (BorderSpacing / 2)
            cmdOkay.Left = cmdToday.Left + cmdToday.Width + BorderSpacing
        ElseIf OkayEnabled Then 'Only Okay button enabled
            cmdOkay.Left = (frameCalendar.Width - cmdOkay.Width) / 2
        ElseIf TodayEnabled Then 'Only Today button enabled
            cmdToday.Top = bgDate61.Top + bgDate61.Height + BorderSpacing
            cmdToday.Left = (frameCalendar.Width - cmdToday.Width) / 2
        End If

        'Set userform height, depending on which buttons are enabled
        HeightOffset = Me.Height - Me.InsideHeight
        If OkayEnabled Then
            frameCalendar.Height = cmdOkay.Top + cmdOkay.Height + HeightOffset + BorderSpacing
        ElseIf TodayEnabled Then 'Only Today button enabled
            frameCalendar.Height = cmdToday.Top + cmdToday.Height + HeightOffset + BorderSpacing
        Else 'Neither button enabled
            frameCalendar.Height = bgDate61.Top + bgDate61.Height + HeightOffset + BorderSpacing
        End If

        'Make sure userform is large enough to show entire calendar
        If Me.InsideHeight < (frameCalendar.Top + frameCalendar.Height) Then
            Me.Height = Me.Height + ((frameCalendar.Top + frameCalendar.Height) - Me.InsideHeight - HeightOffset)
        End If

        'Check if SelectedDateIn was set by user, and ensure it is within min/max range
        If SelectedDate > 0 Then
            If SelectedDate < MinDate Then
                SelectedDate = MinDate
            ElseIf SelectedDate > MaxDate Then
                SelectedDate = MaxDate
            End If
            SelectedDateIn = SelectedDate
            SelectedYear = Year(SelectedDateIn)
            SelectedMonth = Month(SelectedDateIn)
            SelectedDay = Day(SelectedDateIn)
            Call SetSelectionLabel(SelectedDateIn)
        Else 'No SelectedDate provided, default to today's date
            cmdOkay.Enabled = False
            TempDate = Date
            If TempDate < MinDate Then
                TempDate = MinDate
            ElseIf TempDate > MaxDate Then
                TempDate = MaxDate
            End If
            SelectedYear = Year(TempDate)
            SelectedMonth = Month(TempDate)
            SelectedDay = 0 'Don't want to highlight a 'selected date,' since user supplied no date
            Call SetSelectionLabel(Empty)
        End If

        'Initialize month and year comboboxes, as well as month scroll bar. Make sure
        'years are within range of 1900 to 9999. If year combobox falls outside bounds
        'of MinDate and MaxDate, it will be overridden.
        Call SetMonthCombobox(SelectedYear, SelectedMonth)
        scrlMonth.Value = SelectedMonth
        cmbYearMin = SelectedYear - RangeOfYears
        cmbYearMax = SelectedYear + RangeOfYears
        If cmbYearMin < Year(MinDate) Then
            cmbYearMin = Year(MinDate)
        End If
        If cmbYearMax > Year(MaxDate) Then
            cmbYearMax = Year(MaxDate)
        End If
        For i = cmbYearMin To cmbYearMax
            cmbYear.AddItem i
    Next i
        cmbYear.value = SelectedYear

        'Set userform colors and effects
        Me.BackColor = BackgroundColor
        frameCalendar.BackColor = BackgroundColor
        bgHeader.BackColor = HeaderColor
        bgScrollCover.BackColor = HeaderColor
        lblMonth.ForeColor = HeaderFontColor
        lblYear.ForeColor = HeaderFontColor
        lblSelection.ForeColor = SubHeaderFontColor
        lblSelectionDate.ForeColor = SubHeaderFontColor
        bgDayLabels.BackColor = SubHeaderColor
        For i = 1 To 7
            Me("lblDay" & CStr(i)).ForeColor = SubHeaderFontColor
        Next i
        If bWeekNumbers Then
            lblWk.ForeColor = SubHeaderFontColor
            For i = 1 To 6
                Me("lblWeek" & CStr(i)).ForeColor = SubHeaderFontColor
            Next i
        End If
        For i = 1 To 6
            For j = 1 To 7
                With Me("bgDate" & CStr(i) & CStr(j))
                    If DateBorder Then
                        .BorderStyle = fmBorderStyleSingle
                        .BorderColor = DateBorderColor
                    End If
                    .SpecialEffect = DateSpecialEffect
                End With
            Next j
        Next i

        'Initialize subheader day labels, based on selected first day of week
        TempDayOfWeek = StartWeek
        For i = 1 To 7
            Me("lblDay" & CStr(i)).Caption = Choose(TempDayOfWeek, "Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
            TempDayOfWeek = TempDayOfWeek + 1
            If TempDayOfWeek = 8 Then TempDayOfWeek = 1
        Next i

        'Set month and year labels in header, as well as date labels
        Call SetMonthYear(SelectedMonth, SelectedYear)
        Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'This is a THREADED OBJECT to perform background tasks. (I will think of a reason in a calendar selection form one day)

    End Sub

    Private Sub scrlMonth_Scroll(sender As Object, e As ScrollEventArgs) Handles scrlMonth.Scroll

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub cmdOkay_Click(sender As Object, e As EventArgs) Handles cmdOkay.Click
        'OK BUTTON:

    End Sub

    Private Sub cmdToday_Click(sender As Object, e As EventArgs) Handles cmdToday.Click
        'TODAY BUTTON:

    End Sub
End Class