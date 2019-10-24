Public Class clsControls
    'Handles the event changes for each control
    Public WithEvents cbTimeStartEvent As Button
    Public WithEvents cbTimeEndEvent As Button
    Public WithEvents comboAfterUpdate As ComboBox
    Public WithEvents txtBoxAfterUpdate As TextBox
    Public WithEvents comboAfterUpdate2 As ComboBox
    Private _txtboxNameOfChangedValue As String 'TextBoxName
    Private _comboNameOfChangedValue As String 'ComboBoxName
    Private _txtboxTagOfChangedValue As String 'TextBoxTAG
    Private _comboTagOfChangedValue As String 'ComboTAG
    Private _txtboxChangedValue As String 'TextBoxValue
    Private _comboChangedValue As String 'ComboValue
    Private _CtrlID As String 'Delivery Date _ Delivery Ref _ TAG Number
    Private _CtrlAltID As String 'Delivery Date _ Delivery Ref _ Control Name
    Private _CtrlASNID As String 'Delivery Date _ ASN _ TAG Number
    Private _CtrlASNAltID As String 'Delivery Date _ ASN _ Control Name
    Private _CTRL As Control
    Private _CtrlOBJECT As Object
    Private _CtrlName As String
    Private _CtrlFrameName As String
    Private _CtrlTag As String
    Private _CtrlDate As Date
    Private _CtrlLeft As Integer
    Private _CtrlTop As Integer
    Private _CtrlWidth As Integer
    Private _CtrlHeight As Integer
    Private _CtrlType As String
    Private _CtrlDeliveryDate As Date
    Private _CtrlDeliveryRef As String
    Private _CtrlASN As String
    Private _CtrlObjNumber As Long 'Count Number of Objects within one FRAME control = Number of CHILD RECORDS in Database
    Private _CtrlValue As String
    Private _CtrlList As Object
    Private _CtrlStartTag As String
    Private _CtrlEndTag As String
    Private _CtrlFrameRowNumber As Long
    Private _CtrlOpRowNumber As Long
    Private _CtrlShortRowNumber As Long
    Private _CtrlExtraRowNumber As Long
    Private _CtrlTotalRows As Long
    Private _CtrlFLMName As String
    Private _CtrlFLMStartDateTime As Date
    Private _CtrlFLMEndDateTime As Date
    Private _CtrlOpName As String
    Private _CtrlOpActivity As String
    Private _CtrlOpStartDateTime As Date
    Private _CtrlOpEndDateTime As Date
    Private _CtrlPartNo As String
    Private _CtrlQty As Long
    Private _CtrlExtraShort As String
    Private _CtrlDBTable As String
    Private _CtrlFieldName As String
    Private _CtrlLastSaved As Date
    Private _CtrlBackColour As Long
    Private _CtrlLeftMargin As Boolean
    Private _CtrlErrMessage As String
    Private _CtrlPrimaryKey As Long
    Private _CtrlFontName As String
    Private _CtrlFontSize As Integer
    Private _CtrlForeColour As Integer
    Private _CtrlFontStyle As String
    Private _CtrlUpdatedByUsername As String
    Private _CtrlUpdatedByEmpNo As String
    Private _CtrlUpdatedByName As String
    Private _CtrlTabIndex As Long
    Private _CtrlOPdblTotalHrs As Double
    Private _CtrlOpdblTotalSecs As Double
    Private _CtrlOPstrTotalHrs As String
    Private _CtrlFLMdblTotalHrs As Double
    Private _CtrlFLMstrTotalHrs As String
    Private _CtrlEnableControl As Boolean



    Private Sub cbTimeStartEvent_click(sender As Object, e As System.EventArgs) Handles cbTimeStartEvent.Click
        Dim myButton As Button
        Dim TagNumber As String
        Dim TimeControl As Control
        Dim SearchControlName As String
        Dim TimeOut As DateTime
        Dim ChangeProperty As String
        Dim testCollection As Object = Nothing
        Dim txtControl As TextBox

        SearchControlName = Replace(UCase(cbTimeStartEvent.Name), UCase("btn"), "txt")
        'Need FindFormControl()

        If TypeOf (sender) Is System.Windows.Forms.Button Then
            myButton = CType(sender, System.Windows.Forms.Button)
            'MsgBox("clsControls TIME START BUtton Pressed: " & myButton.Name & ", Caption: " & myButton.Text) 'YES THIS WORKS !!!!!
            'TimeControl = FindControl(frmGI_RP_Userform, SearchControlName) 'Cannot find for some reason.
            'TimeControl = SearchFrame(frmGI_RP_Userform.Frame_Operatives, SearchControlName)
            TimeControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", SearchControlName)
            If TimeControl Is Nothing Then
                MsgBox("CANNOT FIND START TIME CONTROL")
            Else
                txtControl = CType(TimeControl, System.Windows.Forms.TextBox)
                TagNumber = txtControl.Tag
                TimeOut = Now()
                If InStr(UCase(txtControl.Name), "FLM") > 0 Then
                    Me.ControlFLMStartDateTime = TimeOut
                    txtControl.Text = Format(TimeOut, "HH:mm:ss")
                    ChangeProperty = "TheControl"
                    testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "FLMSTARTTIME"
                    testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                ElseIf InStr(ucase(txtControl.Name), "OP") > 0 Then
                    Me.ControlOpStartDateTime = TimeOut
                    txtControl.Text = Format(TimeOut, "HH:mm:ss")
                    ChangeProperty = "TheControl"
                    testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "OpStartTime"
                    testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                Else
                    Me.ControlErrMessage = "CONTROL: " & txtControl.Name & " , Not handled"
                    MsgBox("CONTROL NOT HANDLED: " & txtControl.Name)
                End If
                cbTimeStartEvent.Enabled = False
            End If
        End If

    End Sub

    Private Sub cbTimeEndEvent_click(sender As Object, e As System.EventArgs) Handles cbTimeEndEvent.Click
        Dim myButton As Button
        Dim TagNumber As String
        Dim TimeControl As Control
        Dim StartTimeControl As Control
        Dim SearchControlName As String
        Dim TimeOut As DateTime
        Dim ChangeProperty As String
        Dim testCollection As Object = Nothing
        Dim txtControl As TextBox
        Dim FLMDetailsFrame As ScrollableControl
        Dim OpFrame As ScrollableControl
        Dim dblTotalHours As Double
        Dim TimeSheetHrs() As String = Nothing

        SearchControlName = Replace(UCase(cbTimeEndEvent.Name), UCase("btn"), "txt")
        'Need FindFormControl()

        If TypeOf (sender) Is System.Windows.Forms.Button Then
            myButton = CType(sender, System.Windows.Forms.Button)
            'MsgBox("clsControls TIME START BUtton Pressed: " & myButton.Name & ", Caption: " & myButton.Text) 'YES THIS WORKS !!!!!
            'TimeControl = FindControl(frmGI_RP_Userform, SearchControlName) 'Cannot find for some reason.
            'TimeControl = SearchFrame(frmGI_RP_Userform.Frame_Operatives, SearchControlName)
            TimeControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", SearchControlName)
            StartTimeControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", "", CStr(TimeControl.Tag - 1))
            If StartTimeControl IsNot Nothing Then
                'MsgBox("Found Start Control: " & StartTimeControl.Name)
            End If
            If TimeControl Is Nothing Then
                MsgBox("CANNOT FIND END TIME CONTROL")
            Else
                txtControl = CType(TimeControl, System.Windows.Forms.TextBox)
                TagNumber = txtControl.Tag
                TimeOut = Now()
                If InStr(UCase(txtControl.Name), "FLM") > 0 Then
                    Me.ControlFLMEndDateTime = TimeOut
                    txtControl.Text = Format(TimeOut, "HH:mm:ss")
                    ChangeProperty = "TheControl"
                    testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "FLMFINISHTIME"
                    testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    FLMDetailsFrame = GetFrameControl("GI_TIMESHEET", "Frame_FLMDetails")
                    dblTotalHours = Get_TotalHours(FLMDetailsFrame, Me.ControlDeliveryDate, Me.ControlDeliveryRef, TimeSheetHrs)
                ElseIf InStr(ucase(txtControl.Name), "OP") > 0 Then


                    Me.ControlOpEndDateTime = TimeOut
                    txtControl.Text = Format(TimeOut, "HH:mm:ss")
                    ChangeProperty = "TheControl"
                    testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, txtControl.Tag, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "OPFINISHTIME"
                    testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, txtControl.Tag, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "FIELDNAME"
                    testCollection = UpdateCollection(dic_Controls, Nothing, "OP_FINISHTIME", Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, txtControl.Tag, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    OpFrame = GetFrameControl("GI_TIMESHEET", "Frame_Operatives")
                    dblTotalHours = Get_TotalHours(OpFrame, Me.ControlDeliveryDate, Me.ControlDeliveryRef, TimeSheetHrs)
                    'Insert into Comments here:

                Else
                    Me.ControlErrMessage = "CONTROL: " & txtControl.Name & " , Not handled"
                    MsgBox("CONTROL NOT HANDLED: " & txtControl.Name)
                End If

            End If
        End If

    End Sub

    Private Sub txtBoxAfterUpdate_change(sender As Object, e As System.EventArgs) Handles txtBoxAfterUpdate.TextChanged
        Dim myTEXTBOX As TextBox
        Dim NewValue As String
        Dim CollectionKey As String
        Dim testCollection As Object = Nothing
        Dim NewFriendlyName As String
        Dim lblControl As Control
        Dim FoundControl As Control
        Dim ChangeProperty As String
        Dim TagNumber As String

        If TypeOf (sender) Is System.Windows.Forms.TextBox Then
            myTEXTBOX = CType(sender, System.Windows.Forms.TextBox)
            'MsgBox("clsCONTROLS = TEXT CHANGED - Control:" & myTEXTBOX.Name & ", value= " & myTEXTBOX.Text) ' YES THIS WORKS !!!!!!!!
            Me.TextboxName = txtBoxAfterUpdate.Name
            Me.TextBoxTAG = txtBoxAfterUpdate.Tag
            Me.TextBoxValue = txtBoxAfterUpdate.Text
            NewValue = Me.TextBoxValue

            If UCase(Me.ControlName) = UCase("OP_STARTTIME") Then
                NewValue = Me.ControlDeliveryDate & " " & NewValue
                FoundControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", Me.ControlName)
                myTEXTBOX = CType(FoundControl, System.Windows.Forms.TextBox)
                myTEXTBOX.Text = Format(NewValue, "HH:mm:ss")
                ChangeProperty = "TheControl"
                TagNumber = myTEXTBOX.Tag
                testCollection = UpdateCollection(dic_Controls, Nothing, myTEXTBOX, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                ChangeProperty = "OpStartTime"
                testCollection = UpdateCollection(dic_Controls, Nothing, myTEXTBOX, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
            End If
            If UCase(Me.ControlName) = UCase("OP_FINISHTIME") Then
                NewValue = Me.ControlDeliveryDate & " " & NewValue
                FoundControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", Me.ControlName)
                myTEXTBOX = CType(FoundControl, System.Windows.Forms.TextBox)
                myTEXTBOX.Text = Format(NewValue, "HH:mm:ss")
                ChangeProperty = "TheControl"
                TagNumber = myTEXTBOX.Tag
                testCollection = UpdateCollection(dic_Controls, Nothing, myTEXTBOX, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
                ChangeProperty = "OPFINISHTIME"
                testCollection = UpdateCollection(dic_Controls, Nothing, myTEXTBOX, Me.ControlDeliveryDate, Me.ControlASN,
                                                      Me.ControlDeliveryRef, TagNumber, ChangeProperty)
            End If
            'EMPLOYEE NUMBER ENTERED BY USER:
            If UCase(Me.ControlName) = UCase("txtEmployeeNo") Then
                If Len(NewValue) > 6 Then
                    'MsgBox("clsControls: Employee Number Entry - " & NewValue)
                    'Entry = Me.txtEmployeeNo.Text
                    SearchAndInsertName(NewValue, Me.ControlDeliveryDate, Me.ControlDeliveryRef, Len(NewValue))
                End If
            End If
            CollectionKey = Me.ControlDeliveryDate & "_" & Me.ControlDeliveryRef & "_" & Me.TextBoxTAG
            testCollection = UpdateCollection(dic_Controls, CollectionKey, NewValue, Me.ControlDeliveryDate)
            If testCollection IsNot Nothing Then
                dic_Controls = testCollection
            Else
                MsgBox("clsControls: CONTROL IS Nothing")
            End If
            NewFriendlyName = Replace(Me.TextboxName, "txt", "lbl")
            lblControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", NewFriendlyName)
            If lblControl IsNot Nothing Then
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", lblControl.Text & ":")
            Else
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", Me.TextboxName & ":")
            End If
            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlValue", NewValue)
        End If

    End Sub

    Private Sub comboAfterUpdate_change(sender As Object, e As System.EventArgs) Handles comboAfterUpdate.TextChanged
        Dim myComboBox As ComboBox
        Dim NewValue As String
        Dim CollectionKey As String
        Dim testCollection As Object = Nothing
        Dim NewFriendlyName As String
        Dim lblControl As Control

        If TypeOf (sender) Is System.Windows.Forms.ComboBox Then
            myComboBox = CType(sender, System.Windows.Forms.ComboBox)
            'MsgBox("clsControls = COMBO CHANGED - Control:" & myComboBox.Name & ", value= " & myComboBox.Text) 'YEA THIS WORKS !!!!!!
            Me.ComboBoxName = comboAfterUpdate.Name
            Me.ComboTAG = comboAfterUpdate.Tag
            Me.ComboValue = comboAfterUpdate.Text
            NewValue = Me.ComboValue
            CollectionKey = Me.ControlDeliveryDate & "_" & Me.ControlDeliveryRef & "_" & Me.ComboTAG
            testCollection = UpdateCollection(dic_Controls, CollectionKey, NewValue, Me.ControlDeliveryDate) 'VALUE is default
            If testCollection IsNot Nothing Then
                dic_Controls = testCollection
            Else
                MsgBox("clsControls: Combo IS Nothing")
            End If
            'If Me.ComboTAG = "40" Then
            ' testCollection = UpdateCollection(dic_Controls, CollectionKey, NewValue, Me.ControlDeliveryDate, "", "", "", "FLM_NAME") 'VALUE is default
            'If testCollection IsNot Nothing Then
            'dic_Controls = testCollection
            'Else
            'MsgBox("clsControls: Combo IS Nothing")
            'End If
            'End If
            NewFriendlyName = Replace(Me.ComboBoxName, "com", "lbl")
            lblControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", NewFriendlyName)
            If lblControl IsNot Nothing Then
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", lblControl.Text)
            Else
                frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", Me.TextboxName)
            End If
            frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlValue", NewValue)
            'lblControl = frmMainGIForm.FindFrameControls("GI_TIMESHEET", NewFriendlyName)
            'If lblControl IsNot Nothing Then
            'frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", lblControl.Text & ":")
            'Else
            'frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlName", Me.ComboBoxName & ":")
            'End If
            'frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET", "txtControlValue", NewValue)
        End If
    End Sub

    Private Sub comboAfterUpdate2_MouseWheel(sender As Object, e As MouseEventArgs) Handles comboAfterUpdate2.MouseWheel
        Dim HMEA As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)

        MsgBox("IN MOUSE WHEEL EVENT")
        HMEA.Handled = True
    End Sub

    Public Property TextboxName() As String
        Get
            Return _txtboxNameOfChangedValue
        End Get
        Set(value As String)
            _txtboxNameOfChangedValue = value
        End Set
    End Property

    Public Property ComboBoxName() As String
        Get
            Return _comboNameOfChangedValue
        End Get
        Set(value As String)
            _comboNameOfChangedValue = value
        End Set
    End Property

    Public Property TextBoxTAG() As String
        Get
            Return _txtboxTagOfChangedValue
        End Get
        Set(value As String)
            _txtboxTagOfChangedValue = value
        End Set
    End Property
    '_CtrlOBJECT

    Public Property ControlObject() As Object
        Get
            Return _CtrlOBJECT
        End Get
        Set(value As Object)
            _CtrlOBJECT = value
        End Set
    End Property

    Public Property ComboTAG() As String
        Get
            Return _comboTagOfChangedValue
        End Get
        Set(value As String)
            _comboTagOfChangedValue = value
        End Set
    End Property

    Public Property TextBoxValue() As String
        Get
            Return _txtboxChangedValue
        End Get
        Set(value As String)
            _txtboxChangedValue = value
        End Set
    End Property

    Public Property ComboValue() As String
        Get
            Return _comboChangedValue
        End Get
        Set(value As String)
            _comboChangedValue = value
        End Set
    End Property

    Public Property ControlID() As String
        Get
            Return _CtrlID
        End Get
        Set(value As String)
            _CtrlID = value
        End Set
    End Property

    Public Property ControlAltID() As String
        Get
            Return _CtrlAltID
        End Get
        Set(value As String)
            _CtrlAltID = value
        End Set
    End Property

    Public Property ControlASNID() As String
        Get
            Return _CtrlASNID
        End Get
        Set(value As String)
            _CtrlASNID = value
        End Set
    End Property

    Public Property ControlASNAltID() As String
        Get
            Return _CtrlASNAltID
        End Get
        Set(value As String)
            _CtrlASNAltID = value
        End Set
    End Property

    Public Property TheControl() As Control
        Get
            Return _CTRL
        End Get
        Set(value As Control)
            _CTRL = value
        End Set
    End Property

    Public Property ControlName() As String
        Get
            Return _CtrlName
        End Get
        Set(value As String)
            _CtrlName = value
        End Set
    End Property

    Public Property ControlFrameName() As String
        Get
            Return _CtrlFrameName
        End Get
        Set(value As String)
            _CtrlFrameName = value
        End Set
    End Property

    Public Property ControlTAG() As String
        Get
            Return _CtrlTag
        End Get
        Set(value As String)
            _CtrlTag = value
        End Set
    End Property

    Public Property ControlDate() As Date
        Get
            Return _CtrlDate
        End Get
        Set(value As Date)
            _CtrlDate = value
        End Set
    End Property

    Public Property ControlLeft() As Integer
        Get
            Return _CtrlLeft
        End Get
        Set(value As Integer)
            _CtrlLeft = value
        End Set
    End Property

    Public Property ControlTop() As Integer
        Get
            Return _CtrlTop
        End Get
        Set(value As Integer)
            _CtrlTop = value
        End Set
    End Property

    Public Property ControlWidth() As Integer
        Get
            Return _CtrlWidth
        End Get
        Set(value As Integer)
            _CtrlWidth = value
        End Set
    End Property

    Public Property ControlHeight() As Integer
        Get
            Return _CtrlHeight
        End Get
        Set(value As Integer)
            _CtrlHeight = value
        End Set
    End Property

    Public Property ControlType() As String
        Get
            Return _CtrlType
        End Get
        Set(value As String)
            _CtrlType = value
        End Set
    End Property

    Public Property ControlDeliveryDate() As Date
        Get
            Return _CtrlDeliveryDate
        End Get
        Set(value As Date)
            _CtrlDeliveryDate = value
        End Set
    End Property

    Public Property ControlDeliveryRef() As String
        Get
            Return _CtrlDeliveryRef
        End Get
        Set(value As String)
            _CtrlDeliveryRef = value
        End Set
    End Property

    Public Property ControlASN() As String
        Get
            Return _CtrlASN
        End Get
        Set(value As String)
            _CtrlASN = value
        End Set
    End Property

    Public Property ControlObjNumber() As Long
        Get
            Return _CtrlObjNumber
        End Get
        Set(value As Long)
            _CtrlObjNumber = value
        End Set
    End Property

    Public Property ControlValue() As String
        Get
            Return _CtrlValue
        End Get
        Set(value As String)
            _CtrlValue = value
        End Set
    End Property

    Public Property ControlList() As Object
        Get
            Return _CtrlList
        End Get
        Set(value As Object)
            _CtrlList = value
        End Set
    End Property

    Public Property ControlStartTAG() As String
        Get
            Return _CtrlStartTag
        End Get
        Set(value As String)
            _CtrlStartTag = value
        End Set
    End Property

    Public Property ControlEndTAG() As String
        Get
            Return _CtrlEndTag
        End Get
        Set(value As String)
            _CtrlEndTag = value
        End Set
    End Property

    Public Property ControlFrameRowNumber() As Long
        Get
            Return _CtrlFrameRowNumber
        End Get
        Set(value As Long)
            _CtrlFrameRowNumber = value
        End Set
    End Property

    Public Property ControlTotalRows() As Long
        Get
            Return _CtrlTotalRows
        End Get
        Set(value As Long)
            _CtrlTotalRows = value
        End Set
    End Property

    Public Property ControlFLMName() As String
        Get
            Return _CtrlFLMName
        End Get
        Set(value As String)
            _CtrlFLMName = value
        End Set
    End Property

    Public Property ControlFLMStartDateTime() As Date
        Get
            Return _CtrlFLMStartDateTime
        End Get
        Set(value As Date)
            _CtrlFLMStartDateTime = value
        End Set
    End Property

    Public Property ControlFLMEndDateTime() As Date
        Get
            Return _CtrlFLMEndDateTime
        End Get
        Set(value As Date)
            _CtrlFLMEndDateTime = value
        End Set
    End Property

    Public Property ControlOpName() As String
        Get
            Return _CtrlOpName
        End Get
        Set(value As String)
            _CtrlOpName = value
        End Set
    End Property

    Public Property ControlOpActivity() As String
        Get
            Return _CtrlOpActivity
        End Get
        Set(value As String)
            _CtrlOpActivity = value
        End Set
    End Property

    Public Property ControlOpStartDateTime() As Date
        Get
            Return _CtrlOpStartDateTime
        End Get
        Set(value As Date)
            _CtrlOpStartDateTime = value
        End Set
    End Property

    Public Property ControlOpEndDateTime() As Date
        Get
            Return _CtrlOpEndDateTime
        End Get
        Set(value As Date)
            _CtrlOpEndDateTime = value
        End Set
    End Property

    Public Property ControlOPdblTotalHrs() As Double
        Get
            Return _CtrlOPdblTotalHrs
        End Get
        Set(value As Double)
            _CtrlOPdblTotalHrs = value
        End Set
    End Property

    Public Property ControlOpdblTotalSecs() As Double
        Get
            Return _CtrlOpdblTotalSecs
        End Get
        Set(value As Double)
            _CtrlOpdblTotalSecs = value
        End Set
    End Property

    Public Property ControlOPstrTotalHrs() As String
        Get
            Return _CtrlOPstrTotalHrs
        End Get
        Set(value As String)
            _CtrlOPstrTotalHrs = value
        End Set
    End Property

    Public Property ControlFLMdblTotalHrs() As Double
        Get
            Return _CtrlFLMdblTotalHrs
        End Get
        Set(value As Double)
            _CtrlFLMdblTotalHrs = value
        End Set
    End Property

    Public Property ControlFLMstrTotalHrs() As String
        Get
            Return _CtrlFLMstrTotalHrs
        End Get
        Set(value As String)
            _CtrlFLMstrTotalHrs = value
        End Set
    End Property

    Public Property ControlPartNo() As String
        Get
            Return _CtrlPartNo
        End Get
        Set(value As String)
            _CtrlPartNo = value
        End Set
    End Property

    Public Property ControlQty() As String
        Get
            Return _CtrlQty
        End Get
        Set(value As String)
            _CtrlQty = value
        End Set
    End Property

    Public Property ControlExtraOrShort() As String
        Get
            Return _CtrlExtraShort
        End Get
        Set(value As String)
            _CtrlExtraShort = value
        End Set
    End Property

    Public Property ControlDBTable() As String
        Get
            Return _CtrlDBTable
        End Get
        Set(value As String)
            _CtrlDBTable = value
        End Set
    End Property

    Public Property ControlFieldName() As String
        Get
            Return _CtrlFieldName
        End Get
        Set(value As String)
            _CtrlFieldName = value
        End Set
    End Property

    Public Property ControlLastSaved() As Date
        Get
            Return _CtrlLastSaved
        End Get
        Set(value As Date)
            _CtrlLastSaved = value
        End Set
    End Property

    Public Property ControlBackColor() As Integer
        Get
            Return _CtrlBackColour
        End Get
        Set(value As Integer)
            _CtrlBackColour = value
        End Set
    End Property

    Public Property ControlLeftMargin() As Boolean
        Get
            Return _CtrlLeftMargin
        End Get
        Set(value As Boolean)
            _CtrlLeftMargin = value
        End Set
    End Property

    Public Property ControlErrMessage() As String
        Get
            Return _CtrlErrMessage
        End Get
        Set(value As String)
            _CtrlErrMessage = value
        End Set
    End Property

    Public Property ControlPrimaryKey() As Long
        Get
            Return _CtrlPrimaryKey
        End Get
        Set(value As Long)
            _CtrlPrimaryKey = value
        End Set
    End Property

    Public Property ControlFontName() As String
        Get
            Return _CtrlFontName
        End Get
        Set(value As String)
            _CtrlFontName = value
        End Set
    End Property

    Public Property ControlFontSize() As Integer
        Get
            Return _CtrlFontSize
        End Get
        Set(value As Integer)
            _CtrlFontSize = value
        End Set
    End Property

    Public Property ControlForeColor() As Integer
        Get
            Return _CtrlForeColour
        End Get
        Set(value As Integer)
            _CtrlForeColour = value
        End Set
    End Property

    Public Property ControlFontStyle() As String
        Get
            Return _CtrlFontStyle
        End Get
        Set(value As String)
            _CtrlFontStyle = value
        End Set
    End Property

    Public Property ControlOpRowNumber() As Long
        Get
            Return _CtrlOpRowNumber
        End Get
        Set(value As Long)
            _CtrlOpRowNumber = value
        End Set
    End Property

    Public Property ControlShortRowNumber() As Long
        Get
            Return _CtrlShortRowNumber
        End Get
        Set(value As Long)
            _CtrlShortRowNumber = value
        End Set
    End Property

    Public Property ControlExtraRowNumber() As Long
        Get
            Return _CtrlExtraRowNumber
        End Get
        Set(value As Long)
            _CtrlExtraRowNumber = value
        End Set
    End Property

    'Private _CtrlUpdatedByUsername As String
    'Private _CtrlUpdatedByEmpNo As String
    'Private _CtrlUpdatedByName As String

    Public Property ControlUpdatedByUsername() As String
        Get
            Return _CtrlUpdatedByUsername
        End Get
        Set(value As String)
            _CtrlUpdatedByUsername = value
        End Set
    End Property

    Public Property ControlUpdatedByEmpNo() As String
        Get
            Return _CtrlUpdatedByEmpNo
        End Get
        Set(value As String)
            _CtrlUpdatedByEmpNo = value
        End Set
    End Property

    Public Property ControlUpdatedByName() As String
        Get
            Return _CtrlUpdatedByName
        End Get
        Set(value As String)
            _CtrlUpdatedByName = value
        End Set
    End Property

    '_CtrlTabIndex

    Public Property ControlTabIndex() As Long
        Get
            Return _CtrlTabIndex
        End Get
        Set(value As Long)
            _CtrlTabIndex = value
        End Set
    End Property

    '_CtrlEnableControl

    Public Property ControlEnableControl() As Boolean
        Get
            Return _CtrlEnableControl
        End Get
        Set(value As Boolean)
            _CtrlEnableControl = value
        End Set
    End Property

End Class
