Module ControlsManager
    Public myControl As Control
    Public dic_Controls As Object
    Public Dic_DeliveryInfoFieldTypes As Object
    Public Dic_LabourFieldTypes As Object
    Public Dic_OpFieldTypes As Object
    Public Dic_SEFieldTypes As Object
    Public Dic_SCFieldTypes As Object
    Public Dic_FieldTypes As Object
    Public DeliveryInfoFieldnames As String
    Public LabourFieldnames As String
    Public OpFieldnames As String
    Public ShortAndExtraFieldnames As String
    Public SCFieldnames As String

    Sub Initialise()
        Dim tempTotals As New clsTotals
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim Type1 As String
        Dim Type2 As String
        Dim Type3 As String
        Dim Type4 As String
        Dim Type5 As String
        Dim ErrMessage As String

        dic_Controls = CreateObject("Scripting.Dictionary")
        dic_Controls.CompareMode = vbTextCompare
        dic_Totals = CreateObject("Scripting.Dictionary")
        dic_Totals.comparemode = vbTextCompare
        Dic_DeliveryInfoFieldTypes = CreateObject("Scripting.Dictionary")
        Dic_DeliveryInfoFieldTypes.comparemode = vbTextCompare
        Dic_LabourFieldTypes = CreateObject("Scripting.Dictionary")
        Dic_LabourFieldTypes.comparemode = vbTextCompare
        Dic_OpFieldTypes = CreateObject("Scripting.Dictionary")
        Dic_OpFieldTypes.comparemode = vbTextCompare
        Dic_SEFieldTypes = CreateObject("Scripting.Dictionary")
        Dic_SEFieldTypes.comparemode = vbTextCompare
        Dic_SCFieldTypes = CreateObject("Scripting.Dictionary")
        Dic_SCFieldTypes.comparemode = vbTextCompare
        Dic_FieldTypes = CreateObject("Scripting.Dictionary")
        Dic_FieldTypes.comparemode = vbTextCompare
        dic_AnyChanges = CreateObject("Scripting.Dictionary")
        dic_AnyChanges.comparemode = vbTextCompare
        dic_AnyChanges(frmMainGIForm.ControlPanelIdx) = "NO"

        tempTotals.Total_Operatives = 0
        tempTotals.Total_ShortParts = 0
        tempTotals.Total_ExtraParts = 0
        tempTotals.HighestOpTAGID = 43
        tempTotals.HighestOpBtnTAGID = 10
        tempTotals.HighestShortTAGID = 1001
        tempTotals.HighestExtraTAGID = 2001
        tempTotals.TotalOpHours = 0
        tempTotals.TotalFLMHours = 0

        strDeliveryDate = CDate("1970-01-01")
        strDeliveryRef = ""

        If Not dic_Totals.exists(strDeliveryDate & "_" & strDeliveryRef) Then
            dic_Totals(strDeliveryDate & "_" & strDeliveryRef) = tempTotals
        End If

        DeliveryInfoFieldnames = GetMyFields("tblDeliveryInfo", frmMainGIForm.myConnString, ErrMessage, "", Type1)
        SCFieldnames = GetMyFields("tblSupplierCompliance", frmMainGIForm.myConnString, ErrMessage, "", Type2)
        LabourFieldnames = GetMyFields("tblLabourHours", frmMainGIForm.myConnString, ErrMessage, "", Type3)
        OpFieldnames = GetMyFields("tblOperatives", frmMainGIForm.myConnString, ErrMessage, "", Type4)
        ShortAndExtraFieldnames = GetMyFields("tblshortsandextraparts", frmMainGIForm.myConnString, ErrMessage, "", Type5)

        Call Create_Dic_FieldTypes(Dic_DeliveryInfoFieldTypes, "tblDeliveryInfo", DeliveryInfoFieldnames, Type1)
        Call Create_Dic_FieldTypes(Dic_SCFieldTypes, "tblSupplierCompliance", SCFieldnames, Type2)
        Call Create_Dic_FieldTypes(Dic_LabourFieldTypes, "tblLabourHours", LabourFieldnames, Type3)
        Call Create_Dic_FieldTypes(Dic_OpFieldTypes, "tblOperatives", OpFieldnames, Type4)
        Call Create_Dic_FieldTypes(Dic_SEFieldTypes, "tblshortsandextraparts", ShortAndExtraFieldnames, Type5)

        Call Create_Dic_FieldTypes(Dic_FieldTypes, "tblDeliveryInfo", DeliveryInfoFieldnames, Type1)
        Call Create_Dic_FieldTypes(Dic_FieldTypes, "tblSupplierCompliance", SCFieldnames, Type2)
        Call Create_Dic_FieldTypes(Dic_FieldTypes, "tblLabourHours", LabourFieldnames, Type3)
        Call Create_Dic_FieldTypes(Dic_FieldTypes, "tblOperatives", OpFieldnames, Type4)
        Call Create_Dic_FieldTypes(Dic_FieldTypes, "tblshortsandextraparts", ShortAndExtraFieldnames, Type5)

    End Sub

    Sub AddFormControls(myForm As Form, ControlType As String, ControlName As String, InitialValue As String, TagNumber As String,
                        xx As Integer, yy As Integer, Width As Integer, Height As Integer)
        Dim tempControl As clsControls 'replaces mycontrol = OR tempControl = myControl

        If UCase(ControlType) = "TEXTBOX" Then
            myControl = New TextBox
            myControl.Name = ControlName
            myControl.Tag = TagNumber
            myControl.Left = xx
            myControl.Top = yy
            myControl.Width = Width
            myControl.Height = Height
            myControl.Text = InitialValue
            AddHandler myControl.TextChanged, AddressOf myControl_textchanged
            myForm.Controls.Add(myControl)

        End If
        If UCase(ControlType) = "BUTTON" Then
            myControl = New Button
            myControl.Name = ControlName
            myControl.Tag = TagNumber
            myControl.Left = xx
            myControl.Top = yy
            myControl.Width = Width
            myControl.Height = Height
            myControl.Text = InitialValue
            AddHandler myControl.Click, AddressOf myControl_Click
            myForm.Controls.Add(myControl)

        End If
        If UCase(ControlType) = "COMBOBOX" Then
            myControl = New ComboBox
            myControl.Name = ControlName
            myControl.Tag = TagNumber
            myControl.Left = xx
            myControl.Top = yy
            myControl.Width = Width
            myControl.Height = Height
            myControl.Text = InitialValue

            AddHandler myControl.TextChanged, AddressOf myControl_textchanged
            myForm.Controls.Add(myControl)

        End If
    End Sub

    Friend Sub myControl_textchanged(sender As Object, e As EventArgs)
        Dim txtbox As TextBox = DirectCast(sender, TextBox)

        MsgBox("Control Name: " & txtbox.Name & ", Value=" & txtbox.Text)
    End Sub

    Friend Sub myControl_Click(sender As Object, e As System.EventArgs)
        Dim btn As Button = DirectCast(sender, Button)

        'INSERT the TIME into the relevant control here:
        MsgBox("Control Name: " & btn.Name & ", Caption=" & btn.Text)
    End Sub

    Function AddNewControl(IsNewControl As Boolean, ScrollControlFrame As ScrollableControl, ControlDBTable As String, ControlFieldname As String, IDPrefix As String,
                                TheControl As Control, ControlName As String, ControlText As String, ControlType As String, ControlTAG As String, ControlTABIndex As Long,
        ControlDate As Date,
        ControlLeft As Integer, ControlTop As Integer, ControlWidth As Integer, ControlHeight As Integer,
        ControlDeliveryDate As Date, ControlDeliveryRef As String, controlUpdatedByEmpNo As String, Controlupdatedbyusername As String, controlupdatedbyname As String,
        Optional ControlASN As String = "", Optional ControlOBJCount As Long = 0,
        Optional ControlStartTAG As String = "", Optional ControlEndTAG As String = "",
        Optional ControlRowNumber As Long = 0, Optional ControlTotalRows As Long = 0, Optional MakeVisible As Boolean = True,
        Optional ByRef ListArray() As Object = Nothing, Optional ControlBACKCOLOUR As Integer = 0, Optional ControlAddLeftMargin As Boolean = True,
        Optional ControlFontName As String = "TAHOMA", Optional ControlFontSize As Integer = 10, Optional Fontstyle As String = "", Optional ControlForecolor As Integer = 0,
        Optional TextAlign As String = "LEFT", Optional LabelAlign As String = "MIDDLE CENTER", Optional MakeReadOnly As Boolean = False,
        Optional AddComboList As Boolean = False, Optional ByVal ControlStartDateTime As Date = #1970-01-01#, Optional ByVal ControlEndDateTime As Date = #1970-01-01#,
                           Optional ControlPrimaryKey As Long = 0) As Long

        Dim tempControl As New clsControls
        Dim NewControl As Control = Nothing
        Dim NewOBJ As Control
        Dim NewCombo As New ComboBox
        Dim NewButton As New Button
        Dim NewTextBox As New TextBox
        Dim NewLabel As New Label
        Dim IDX As Long
        Dim CollectionKey As Object
        Dim ControlFrameName As String
        'Dim font = New Font(tb_remark.Font, FontStyle.Italic)
        Dim NewFont As Font

        'On Error GoTo Err_AddNewControl
        If ScrollControlFrame IsNot Nothing Then
            ControlFrameName = ScrollControlFrame.Name
        Else
            ControlFrameName = "NA"
        End If
        'Set Dic_Collection = CreateObject("Scripting.Dictionary")
        'Dic_Collection.RemoveAll
        NewOBJ = Nothing
        If IsNewControl Then

            If UCase(ControlType) = "COMBOBOX" Or InStr(UCase(ControlType), "COMBOBOX") > 0 Then
                NewCombo = New ComboBox

                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.ComboBox.1", ControlName, MakeVisible)
                NewCombo.Name = ControlName
                NewCombo.Top = ControlTop
                NewCombo.Left = ControlLeft
                NewCombo.Width = ControlWidth
                NewCombo.AutoSize = False
                NewCombo.TabIndex = ControlTABIndex
                Select Case UCase(Fontstyle)
                    Case "BOLD"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewCombo.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                NewCombo.Tag = ControlTAG
                NewCombo.Text = ControlText
                NewCombo.ForeColor = ColorTranslator.FromWin32(ControlForecolor)
                NewCombo.BackColor = ColorTranslator.FromWin32(ControlBACKCOLOUR)
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If ControlHeight > 0 Then
                    NewCombo.Height = ControlHeight
                End If
                NewCombo.Items.Clear()

                If Not IsNothing(ListArray) Then
                    For IDX = 0 To UBound(ListArray)
                        If Len(ListArray(IDX)) > 0 Then
                            NewCombo.Items.Add(ListArray(IDX))
                        End If
                    Next
                End If
                'NewControl = DirectCast(NewOBJ, ComboBox)
                'NewControl = NewOBJ
                'AddHandler NewCombo.TextChanged, AddressOf NewCombo_textchanged

                tempControl.comboAfterUpdate = NewCombo
                tempControl.comboAfterUpdate2 = NewCombo
                NewOBJ = DirectCast(NewCombo, System.Windows.Forms.ComboBox)
            End If


            If UCase(ControlType) = "BTN" Or InStr(UCase(ControlType), "BTN") > 0 Or InStr(UCase(ControlType), "BUTTON") > 0 Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.CommandButton.1", ControlName, MakeVisible)
                'NewOBJ = New Button
                NewButton.Name = ControlName
                NewButton.Top = ControlTop
                NewButton.Left = ControlLeft
                NewButton.Width = ControlWidth
                NewButton.Tag = ControlTAG

                NewButton.Text = ControlText
                NewButton.ForeColor = ColorTranslator.FromWin32(ControlForecolor)
                NewButton.BackColor = ColorTranslator.FromWin32(ControlBACKCOLOUR)
                NewButton.TabIndex = ControlTABIndex

                'can also use .FromName()
                'NewOBJ.BackColor = ControlBACKCOLOUR
                Select Case UCase(Fontstyle)
                    Case "BOLD"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewButton.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Regular)

                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                'NewCtrl.SelectionMargin = ControlAddLeftMargin 'NOT applicable to BUttons !

                If ControlHeight > 0 Then
                    NewButton.Height = ControlHeight
                End If
                If MakeReadOnly Then
                    NewButton.Enabled = False
                End If
                'AddHandler NewButton.Click, AddressOf Newbutton_click
                If InStr(1, UCase(NewButton.Name), UCase("StartTime"), vbTextCompare) > 0 Then
                    tempControl.cbTimeStartEvent = NewButton
                End If
                If InStr(1, UCase(NewButton.Name), UCase("FinishTime"), vbTextCompare) > 0 Then
                    tempControl.cbTimeEndEvent = NewButton
                End If
                'If InStr(1, UCase(NewButton.Name), UCase("btnOperativeStartTime"), vbTextCompare) > 0 Then
                'tempControl.cbTimeStartEvent = NewButton
                'End If
                'If InStr(1, UCase(NewButton.Name), UCase("btnOperativeFinishTime"), vbTextCompare) > 0 Then
                'tempControl.cbTimeEndEvent = NewButton
                'End If
                NewOBJ = DirectCast(NewButton, System.Windows.Forms.Button)
                'ScrollControlFrame.Controls.Add(NewButton)
                'If InStr(1, UCase(NewOBJ.Name), "START", vbTextCompare) > 0 Then
                'AddHandler NewOBJ.Click, AddressOf NewOBJ_click
                'tempControl.cbTimeStartEvent = NewOBJ
                'End If

                'If InStr(1, UCase(NewOBJ.Name), "END", vbTextCompare) > 0 Then
                'AddHandler NewOBJ.Click, AddressOf NewOBJ_click
                'tempControl.cbTimeEndEvent = NewOBJ
                'End If
            End If

            If UCase(ControlType) = "LABEL" Or InStr(UCase(ControlType), "LABEL") > 0 Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
                'NewOBJ = New TextBox

                NewLabel.Name = ControlName
                NewLabel.Top = ControlTop
                NewLabel.Left = ControlLeft
                NewLabel.Width = ControlWidth
                NewLabel.AutoSize = False
                'NewTextBox.TextAlign = HorizontalAlignment.Center/.left/.right
                Select Case UCase(Fontstyle)
                    Case "BOLD"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewLabel.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                NewLabel.Tag = ControlTAG
                NewLabel.Text = ControlText

                If UCase(LabelAlign) = "BOTTOM LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.BottomLeft
                ElseIf UCase(LabelAlign) = "BOTTOM CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.BottomCenter
                ElseIf UCase(LabelAlign) = "BOTTOM RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.BottomRight
                ElseIf UCase(LabelAlign) = "MIDDLE LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleLeft
                ElseIf UCase(LabelAlign) = "MIDDLE CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleCenter
                ElseIf UCase(LabelAlign) = "MIDDLE RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleRight
                ElseIf UCase(LabelAlign) = "TOP LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.TopLeft
                ElseIf UCase(LabelAlign) = "TOP CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.TopCenter
                ElseIf UCase(LabelAlign) = "TOP RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.TopRight
                Else
                    NewLabel.TextAlign = ContentAlignment.MiddleCenter
                End If


                NewLabel.ForeColor = ColorTranslator.FromWin32(ControlForecolor)
                NewLabel.BackColor = ColorTranslator.FromWin32(ControlBACKCOLOUR)
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If ControlHeight > 0 Then
                    NewLabel.Height = ControlHeight
                End If
                'txtBoxAfterUpdate_change in clsControls
                'tempControl.txtBoxAfterUpdate = NewTextBox
                'AddHandler NewTextBox.TextChanged, AddressOf NewTextBox_textchanged

                NewOBJ = DirectCast(NewLabel, System.Windows.Forms.Label)
            End If

            If UCase(ControlType) = "TEXTBOX" Or InStr(UCase(ControlType), "TEXTBOX") > 0 Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
                'NewOBJ = New TextBox

                NewTextBox.Name = ControlName
                NewTextBox.Top = ControlTop
                NewTextBox.Left = ControlLeft
                NewTextBox.Width = ControlWidth
                NewTextBox.AutoSize = False
                NewTextBox.TabIndex = ControlTABIndex
                'NewTextBox.TextAlign = HorizontalAlignment.Center/.left/.right
                Select Case UCase(Fontstyle)
                    Case "BOLD"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewTextBox.Font = New Font(ControlFontName, ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                NewTextBox.Tag = ControlTAG
                NewTextBox.Text = ControlText
                NewTextBox.ForeColor = ColorTranslator.FromWin32(ControlForecolor)
                NewTextBox.BackColor = ColorTranslator.FromWin32(ControlBACKCOLOUR)

                If UCase(TextAlign) = "LEFT" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Left
                ElseIf UCase(TextAlign) = "CENTER" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Center
                ElseIf UCase(TextAlign) = "RIGHT" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Right
                Else
                    NewTextBox.TextAlign = HorizontalAlignment.Left
                End If
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If ControlHeight > 0 Then
                    NewTextBox.Height = ControlHeight
                End If
                'txtBoxAfterUpdate_change in clsControls
                tempControl.txtBoxAfterUpdate = NewTextBox
                'AddHandler NewTextBox.TextChanged, AddressOf NewTextBox_textchanged

                If MakeReadOnly Then
                    NewTextBox.ReadOnly = True
                Else
                    NewTextBox.ReadOnly = False
                End If

                NewOBJ = DirectCast(NewTextBox, System.Windows.Forms.TextBox)

            End If
            ScrollControlFrame.Controls.Add(NewOBJ)
        Else
            'FIXED CONTROLS done at Design-time:
            'Use Found Textbox - control passed through
            If UCase(ControlType) = "TEXTBOX" Or InStr(UCase(ControlType), "TEXTBOX") > 0 Then
                NewTextBox = New TextBox
                If TypeName(TheControl) = "TextBox" Then
                    NewTextBox = TheControl
                    tempControl.txtBoxAfterUpdate = NewTextBox
                    NewOBJ = DirectCast(NewTextBox, System.Windows.Forms.TextBox)
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - TEXTBOX being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            ElseIf UCase(ControlType) = "LABEL" Or InStr(UCase(ControlType), "LABEL") > 0 Then
                NewLabel = New Label
                If TypeName(TheControl) = "Label" Then
                    NewLabel = TheControl
                    NewOBJ = DirectCast(NewLabel, System.Windows.Forms.Label)
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - LABEL being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            ElseIf UCase(ControlType) = "BTN" Or InStr(UCase(ControlType), "BTN") > 0 Or InStr(UCase(ControlType), "BUTTON") > 0 Then
                NewButton = New Button
                If TypeName(TheControl) = "Button" Then
                    NewButton = TheControl

                    If InStr(1, UCase(NewButton.Name), UCase("StartTime"), vbTextCompare) > 0 Then
                        tempControl.cbTimeStartEvent = NewButton
                    End If
                    If InStr(1, UCase(NewButton.Name), UCase("EndTime"), vbTextCompare) > 0 Or InStr(1, UCase(NewButton.Name), UCase("FinishTime"), vbTextCompare) > 0 Then
                        tempControl.cbTimeEndEvent = NewButton
                    End If

                    NewOBJ = DirectCast(NewButton, System.Windows.Forms.Button)
                    'Need a static button event in clsControls:
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - BUTTON being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If

            ElseIf UCase(ControlType) = "COMBOBOX" Or InStr(UCase(ControlType), "COMBOBOX") > 0 Then
                NewCombo = New ComboBox
                If TypeName(TheControl) = "ComboBox" Then
                    NewCombo = TheControl
                    If AddComboList Then
                        NewCombo.Items.Clear()

                        If Not IsNothing(ListArray) Then
                            For IDX = 0 To UBound(ListArray)
                                If Len(ListArray(IDX)) > 0 Then
                                    NewCombo.Items.Add(ListArray(IDX))
                                End If
                            Next
                        End If
                    End If
                    'AddHandler NewCombo.TextChanged, AddressOf NewCombo_textchanged
                    tempControl.comboAfterUpdate = NewCombo
                    NewOBJ = DirectCast(NewCombo, System.Windows.Forms.ComboBox)
                Else
                    'Control of wrong type here:
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - TEXTBOX being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            End If
        End If

        If TheControl Is Nothing Then
            tempControl.TheControl = NewOBJ
        Else
            tempControl.TheControl = TheControl
        End If
        tempControl.ControlName = ControlName
        tempControl.ControlFrameName = ControlFrameName
        tempControl.ControlDBTable = ControlDBTable
        tempControl.ControlFieldName = ControlFieldname
        tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
        tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
        tempControl.ControlValue = ControlText
        tempControl.ControlType = ControlType
        tempControl.ControlTAG = ControlTAG
        tempControl.ControlTabIndex = ControlTABIndex
        tempControl.ControlDate = ControlDate
        tempControl.ControlLeft = ControlLeft
        tempControl.ControlTop = ControlTop
        tempControl.ControlWidth = ControlWidth
        tempControl.ControlHeight = ControlHeight
        tempControl.ControlDeliveryDate = ControlDeliveryDate
        tempControl.ControlDeliveryRef = ControlDeliveryRef
        tempControl.ControlASN = ControlASN
        tempControl.ControlObjNumber = ControlOBJCount
        tempControl.ControlStartTAG = ControlStartTAG
        tempControl.ControlEndTAG = ControlEndTAG
        tempControl.ControlFrameRowNumber = ControlRowNumber
        tempControl.ControlTotalRows = ControlTotalRows
        tempControl.ControlBackColor = ControlBACKCOLOUR
        tempControl.ControlLeftMargin = ControlAddLeftMargin
        tempControl.ControlFontName = ControlFontName
        tempControl.ControlFontSize = ControlFontSize
        tempControl.ControlUpdatedByEmpNo = controlUpdatedByEmpNo
        tempControl.ControlUpdatedByUsername = Controlupdatedbyusername
        tempControl.ControlUpdatedByName = controlupdatedbyname
        If ControlPrimaryKey > 0 Then
            tempControl.ControlPrimaryKey = ControlPrimaryKey
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_FLMDETAILS") And UCase(ControlName) = UCase("txtFLMStartTime") Then
            tempControl.ControlFLMStartDateTime = ControlStartDateTime
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_FLMDETAILS") And UCase(ControlName) = UCase("txtFLMFinishTime") Then
            tempControl.ControlFLMEndDateTime = ControlEndDateTime
        End If 'FRAME_OPERATIVES not being used now. 03-JUL-2019
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_OPERATIVES") And UCase(ControlName) = UCase("txtOperativeStartTime:" & CStr(ControlRowNumber)) Then
            tempControl.ControlOpStartDateTime = ControlStartDateTime
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_OPERATIVES") And UCase(ControlName) = UCase("txtOperativeFinishTime:" & CStr(ControlRowNumber)) Then
            tempControl.ControlOpEndDateTime = ControlEndDateTime
        End If

        CollectionKey = tempControl.ControlTAG
        If Len(ControlDeliveryDate) > 0 Then
            CollectionKey = ControlDeliveryDate.ToString("dd/MM/yyyy") & "_" & ControlDeliveryRef & "_" & tempControl.ControlTAG
        ElseIf Len(ControlASN) > 0 Then
            CollectionKey = ControlDeliveryDate.ToString("dd/MM/yyyy") & "_" & ControlASN & "_" & tempControl.ControlTAG
        End If
        'ctrlCollection.Add tempControl, tempControl.ControlTag
        'If Not InCollection("MISSING", ctrlCollection, CollectionKey) Then
        'ctrlCollection.Add tempControl, CollectionKey
        'End If
        If Not dic_Controls.Exists(CollectionKey) Then
            dic_Controls(CollectionKey) = tempControl
        Else
            dic_Controls(CollectionKey) = tempControl
        End If

        'Write_Properties_To_File("Property Log.log", tempControl)

        AddNewControl = AddNewControl + 1

        'Call Error_Report("AddNewControl")

    End Function

    Function AddNewControls(IsNewControl As Boolean, ScrollControlFrame As ScrollableControl, TheControl As Control, ctrlProperty As clsControls,
        Optional ByVal MakeVisible As Boolean = True, Optional ByRef ListArray() As Object = Nothing, Optional TextAlign As String = "LEFT",
        Optional LabelAlign As String = "MIDDLE CENTER", Optional MakeReadOnly As Boolean = False, Optional AddComboList As Boolean = False,
        Optional ByVal ControlStartDateTime As Date = #1970-01-01#, Optional ByVal ControlEndDateTime As Date = #1970-01-01#) As Long

        Dim tempControl As New clsControls
        Dim NewControl As Control = Nothing
        Dim NewOBJ As Control
        Dim NewCombo As New ComboBox
        Dim NewButton As New Button
        Dim NewTextBox As New TextBox
        Dim NewLabel As New Label
        Dim IDX As Long
        Dim CollectionKey As Object
        Dim ControlFrameName As String
        'Dim font = New Font(tb_remark.Font, FontStyle.Italic)
        Dim NewFont As Font

        tempControl = ctrlProperty
        'On Error GoTo Err_AddNewControl
        If ScrollControlFrame IsNot Nothing Then
            ControlFrameName = ScrollControlFrame.Name
        Else
            ControlFrameName = "NA"
        End If
        'Set Dic_Collection = CreateObject("Scripting.Dictionary")
        'Dic_Collection.RemoveAll
        NewOBJ = Nothing
        If IsNewControl Then

            If UCase(tempControl.ControlType) = "COMBOBOX" Then
                NewCombo = New ComboBox

                'NewOBJ = New ComboBox
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.ComboBox.1", ControlName, MakeVisible)
                NewCombo.Name = tempControl.ControlName
                NewCombo.Top = tempControl.ControlTop
                NewCombo.Left = tempControl.ControlLeft
                NewCombo.Width = tempControl.ControlWidth
                NewCombo.AutoSize = False

                Select Case UCase(tempControl.ControlFontStyle)
                    Case "BOLD"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewCombo.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                NewCombo.Tag = tempControl.ControlTAG
                NewCombo.Text = tempControl.ControlValue
                NewCombo.ForeColor = ColorTranslator.FromWin32(tempControl.ControlForeColor)
                NewCombo.BackColor = ColorTranslator.FromWin32(tempControl.ControlBackColor)
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If tempControl.ControlHeight > 0 Then
                    NewCombo.Height = tempControl.ControlHeight
                End If
                NewCombo.Items.Clear()

                If Not IsNothing(ListArray) Then
                    For IDX = 0 To UBound(ListArray)
                        If Len(ListArray(IDX)) > 0 Then
                            NewCombo.Items.Add(ListArray(IDX))
                        End If
                    Next
                End If
                'NewControl = DirectCast(NewOBJ, ComboBox)
                'NewControl = NewOBJ
                'AddHandler NewCombo.TextChanged, AddressOf NewCombo_textchanged

                tempControl.comboAfterUpdate = NewCombo
                NewOBJ = DirectCast(NewCombo, System.Windows.Forms.ComboBox)
            End If


            If UCase(tempControl.ControlType) = "BTN" Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.CommandButton.1", ControlName, MakeVisible)
                'NewOBJ = New Button
                NewButton.Name = tempControl.ControlName
                NewButton.Top = tempControl.ControlTop
                NewButton.Left = tempControl.ControlLeft
                NewButton.Width = tempControl.ControlWidth
                NewButton.Tag = tempControl.ControlTAG

                NewButton.Text = tempControl.ControlValue
                NewButton.ForeColor = ColorTranslator.FromWin32(tempControl.ControlForeColor)
                NewButton.BackColor = ColorTranslator.FromWin32(tempControl.ControlBackColor)

                'can also use .FromName()
                'NewOBJ.BackColor = ControlBACKCOLOUR
                Select Case UCase(tempControl.ControlFontStyle)
                    Case "BOLD"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewButton.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Regular)

                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                'NewCtrl.SelectionMargin = ControlAddLeftMargin 'NOT applicable to BUttons !

                If tempControl.ControlHeight > 0 Then
                    NewButton.Height = tempControl.ControlHeight
                End If
                'AddHandler NewButton.Click, AddressOf Newbutton_click
                If InStr(1, UCase(NewButton.Name), UCase("StartTime"), vbTextCompare) > 0 Then
                    tempControl.cbTimeStartEvent = NewButton
                End If
                If InStr(1, UCase(NewButton.Name), UCase("FinishTime"), vbTextCompare) > 0 Then
                    tempControl.cbTimeEndEvent = NewButton
                End If
                'If InStr(1, UCase(NewButton.Name), UCase("btnOperativeStartTime"), vbTextCompare) > 0 Then
                'tempControl.cbTimeStartEvent = NewButton
                'End If
                'If InStr(1, UCase(NewButton.Name), UCase("btnOperativeFinishTime"), vbTextCompare) > 0 Then
                'tempControl.cbTimeEndEvent = NewButton
                'End If
                NewOBJ = DirectCast(NewButton, System.Windows.Forms.Button)
                'ScrollControlFrame.Controls.Add(NewButton)
                'If InStr(1, UCase(NewOBJ.Name), "START", vbTextCompare) > 0 Then
                'AddHandler NewOBJ.Click, AddressOf NewOBJ_click
                'tempControl.cbTimeStartEvent = NewOBJ
                'End If

                'If InStr(1, UCase(NewOBJ.Name), "END", vbTextCompare) > 0 Then
                'AddHandler NewOBJ.Click, AddressOf NewOBJ_click
                'tempControl.cbTimeEndEvent = NewOBJ
                'End If
            End If

            If UCase(tempControl.ControlType) = "LABEL" Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
                'NewOBJ = New TextBox

                NewLabel.Name = tempControl.ControlName
                NewLabel.Top = tempControl.ControlTop
                NewLabel.Left = tempControl.ControlLeft
                NewLabel.Width = tempControl.ControlWidth
                NewLabel.AutoSize = False
                'NewTextBox.TextAlign = HorizontalAlignment.Center/.left/.right
                Select Case UCase(tempControl.ControlFontStyle)
                    Case "BOLD"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewLabel.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                NewLabel.Tag = tempControl.ControlTAG
                NewLabel.Text = tempControl.ControlValue

                If UCase(LabelAlign) = "BOTTOM LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.BottomLeft
                ElseIf UCase(LabelAlign) = "BOTTOM CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.BottomCenter
                ElseIf UCase(LabelAlign) = "BOTTOM RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.BottomRight
                ElseIf UCase(LabelAlign) = "MIDDLE LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleLeft
                ElseIf UCase(LabelAlign) = "MIDDLE CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleCenter
                ElseIf UCase(LabelAlign) = "MIDDLE RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.MiddleRight
                ElseIf UCase(LabelAlign) = "TOP LEFT" Then
                    NewLabel.TextAlign = ContentAlignment.TopLeft
                ElseIf UCase(LabelAlign) = "TOP CENTER" Then
                    NewLabel.TextAlign = ContentAlignment.TopCenter
                ElseIf UCase(LabelAlign) = "TOP RIGHT" Then
                    NewLabel.TextAlign = ContentAlignment.TopRight
                Else
                    NewLabel.TextAlign = ContentAlignment.MiddleCenter
                End If


                NewLabel.ForeColor = ColorTranslator.FromWin32(tempControl.ControlForeColor)
                NewLabel.BackColor = ColorTranslator.FromWin32(tempControl.ControlBackColor)
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If tempControl.ControlHeight > 0 Then
                    NewLabel.Height = tempControl.ControlHeight
                End If
                'txtBoxAfterUpdate_change in clsControls
                'tempControl.txtBoxAfterUpdate = NewTextBox
                'AddHandler NewTextBox.TextChanged, AddressOf NewTextBox_textchanged

                NewOBJ = DirectCast(NewLabel, System.Windows.Forms.Label)
            End If

            If UCase(tempControl.ControlType) = "TEXTBOX" Then
                'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
                'NewOBJ = New TextBox

                NewTextBox.Name = tempControl.ControlName
                NewTextBox.Top = tempControl.ControlTop
                NewTextBox.Left = tempControl.ControlLeft
                NewTextBox.Width = tempControl.ControlWidth
                NewTextBox.AutoSize = False
                'NewTextBox.TextAlign = HorizontalAlignment.Center/.left/.right
                Select Case UCase(tempControl.ControlFontStyle)
                    Case "BOLD"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Bold)
                    Case "ITALIC"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic)
                    Case "STRIKEOUT"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Strikeout)
                    Case "UNDERLINE"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline)
                    Case "BOLD,UNDERLINE"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "UNDERLINE,BOLD"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)
                    Case "BOLD,ITALIC"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case "ITALIC,BOLD"
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Italic Or Drawing.FontStyle.Bold)
                    Case Else
                        NewTextBox.Font = New Font(tempControl.ControlFontName, tempControl.ControlFontSize, Drawing.FontStyle.Regular)
                End Select
                'NewOBJ.Font = New Font(ControlFontName, ControlFontSize)
                NewTextBox.Tag = tempControl.ControlTAG
                NewTextBox.Text = tempControl.ControlValue
                NewTextBox.ForeColor = ColorTranslator.FromWin32(tempControl.ControlForeColor)
                NewTextBox.BackColor = ColorTranslator.FromWin32(tempControl.ControlBackColor)

                If UCase(TextAlign) = "LEFT" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Left
                ElseIf UCase(TextAlign) = "CENTER" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Center
                ElseIf UCase(TextAlign) = "RIGHT" Then
                    NewTextBox.TextAlign = HorizontalAlignment.Right
                Else
                    NewTextBox.TextAlign = HorizontalAlignment.Left
                End If
                'NewCtrl.ControlAddLeftMargin does NOT exist in VB.NET - but padding does.
                If tempControl.ControlHeight > 0 Then
                    NewTextBox.Height = tempControl.ControlHeight
                End If
                'txtBoxAfterUpdate_change in clsControls
                tempControl.txtBoxAfterUpdate = NewTextBox
                'AddHandler NewTextBox.TextChanged, AddressOf NewTextBox_textchanged

                If MakeReadOnly Then
                    NewTextBox.ReadOnly = True
                Else
                    NewTextBox.ReadOnly = False
                End If

                NewOBJ = DirectCast(NewTextBox, System.Windows.Forms.TextBox)

            End If
            ScrollControlFrame.Controls.Add(NewOBJ)
        Else
            'FIXED CONTROLS done at Design-time:
            'Use Found Textbox - control passed through
            If UCase(tempControl.ControlType) = "TEXTBOX" Then
                NewTextBox = New TextBox
                If TypeName(TheControl) = "TextBox" Then
                    NewTextBox = TheControl
                    tempControl.txtBoxAfterUpdate = NewTextBox
                    NewOBJ = DirectCast(NewTextBox, System.Windows.Forms.TextBox)
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - TEXTBOX being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            ElseIf UCase(tempControl.ControlType) = "LABEL" Then
                NewLabel = New Label
                If TypeName(TheControl) = "Label" Then
                    NewLabel = TheControl
                    NewOBJ = DirectCast(NewLabel, System.Windows.Forms.Label)
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - LABEL being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            ElseIf UCase(tempControl.ControlType) = "BTN" Then
                NewButton = New Button
                If TypeName(TheControl) = "Button" Then
                    NewButton = TheControl
                    NewOBJ = DirectCast(NewButton, System.Windows.Forms.Button)
                    'Need a static button event in clsControls:
                Else
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - BUTTON being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If

            ElseIf UCase(tempControl.ControlType) = "COMBOBOX" Then
                NewCombo = New ComboBox
                If TypeName(TheControl) = "ComboBox" Then
                    NewCombo = TheControl
                    If AddComboList Then
                        NewCombo.Items.Clear()

                        If Not IsNothing(ListArray) Then
                            For IDX = 0 To UBound(ListArray)
                                If Len(ListArray(IDX)) > 0 Then
                                    NewCombo.Items.Add(ListArray(IDX))
                                End If
                            Next
                        End If
                    End If
                    'AddHandler NewCombo.TextChanged, AddressOf NewCombo_textchanged
                    tempControl.comboAfterUpdate = NewCombo
                    NewOBJ = DirectCast(NewCombo, System.Windows.Forms.ComboBox)
                Else
                    'Control of wrong type here:
                    frmMainGIForm.logger.LogError("GI_Errors_v1_3.log", Application.StartupPath, "Wrong Type - TEXTBOX being cast to:" & TypeName(TheControl), "AddNewControl()", frmMainGIForm.GetComputerName() & "," & frmMainGIForm.GetIPv4Address() & "," & frmMainGIForm.GetIPv6Address() & ", OPERATOR LOGGED OUT:" & frmMainGIForm.myUsername)
                End If
            End If
        End If

        tempControl.TheControl = NewOBJ
        tempControl.ControlFrameName = ControlFrameName
        tempControl.ControlID = CStr(tempControl.ControlDeliveryDate) & "_" & tempControl.ControlDeliveryRef & "_" & tempControl.ControlTAG
        tempControl.ControlAltID = CStr(tempControl.ControlDeliveryDate) & "_" & tempControl.ControlDeliveryRef & "_" & tempControl.ControlName
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_FLMDETAILS") And UCase(tempControl.ControlName) = UCase("txtFLMStartTime") Then
            tempControl.ControlFLMStartDateTime = ControlStartDateTime
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_FLMDETAILS") And UCase(tempControl.ControlName) = UCase("txtFLMFinishTime") Then
            tempControl.ControlFLMEndDateTime = ControlEndDateTime
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_OPERATIVES") And UCase(tempControl.ControlName) = UCase("txtOperativeStartTime:" & CStr(tempControl.ControlFrameRowNumber)) Then
            tempControl.ControlOpStartDateTime = ControlStartDateTime
        End If
        If UCase(ScrollControlFrame.Name) = UCase("FRAME_OPERATIVES") And UCase(tempControl.ControlName) = UCase("txtOperativeFinishTime:" & CStr(tempControl.ControlFrameRowNumber)) Then
            tempControl.ControlOpEndDateTime = ControlEndDateTime
        End If

        CollectionKey = tempControl.ControlTAG
        If Len(tempControl.ControlDeliveryDate) > 0 Then
            CollectionKey = tempControl.ControlDeliveryDate & "_" & tempControl.ControlDeliveryRef & "_" & tempControl.ControlTAG
        ElseIf Len(tempControl.ControlASN) > 0 Then
            CollectionKey = tempControl.ControlDeliveryDate & "_" & tempControl.ControlASN & "_" & tempControl.ControlTAG
        End If
        'ctrlCollection.Add tempControl, tempControl.ControlTag
        'If Not InCollection("MISSING", ctrlCollection, CollectionKey) Then
        'ctrlCollection.Add tempControl, CollectionKey
        'End If
        If Not dic_Controls.Exists(CollectionKey) Then
            dic_Controls(CollectionKey) = tempControl
        Else
            dic_Controls(CollectionKey) = tempControl
        End If

        Return AddNewControls + 1

        'Call Error_Report("AddNewControls")

    End Function

    Sub InsertFLMButtons(SaveFLMDetails As Boolean, ScrollFrame As ScrollableControl, strDeliveryDate As String, strDeliveryRef As String, strASNNo As String,
                         Optional ByVal EmployeeNo As String = "", Optional FLMName As String = "", Optional ByVal dtStartTime As Date = #1970-01-01#, Optional ByVal dtFinishTime As Date = #1970-01-01#,
                         Optional ByRef StartTABIndex As Long = 0, Optional DisableStartTimeButton As Boolean = False, Optional DisableFinishTimeButton As Boolean = False)
        Dim FLMStartButton As New Button
        Dim FLMEndButton As New Button
        Dim ctrls() As Control
        Dim btnFLMstarttime As Control
        Dim btnFLMendtime As Control
        Dim FLMsControl As String
        Dim FLMeControl As String
        Dim tempControl As New clsControls
        Dim ControlLeft As Integer
        Dim ControlTop As Integer
        Dim ControlWidth As Integer
        Dim ControlHeight As Integer
        Dim ControlTAG As String
        Dim ComboArray() As String
        Dim ComboIDX As Long
        Dim Firstname As String = ""
        Dim Lastname As String = ""
        Dim Fullname As String = ""
        Dim ListArray As Object
        Dim ErrMessage As String = ""
        Dim ControlType As String = ""
        Dim ControlText As String = ""
        Dim ControlStyle As String = ""
        Dim ControlFontName = "Cambria"
        Dim ControlFontSize As Integer = 9
        Dim ControlForeColor As Integer = RGB(0, 0, 0) 'black
        Dim ControlBackColor As Integer = RGB(240, 248, 255) 'ALICEBLUE
        Dim ControlFieldname As String
        Dim FieldnameArr() As String
        Dim ControlDBTable As String
        Dim ControlDate As DateTime
        Dim ControlDeliveryDate As DateTime
        Dim ControlDeliveryRef As String
        Dim ControlASN As String
        Dim ControlOBJCount As Long
        Dim ControlStartTAG As String
        Dim ControlEndTAG As String
        Dim ControlRowNumber As Long
        Dim ControlTotalRows As Long
        Dim MakeVisible As Boolean = True
        Dim ControlLeftMargin As Boolean = False
        Dim Criteria As String = ""
        Dim strSQL As String = ""
        Dim LoadedOK As Boolean = False
        Dim LabourFieldsArr() As String = Nothing
        Dim SearchText As String = ""
        Dim SearchField As String
        Dim FieldType As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim SearchCriteria As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ExcludeFields As String
        Dim AllValues() As Object = Nothing
        Dim AllFields() As String = Nothing
        Dim LabourFields As String
        Dim LabourFieldValues As String
        Dim SavedOK As Boolean = False
        Dim SaveCriteria As String
        Dim dtDeliveryDate As DateTime
        Dim dtNow As DateTime
        Dim RecID As String
        Dim ExtractTotals As New clsTotals
        Dim dtLastSaved As DateTime
        Dim strLastSaved As String
        Dim SaveDeliveryDate As String
        Dim ControlTABIndex As Long
        Dim UpdatedByName As String
        Dim UpdatedByEmpno As String
        Dim UpdatedByUsername As String

        'EMPLOYEE NO TEXTBOX:

        ReDim FieldnameArr(1)

        'ADD LinkID from ControlPrimaryKey from Dic_Controls(strDeliveryDate_strDeliveryRef_1)


        ControlTABIndex = StartTABIndex
        If IsDate(strDeliveryDate) Then
            ControlDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
            strDeliveryDate = ControlDeliveryDate.ToString("dd/MM/yyyy")
            SaveDeliveryDate = ControlDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")
        Else
            'MsgBox ("Need to pass proper delivery date")
            'Exit Sub
            ControlDeliveryDate = CDate("01/01/1970")
            SaveDeliveryDate = ControlDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")
        End If
        ControlDeliveryRef = strDeliveryRef
        ControlDate = Now()

        tempControl = dic_Controls(strDeliveryDate & "_" & strDeliveryRef & "_" & "1")
        If tempControl IsNot Nothing Then
            UpdatedByName = tempControl.ControlUpdatedByName
            UpdatedByEmpno = tempControl.ControlUpdatedByEmpNo
            UpdatedByUsername = tempControl.ControlUpdatedByUsername
        End If

        ControlDBTable = "tblLabourHours"

        If Len(ControlDeliveryRef) > 0 Then
            SearchText = ControlDeliveryRef
            SearchField = "DeliveryReference"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = ""
        Else
            SearchText = ControlASN
            SearchField = "ASN_Number"
            FieldType = "STRING"
            ReturnField = "ID"
            ReturnValue = ""
            SearchCriteria = ""
        End If

        SortFields = ""

        Reversed = False
        ExcludeFields = ""
        RecID = ""
        LoadedOK = Find_myQuery(frmMainGIForm.myConnString, ControlDBTable, SearchField, SearchText, FieldType, ReturnField, RecID, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)

        'ControlTotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1
        If Not LoadedOK Then
            ControlTotalRows = 0
        Else
            ControlTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
        End If
        ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
        If ExtractTotals IsNot Nothing Then
            ExtractTotals.Total_Operatives = ControlTotalRows
        Else
            'ExtractTotals.Total_Operatives = 2
        End If
        ControlLeft = 7
        ControlTop = 22
        ControlWidth = 150
        ControlHeight = 21
        ControlTAG = ""
        ControlStartTAG = "40"
        ControlEndTAG = "442"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "TEXTBOX"
        ControlText = EmployeeNo
        ControlStyle = ""
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(216, 218, 214) 'Lt Grey
        ControlFieldname = ""
        ControlASN = strASNNo
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtEmployeeNo", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER")

        ControlTABIndex = ControlTABIndex + 1

        'Label: lblEmployeeNo

        ControlLeft = 7
        ControlTop = 1
        ControlWidth = 150
        ControlHeight = 20
        ControlTAG = "lblEmployeeNo"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "LABEL"
        ControlText = "Employee No"
        ControlStyle = ""
        ControlForeColor = RGB(255, 240, 60) 'YELLOW
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlFieldname = ""
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "lblEmployeeNo", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)


        'Label: lblFLMName

        ControlLeft = 162
        ControlTop = 1
        ControlWidth = 180
        ControlHeight = 18
        ControlTAG = "lblFLMName"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "LABEL"
        ControlText = "FLM Name"
        ControlStyle = ""
        ControlForeColor = RGB(255, 240, 60) 'YELLOW
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlFieldname = ""
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "lblFLMName", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)

        'Label: lblFLMStartTime

        ControlLeft = 406
        ControlTop = 1
        ControlWidth = 136
        ControlHeight = 18
        ControlTAG = "lblFLMStartTime"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "LABEL"
        ControlText = "Start Time"
        ControlStyle = ""
        ControlForeColor = RGB(255, 240, 60) 'YELLOW
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlFieldname = ""
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "lblFLMStartTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)

        'Label: lblFLMFinishTime

        ControlLeft = 547
        ControlTop = 1
        ControlWidth = 136
        ControlHeight = 18
        ControlTAG = "lblFLMFinishTime"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "LABEL"
        ControlText = "Finish Time"
        ControlStyle = ""
        ControlForeColor = RGB(255, 240, 60) 'YELLOW
        ControlBackColor = RGB(0, 112, 192) 'BLUE
        ControlFieldname = ""
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "lblFLMFinishTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)

        'comboBOX -FLM NAME
        ControlLeft = 162
        ControlTop = 22
        ControlWidth = 180
        ControlHeight = 20
        ControlTAG = "40"
        ControlStartTAG = "40"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "COMBOBOX"
        If Len(FLMName) = 0 Then
            ControlText = "Select FLM"
        Else
            ControlText = FLMName
        End If
        ControlStyle = ""
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(240, 248, 255) 'ALICEBLUE
        ControlFieldname = "FLM_Name"
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Firstname = ""
        Lastname = ""
        Fullname = ""
        strSQL = "SELECT Firstname,Lastname FROM tblEmployees ORDER BY Firstname"
        'Criteria = "Description = " & Chr(34) & "FLM" & Chr(34)
        Criteria = ""
        ListArray = Get_DropdownItems(frmMainGIForm.myConnString, strSQL, Criteria, ErrMessage)
        ReDim ComboArray(UBound(ListArray, 2) + 1)
        For RowIDX = 0 To UBound(ListArray, 2)
            If Not IsNothing(ListArray(0, RowIDX)) Then
                Firstname = ListArray(0, RowIDX)
            End If
            If Not IsNothing(ListArray(1, RowIDX)) Then
                Lastname = ListArray(1, RowIDX)
            End If
            Fullname = Firstname & " " & Lastname
            If Len(Fullname) > 1 Then
                'put into combo array
                ComboArray(ComboIDX) = Fullname
                ComboIDX = ComboIDX + 1
            End If
        Next

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "comFLMName", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor)

        ControlTABIndex = ControlTABIndex + 1
        'Start Time BUTTON

        ControlLeft = 406
        ControlTop = 21
        ControlWidth = 40
        ControlHeight = 25
        ControlTAG = "btn8"
        ControlStartTAG = "40"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "BTN"
        ControlText = "@S"
        ControlStyle = "BOLD,ITALIC"
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(255, 215, 0) 'GOLD
        ControlFieldname = "FLM_StartTime"
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "btnFLMStartTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER",
                           "", DisableStartTimeButton, False, dtStartTime, dtFinishTime)

        ControlTABIndex = ControlTABIndex + 1
        'StartTime TEXTBOX:

        ControlLeft = 461
        ControlTop = 22
        ControlWidth = 80
        ControlHeight = 21
        ControlTAG = "441"
        ControlStartTAG = "40"
        ControlEndTAG = "442"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "TEXTBOX"
        If dtStartTime > #1970-01-01# Then
            ControlText = dtStartTime.ToString("HH:mm:ss")
        Else
            ControlText = "00:00:00"
        End If

        ControlStyle = ""
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(216, 218, 214) 'Lt Grey
        ControlFieldname = "FLM_StartTime"
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtFLMStartTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER",
                           "", True, False, dtStartTime, dtFinishTime)

        ControlTABIndex = ControlTABIndex + 1
        'Finish Time BUTTON

        ControlLeft = 547
        ControlTop = 21
        ControlWidth = 40
        ControlHeight = 25
        ControlTAG = "btn9"
        ControlStartTAG = "40"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "BTN"
        ControlText = "@F"
        ControlStyle = "BOLD,ITALIC"
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(255, 215, 0) 'GOLD
        ControlFieldname = "FLM_FinishTime"
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "btnFLMFinishTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER",
                           "", DisableFinishTimeButton, False, dtStartTime, dtFinishTime)

        ControlTABIndex = ControlTABIndex + 1
        'FinishTime TEXTBOX:

        ControlLeft = 603
        ControlTop = 22
        ControlWidth = 80
        ControlHeight = 21
        ControlTAG = "442"
        ControlStartTAG = "40"
        ControlEndTAG = "442"
        ControlFontName = "Cambria"
        ControlFontSize = 10
        ControlType = "TEXTBOX"
        If dtFinishTime > #1970-01-01# Then
            ControlText = dtFinishTime.ToString("HH:mm:ss")
        Else
            ControlText = "00:00:00"
        End If
        ControlStyle = ""
        ControlForeColor = RGB(0, 0, 0) 'black
        ControlBackColor = RGB(216, 218, 214) 'Lt Grey
        ControlFieldname = "FLM_FinishTime"
        'ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
        ComboArray = Nothing
        ComboIDX = 0

        Call AddNewControl(True, ScrollFrame, ControlDBTable, ControlFieldname, "ID", Nothing,
            "txtFLMFinishTime", ControlText, ControlType, ControlTAG, ControlTABIndex,
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef,
                           frmMainGIForm.myEmpNo, frmMainGIForm.myUsername, frmMainGIForm.myName, ControlASN,
        ControlOBJCount, ControlStartTAG, ControlEndTAG, ControlRowNumber, ControlTotalRows, MakeVisible,
        ComboArray, ControlBackColor, ControlLeftMargin, ControlFontName, ControlFontSize, ControlStyle, ControlForeColor, "CENTER", "",
                           True, False, dtStartTime, dtFinishTime)

        ControlTABIndex = ControlTABIndex + 1
        'Loadedok should be true if FLM record already exists :
        LabourFieldsArr = strToStringArray(SCFieldnames, ",", 0, False, False, False, "_", False, 34, 39)


        SaveCriteria = ""
        'tblLabourHours:
        If LoadedOK Then
            'UPDATE record:
            'Need to grab the Total Rows etc.
            If IsNumeric(GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")) Then
                ControlTotalRows = GetMYValuebyFieldname(AllValues, AllFields, "Total_Rows")
            End If
            dtDeliveryDate = CDate(strDeliveryDate)
            'dtLastSaved = GetMYValuebyFieldname(AllValues, AllFields, "LastSaved")
            dtLastSaved = Now()
            'dtNow = Now()
            SaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")
            'ControlTotalRows = ' From database or from dic_Totals() ??????
            If Len(FLMName) > 0 Then
                LabourFields = "DeliveryDate,DeliveryReference,FLM_Name,FLM_StartTime,FLM_FinishTime,Total_Rows,UpdatedByEmpno,UpdatedByUsername,UpdatedByName,LastSaved"
                'LabourFieldValues = Chr(39) & SaveDeliveryDate & Chr(39) & "," & Chr(34) & strDeliveryRef & Chr(34)
                'LabourFieldValues = LabourFieldValues & "," & Chr(34) & FLMName & Chr(34) & "," & Chr(39) & dtStartTime.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39)
                'LabourFieldValues = LabourFieldValues & "," & Chr(39) & dtFinishTime.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39) & "," & Chr(34) & ControlTotalRows & Chr(34)
                'LabourFieldValues = LabourFieldValues & "," & Chr(39) & dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39)

                LabourFieldValues = dtDeliveryDate & "," & strDeliveryRef
                LabourFieldValues = LabourFieldValues & "," & FLMName & "," & dtStartTime
                LabourFieldValues = LabourFieldValues & "," & dtFinishTime & "," & ControlTotalRows
                LabourFieldValues = LabourFieldValues & "," & UpdatedByEmpno & "," & UpdatedByUsername & "," & UpdatedByName
                LabourFieldValues = LabourFieldValues & "," & dtLastSaved

            Else
                LabourFields = "DeliveryDate,DeliveryReference,Total_Rows,UpdatedByEmpno,UpdatedByUsername,UpdatedByName,LastSaved"
                'LabourFieldValues = Chr(39) & SaveDeliveryDate & Chr(39) & "," & Chr(34) & strDeliveryRef & Chr(34) & "," & Chr(34) & ControlTotalRows & Chr(34)
                'LabourFieldValues = LabourFieldValues & "," & Chr(39) & dtLastSaved.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39)

                LabourFieldValues = dtDeliveryDate & "," & strDeliveryRef & "," & ControlTotalRows
                LabourFieldValues = LabourFieldValues & "," & UpdatedByEmpno & "," & UpdatedByUsername & "," & UpdatedByName
                LabourFieldValues = LabourFieldValues & "," & dtLastSaved

            End If
            If Len(RecID) > 0 Then
                SaveCriteria = "ID = " & RecID
            End If
            If Len(strDeliveryRef) > 0 And SaveFLMDetails = True Then
                'LoadedOK should be true for update - check the returned ID:

                'SavedOK = InsertUpdateMyRecord(LoadedOK, frmMainGIForm.myConnString, ControlDBTable, LabourFields, LabourFieldValues, ErrMessage, SaveCriteria,
                'ExcludeFields)
                SavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, LoadedOK, "", ControlDBTable, LabourFields, LabourFieldValues,
                                                               SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            End If
        Else
            dtDeliveryDate = CDate(strDeliveryDate)
            dtNow = Now()
            SaveDeliveryDate = dtDeliveryDate.ToString("yyyy-MM-dd HH:mm:ss")
            LabourFields = "DeliveryDate,DeliveryReference,Total_Rows,UpdatedByEmpno,UpdatedByUsername,UpdatedByName,LastSaved"
            'LabourFieldValues = Chr(39) & SaveDeliveryDate & Chr(39) & "," & Chr(34) & strDeliveryRef & Chr(34) & "," & Chr(34) & ControlTotalRows & Chr(34)
            'LabourFieldValues = LabourFieldValues & "," & Chr(39) & dtNow.ToString("yyyy-MM-dd HH:mm:ss") & Chr(39)
            LabourFieldValues = dtDeliveryDate & "," & strDeliveryRef & "," & ControlTotalRows
            LabourFieldValues = LabourFieldValues & "," & UpdatedByEmpno & "," & UpdatedByUsername & "," & UpdatedByName
            LabourFieldValues = LabourFieldValues & "," & dtNow
            ExcludeFields = ""
            SaveCriteria = ""
            If Len(strDeliveryRef) > 0 Then
                'Only save if record not exists already - loadedok should be true for update - check the returned ID:

                'SavedOK = InsertUpdateMyRecord(LoadedOK, frmMainGIForm.myConnString, ControlDBTable, LabourFields, LabourFieldValues, ErrMessage, SaveCriteria,
                'ExcludeFields)
                SavedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, LoadedOK, "", ControlDBTable, LabourFields, LabourFieldValues,
                                                               SaveCriteria, ExcludeFields, ErrMessage, False, ",")
            End If
        End If
        StartTABIndex = ControlTABIndex

    End Sub

    Sub InsertOperativeRow(strDeliveryDate As String, strDeliveryRef As String, Optional strASNNo As String = "")
        Dim OpID As Long
        Dim LowerTAG As Long
        Dim UpperTAG As Long
        Dim TimeStartTAG As Long
        Dim FieldNames As String
        Dim ErrMessage As String = ""
        Dim TotalRows As Long
        Dim StartFieldIndex As Long
        Dim StartBtnTAG As Long = 0
        Dim HighestOpTag As Long = 0
        Dim ControlsPerRow As Long = 0
        Dim TotalControlsInFrame As Long = 0
        Dim frm As Form
        Dim OperativeFrame As ScrollableControl = Nothing
        Dim FindControl As Control = Nothing

        'Need to calculate the UPPER TAG from TOTAL OPERATIVES and Total Number of controls in the FRAME_OPERATIVES:
        'DO NOT USE THIS !

        LowerTAG = 43
        UpperTAG = 0
        TotalRows = 0
        StartBtnTAG = 0
        HighestOpTag = 0
        TimeStartTAG = 400
        MsgBox("In InsertOperativeRow")
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
        ControlsPerRow = 6
        TotalControlsInFrame = (TotalRows * ControlsPerRow)

        For Each frm In Application.OpenForms
            If UCase(frm.Name) = frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx Then
                FindControl = frmMainGIForm.FindFrameControls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "Frame_Operatives")
            End If
        Next
        If FindControl IsNot Nothing Then
            OperativeFrame = CType(FindControl, System.Windows.Forms.ScrollableControl)
        End If
        If Len(ErrMessage) > 0 Then
            MsgBox("Error while Getting FieldNames")
            Exit Sub
        End If

        MsgBox("INSERT OP ROW")
        StartFieldIndex = 5
        If TotalRows = 0 Then
            UpperTAG = ((1 * ControlsPerRow) - 1) + LowerTAG
        Else
            UpperTAG = ((TotalRows * ControlsPerRow) - 1) + LowerTAG
        End If

        OpID = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef)
        AddNewOperatives(frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef), OpID, OperativeFrame, HighestOpTag, StartBtnTAG,
                         strDeliveryDate, strDeliveryRef, strASNNo, LowerTAG, UpperTAG, TimeStartTAG, OpFieldnames, TotalRows, StartFieldIndex)
        'Me.txtTotalOps.Text = CStr(frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1)
        frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtTotalOps", CStr(frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1))
        frmMainGIForm.Dic_HighestOpTAGID(strDeliveryDate & "_" & strDeliveryRef) = HighestOpTag 'Should be 48 after first row entry (48 = comments TAG)
        frmMainGIForm.Dic_HighestOpBtnTAGID(strDeliveryDate & "_" & strDeliveryRef) = StartBtnTAG


    End Sub

    Sub SearchAndInsertName(Entry As String, strDeliveryDate As String, strDeliveryRef As String, NumberOfCharsEntered As Integer)
        '1) wait until 7 chars have been entered.
        '2) Grab Entry and check against tblEmployees to get name.
        '3) Insert new Operative Row - noting the NEW ROW number
        '4) Insert the name into the Operative Name combo box.
        Dim FirstName As String = ""
        Dim Lastname As String = ""
        Dim SearchValue As String
        Dim SearchField As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim Criteria As String
        Dim ExcludeFields As String
        Dim FoundName As Boolean = False
        Dim AllReturnedValues() As Object
        Dim AllFields() As String
        Dim DBEmployees As String
        Dim TotalRows As Long
        Dim OpNameField As String = ""
        Dim Fullname As String
        Dim OpStartTime As String
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim tempControl As clsControls
        Dim Frame_Operatives As ScrollableControl
        Dim TimeOut As DateTime
        Dim StartTimeControl As Control
        Dim SearchControlName As String
        Dim txtControl As TextBox
        Dim strTagNumber As String
        Dim ChangeProperty As String
        Dim testCollection As Object = Nothing
        Dim NewRow As Long = 0
        Dim EmptyRow As Long = 0
        Dim strEmptyRow As String = ""
        Dim strOpName As String = ""
        Dim strOpActivity As String = ""
        Dim FoundBlankRow As Boolean = False
        Dim BlankRowID As String = ""
        Dim VarKey As String = ""

        DBEmployees = "tblEmployees"
        'Criteria = "Description = " & Chr(34) & "FLM" & Chr(34)
        Criteria = ""
        AllReturnedValues = Nothing
        AllFields = Nothing

        Frame_Operatives = GetFrameControl(CPFormName & FormID, "Frame_Operatives")

        If NumberOfCharsEntered < 7 Then
            SearchField = "EmpNo"
        Else
            SearchField = "EmpNo_7digits"
        End If
        ASNNo = ""
        NewRow = 0
        If Len(Entry) > 4 Then
            SearchValue = Entry
            FoundName = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, DBEmployees, SearchField, SearchValue, "STRING",
                                                             ReturnField, ReturnValue, AllReturnedValues, AllFields, Criteria)
            If FoundName Then
                FirstName = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFields, "Firstname")
                Lastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFields, "Lastname")
                Fullname = FirstName & " " & Lastname
                'MsgBox("Firstname = " & FirstName & " , Lastname=" & Lastname)
                'INSERT NEW OPERATIVE ROw in FRAME_OPERATIVES and insert the name in the OpNAME field:
                'Call frmGI_RP_Userform.InsertOperative()

                'WHAT about searching dic_Controls() ???
                'Instead of TAG - use Control Name ????
                VarKey = strDeliveryDate & "_" & strDeliveryRef & "_" & "44"
                tempControl = dic_Controls(VarKey)
                If Not IsNothing(tempControl) Then
                    strOpName = tempControl.ControlValue
                End If
                VarKey = strDeliveryDate & "_" & strDeliveryRef & "_" & "45"
                tempControl = dic_Controls(VarKey)
                If Not IsNothing(tempControl) Then
                    strOpActivity = tempControl.ControlValue
                End If
                If Len(strOpName) = 0 And Len(strOpActivity) = 0 Or UCase(strOpName) = "SELECT EMPLOYEE" Then
                    strEmptyRow = CStr(Get_NumericPartOfString(tempControl.ControlName))
                    If IsNumeric(strEmptyRow) Then
                        NewRow = CLng(strEmptyRow)
                    End If
                End If
                ' ********* ADD OPERATIVE: *****************

                If IsDate(strDeliveryDate) Then
                    strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
                    ASNNo = ""
                    ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                    If ExtractTotals IsNot Nothing Then
                        TotalRows = ExtractTotals.Total_Operatives
                    End If
                    'TotalRows passed in - does NOT return an updated value !!!
                    If NewRow = 0 Then
                        NewRow = TotalRows + 1
                        TotalRows = InsertOperatives(True, strDeliveryDate, strDeliveryRef, ASNNo, Frame_Operatives, NewRow)
                    Else
                        TotalRows = InsertOperatives(False, strDeliveryDate, strDeliveryRef, ASNNo, Frame_Operatives, NewRow)
                    End If
                    'How do we fill in the existing row ?

                Else
                    MsgBox("Something wrong with the Delivery Date ?")
                End If

                'AfterUpdate procedure is called now because of the trigger in InsertOperative of the text changing.
                'TotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1
                SearchControlName = "txtOperativeStartTime:" & CStr(TotalRows)
                StartTimeControl = frmMainGIForm.FindFrameControls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, SearchControlName)
                If StartTimeControl IsNot Nothing Then
                        txtControl = CType(StartTimeControl, System.Windows.Forms.TextBox)
                        TimeOut = Now()
                        strTagNumber = txtControl.Tag
                        'find the contrl.  Get the tAG number. need to put the full date and time into the controls object clsControls .
                        txtControl.Text = Format(TimeOut, "HH:mm:ss")
                        OpStartTime = Format(TimeOut, "HH:mm:ss")
                        OpNameField = "comOperativeName:" & CStr(TotalRows)
                        Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, OpNameField, Fullname)
                        Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, SearchControlName, OpStartTime) 'START TIME FIELD
                        Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtTotalOps", TotalRows)
                        'Has the new name now been recorded in clsControls ? - does it get called again ?
                        ChangeProperty = "TheControl"
                        testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, strDeliveryDate, ASNNo,
                                                      strDeliveryRef, strTagNumber, ChangeProperty)
                        If testCollection IsNot Nothing Then
                            dic_Controls = testCollection
                        End If
                        ChangeProperty = "OpStartTime"
                        testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, strDeliveryDate, ASNNo,
                                                      strDeliveryRef, strTagNumber, ChangeProperty)
                        If testCollection IsNot Nothing Then
                            dic_Controls = testCollection
                        End If
                    Else
                        MsgBox("Cannot find START TIME CONTROL")
                    End If
                End If
            End If
    End Sub

    Sub SearchAndInsertFirstname(Entry As String, strDeliveryDate As String, strDeliveryRef As String, NumberOfCharsEntered As Integer)
        'As user types into the Firstname combo - it should have an-ever decreasing list of items,
        'as it keeps searching and matching the chars the user keep typing:
        Dim FirstName As String = ""
        Dim Lastname As String = ""
        Dim SearchValue As String
        Dim SearchField As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim Criteria As String
        Dim ExcludeFields As String
        Dim FoundName As Boolean = False
        Dim AllReturnedValues() As Object
        Dim AllFields() As String
        Dim DBEmployees As String
        Dim TotalRows As Long
        Dim OpNameField As String = ""
        Dim Fullname As String
        Dim OpStartTime As String
        Dim ASNNo As String
        Dim ExtractTotals As clsTotals
        Dim tempControl As clsControls
        Dim Frame_Operatives As ScrollableControl
        Dim TimeOut As DateTime
        Dim StartTimeControl As Control
        Dim SearchControlName As String
        Dim txtControl As TextBox
        Dim strTagNumber As String
        Dim ChangeProperty As String
        Dim testCollection As Object = Nothing
        Dim NewRow As Long = 0
        Dim EmptyRow As Long = 0
        Dim strEmptyRow As String = ""
        Dim strOpName As String = ""
        Dim strOpActivity As String = ""
        Dim FoundBlankRow As Boolean = False
        Dim BlankRowID As String = ""
        Dim VarKey As String = ""

        DBEmployees = "tblEmployees"
        'Criteria = "Description = " & Chr(34) & "FLM" & Chr(34)
        Criteria = ""
        AllReturnedValues = Nothing
        AllFields = Nothing

        Frame_Operatives = GetFrameControl(CPFormName & FormID, "Frame_Operatives")

        SearchField = "Firstname"
        ASNNo = ""
        NewRow = 0
        If Len(Entry) > 2 Then
            SearchValue = Entry
            FoundName = Module_DanG_MySQL_Tools.Find_myQuery(frmMainGIForm.myConnString, DBEmployees, SearchField, SearchValue, "STRING",
                                                             ReturnField, ReturnValue, AllReturnedValues, AllFields, Criteria)
            If FoundName Then
                FirstName = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFields, "Firstname")
                Lastname = Module_DanG_MySQL_Tools.GetMYValuebyFieldname(AllReturnedValues, AllFields, "Lastname")
                Fullname = FirstName & " " & Lastname
                'MsgBox("Firstname = " & FirstName & " , Lastname=" & Lastname)
                'INSERT NEW OPERATIVE ROw in FRAME_OPERATIVES and insert the name in the OpNAME field:
                'Call frmGI_RP_Userform.InsertOperative()

                'WHAT about searching dic_Controls() ???
                'Instead of TAG - use Control Name ????
                VarKey = strDeliveryDate & "_" & strDeliveryRef & "_" & "44"
                tempControl = dic_Controls(VarKey)
                strOpName = tempControl.ControlValue
                VarKey = strDeliveryDate & "_" & strDeliveryRef & "_" & "45"
                tempControl = dic_Controls(VarKey)
                strOpActivity = tempControl.ControlValue
                If Len(strOpName) = 0 And Len(strOpActivity) = 0 Or UCase(strOpName) = "SELECT EMPLOYEE" Then
                    strEmptyRow = CStr(Get_NumericPartOfString(tempControl.ControlName))
                    If IsNumeric(strEmptyRow) Then
                        NewRow = CLng(strEmptyRow)
                    End If
                End If
                ' ********* ADD OPERATIVE: *****************

                If IsDate(strDeliveryDate) Then
                    strDeliveryDate = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
                    ASNNo = ""
                    ExtractTotals = dic_Totals(strDeliveryDate & "_" & strDeliveryRef)
                    If ExtractTotals IsNot Nothing Then
                        TotalRows = ExtractTotals.Total_Operatives
                    End If
                    'TotalRows passed in - does NOT return an updated value !!!
                    If NewRow = 0 Then
                        NewRow = TotalRows + 1
                        TotalRows = InsertOperatives(True, strDeliveryDate, strDeliveryRef, ASNNo, Frame_Operatives, NewRow)
                    Else
                        TotalRows = InsertOperatives(False, strDeliveryDate, strDeliveryRef, ASNNo, Frame_Operatives, NewRow)
                    End If
                    'How do we fill in the existing row ?

                Else
                    MsgBox("Something wrong with the Delivery Date ?")
                End If

                'AfterUpdate procedure is called now because of the trigger in InsertOperative of the text changing.
                'TotalRows = frmMainGIForm.Dic_TotalOperatives(strDeliveryDate & "_" & strDeliveryRef) - 1
                SearchControlName = "txtOperativeStartTime:" & CStr(TotalRows)
                StartTimeControl = frmMainGIForm.FindFrameControls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, SearchControlName)
                If StartTimeControl IsNot Nothing Then
                    txtControl = CType(StartTimeControl, System.Windows.Forms.TextBox)
                    TimeOut = Now()
                    strTagNumber = txtControl.Tag
                    'find the contrl.  Get the tAG number. need to put the full date and time into the controls object clsControls .
                    txtControl.Text = Format(TimeOut, "HH:mm:ss")
                    OpStartTime = Format(TimeOut, "HH:mm:ss")
                    OpNameField = "comOperativeName:" & CStr(TotalRows)
                    Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, OpNameField, Fullname)
                    Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, SearchControlName, OpStartTime) 'START TIME FIELD
                    Call frmMainGIForm.InsertValueIntoForm(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "txtTotalOps", TotalRows)
                    'Has the new name now been recorded in clsControls ? - does it get called again ?
                    ChangeProperty = "TheControl"
                    testCollection = UpdateCollection(dic_Controls, Nothing, txtControl, strDeliveryDate, ASNNo,
                                                      strDeliveryRef, strTagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                    ChangeProperty = "OpStartTime"
                    testCollection = UpdateCollection(dic_Controls, Nothing, TimeOut, strDeliveryDate, ASNNo,
                                                      strDeliveryRef, strTagNumber, ChangeProperty)
                    If testCollection IsNot Nothing Then
                        dic_Controls = testCollection
                    End If
                Else
                    MsgBox("Cannot find START TIME CONTROL")
                End If
            End If
        End If
    End Sub

    Friend Sub Newbutton_click(sender As Object, e As System.EventArgs)
        Dim btn As Button = DirectCast(sender, Button)

        'INSERT the TIME into the relevant control here:
        MsgBox("NEWbutton Control Name: " & btn.Name & ", Caption=" & btn.Text & ", tag=" & btn.Tag) 'YEZ THIS WORKS !!!!!
    End Sub

    Friend Sub NewCombo_textchanged(sender As Object, e As System.EventArgs)
        Dim combo As ComboBox = DirectCast(sender, ComboBox)

        MsgBox("Combo Box: " & combo.Name & " Changed value= " & combo.Text) 'YEA THIS WORKS !!!
    End Sub

    Friend Sub NewTextBox_textchanged(sender As Object, e As System.EventArgs)
        Dim txtbox As TextBox = DirectCast(sender, TextBox)
        'NEver seems to execute ??? NewObj_Textchanged takes over instead ????
        'MsgBox("TEXTBOX: " & txtbox.Name & " Changed value= " & txtbox.Text) ' *************************** IS CALLED AND WORKS.
    End Sub

    Friend Sub NewObj_textchanged(sender As Object, e As EventArgs)
        If TypeOf sender Is ComboBox Then
            Dim combo As ComboBox = DirectCast(sender, ComboBox)

            MsgBox("NEW OBJ COMBO Control Name: " & combo.Name & ", Value=" & combo.Text)
        End If

        If TypeOf sender Is TextBox Then
            Dim txtbox As TextBox = DirectCast(sender, TextBox)

            MsgBox("NEW OBJ TEXTBOX Control Name: " & txtbox.Name & ", Value=" & txtbox.Text)
        End If
        If TypeOf sender Is Control Then
            Dim ctrl As Control = DirectCast(sender, Control)

            MsgBox("NEW OBJ CONTROL Name: " & ctrl.Name & ", Value=" & ctrl.Text)
        End If
    End Sub

    Public Function FindControl(ByVal parent As Control, ByVal ControlName As String, Optional ByVal ControlTAG As String = "") As Control
        Dim ControlIDX As Integer = 0
        Dim tmpctrl As Control
        Dim tmpctrl2 As Control
        For Each tmpctrl In parent.Controls
            ControlIDX = ControlIDX + 1
            If Len(ControlName) > 0 Then
                If UCase(tmpctrl.Name) = UCase(ControlName) Then
                    Return tmpctrl
                ElseIf tmpctrl.HasChildren Then
                    tmpctrl2 = FindControl(tmpctrl, ControlName, "")
                    If Not tmpctrl2 Is Nothing Then
                        Return tmpctrl2
                    End If
                End If
            End If
            If Len(ControlTAG) > 0 Then
                If UCase(tmpctrl.Tag) = UCase(ControlTAG) Then
                    Return tmpctrl
                ElseIf tmpctrl.HasChildren Then
                    tmpctrl2 = FindControl(tmpctrl, "", ControlTAG)
                    If Not tmpctrl2 Is Nothing Then
                        Return tmpctrl2
                    End If
                End If
            End If
        Next
        ' Not found
        Return Nothing
    End Function

    Function UpdateCollection(ByRef coll As Object, varKey As Object, ValueToChange As Object, DeliveryDate As DateTime,
    Optional ASNNumber As String = "", Optional DeliveryRef As String = "", Optional TAGNUMBER As String = "",
    Optional PropertyToChange As String = "VALUE") As Object
        Dim NewKey As String = ""
        Dim tempControl As New clsControls
        Dim myDateTime As DateTime
        Dim myDic As New Scripting.Dictionary
        'ADD extra optional parameter here to select which PROPERTY to change:
        ' - default  will be ControlValue but we also need:
        'ControlFLMName
        'ControlFLMStartDateTime
        'ControlFLMEndDateTime
        'ControlOpName
        'ControlOpActivity
        'ControlOpStartDateTime
        'ControlOpEndDateTime

        UpdateCollection = Nothing

        If Len(varKey) > 0 Then
            NewKey = varKey
        Else
            If Len(CStr(DeliveryDate)) > 0 Then
                If Len(DeliveryRef) > 0 Then
                    NewKey = DeliveryDate & "_" & DeliveryRef & "_" & TAGNUMBER
                ElseIf Len(ASNNumber) > 0 Then
                    NewKey = DeliveryDate & "_" & ASNNumber & "_" & TAGNUMBER
                Else
                    NewKey = DeliveryDate & "_" & "" & "_" & TAGNUMBER
                End If
            End If
        End If
        If Right(NewKey, 1) = "0" Then
            If Not IsNumeric(Right(NewKey, 2)) Then
                'Set coll = Nothing
                'Set UpdateCollection = Nothing
                Exit Function
            End If
        End If
        'Need to refer to Dic_Collection here - and use .EXISTS()
        'If Not InCollection("MISSING", coll, NewKey) Then
        'MsgBox ("Not in collection: " & NewKey)
        'Exit Function
        'End If
        tempControl = Nothing
        myDic = coll
        If Not coll.exists(NewKey) Then
            MsgBox("KEY NOT AVAILABLE: " & NewKey)
            Exit Function
        End If

        If Not coll Is Nothing Then
            If Len(NewKey) > 0 And ValueToChange IsNot Nothing Then
                tempControl = coll.Item(NewKey)
                'tempControl = coll(NewKey)
                'tempControl.ControlObject = Nothing
                If UCase(PropertyToChange) = "VALUE" Then
                    If IsNumeric(ValueToChange) Then
                        tempControl.ControlValue = ValueToChange.ToString
                    ElseIf IsDate(ValueToChange) Then
                        myDateTime = ValueToChange
                        tempControl.ControlValue = ValueToChange.ToString
                    Else
                        tempControl.ControlValue = ValueToChange.ToString
                    End If
                ElseIf UCase(PropertyToChange) = "FIELDNAME" Then
                    tempControl.ControlFieldName = ValueToChange
                ElseIf UCase(PropertyToChange) = "CONTROLNAME" Then
                    tempControl.ControlName = ValueToChange
                ElseIf UCase(PropertyToChange) = "FRAMENAME" Then
                    tempControl.ControlFrameName = ValueToChange
                ElseIf UCase(PropertyToChange) = "FLM_NAME" Then
                    tempControl.ControlFLMName = ValueToChange
                ElseIf UCase(PropertyToChange) = "LABOUR_COMMENTS" Then
                    tempControl.ControlLabour_Comments = ValueToChange
                ElseIf UCase(PropertyToChange) = "ASN" Then
                    tempControl.ControlASNID = ValueToChange
                ElseIf UCase(PropertyToChange) = "DELIVERYDATE" Then
                    tempControl.ControlDeliveryDate = ValueToChange
                ElseIf UCase(PropertyToChange) = "DELIVERYREF" Then
                    tempControl.ControlDeliveryRef = ValueToChange
                ElseIf UCase(PropertyToChange) = "STARTTAG" Then
                    tempControl.ControlStartTAG = ValueToChange
                ElseIf UCase(PropertyToChange) = "ENDTAG" Then
                    tempControl.ControlEndTAG = ValueToChange
                ElseIf UCase(PropertyToChange) = "FLMSTARTTIME" Then
                    tempControl.ControlFLMStartDateTime = ValueToChange
                ElseIf UCase(PropertyToChange) = "FLMFINISHTIME" Then
                    tempControl.ControlFLMEndDateTime = ValueToChange
                ElseIf UCase(PropertyToChange) = "OPROWNUMBER" Then
                    tempControl.ControlOpRowNumber = ValueToChange
                ElseIf UCase(PropertyToChange) = "OPNAME" Then
                    tempControl.ControlOpName = ValueToChange
                ElseIf UCase(PropertyToChange) = "ACTIVITYNAME" Then
                    tempControl.ControlOpActivity = ValueToChange
                ElseIf UCase(PropertyToChange) = "OPSTARTTIME" Then
                    tempControl.ControlOpStartDateTime = ValueToChange
                ElseIf UCase(PropertyToChange) = "OPFINISHTIME" Then
                    tempControl.ControlOpEndDateTime = ValueToChange
                ElseIf UCase(PropertyToChange) = "OPCOMMENTS" Then
                    tempControl.Control_Opcomments = ValueToChange
                ElseIf UCase(PropertyToChange) = "TOTALHOURS" Then
                    tempControl.ControlOPdblTotalHrs = ValueToChange
                ElseIf UCase(PropertyToChange) = "EXTRASHORT" Then
                    tempControl.ControlExtraOrShort = ValueToChange
                ElseIf UCase(PropertyToChange) = "PARTNO" Then
                    tempControl.ControlPartNo = ValueToChange
                ElseIf UCase(PropertyToChange) = "PARTQTY" Then
                    tempControl.ControlQty = ValueToChange
                ElseIf UCase(PropertyToChange) = "SHORTROWNUMBER" Then
                    tempControl.ControlShortRowNumber = ValueToChange
                ElseIf UCase(PropertyToChange) = "EXTRAROWNUMBER" Then
                    tempControl.ControlExtraRowNumber = ValueToChange
                ElseIf UCase(PropertyToChange) = "LASTSAVED" Then
                    tempControl.ControlLastSaved = ValueToChange
                ElseIf UCase(PropertyToChange) = "CONTROLLIST" Then
                    tempControl.ControlList = ValueToChange
                ElseIf UCase(PropertyToChange) = "BACKCOLOR" Then
                    tempControl.ControlBackColor = ValueToChange
                ElseIf UCase(PropertyToChange) = "FORECOLOR" Then
                    tempControl.ControlForeColor = ValueToChange
                ElseIf UCase(PropertyToChange) = "FONTNAME" Then
                    tempControl.ControlFontName = ValueToChange
                ElseIf UCase(PropertyToChange) = "FONTSIZE" Then
                    tempControl.ControlFontSize = ValueToChange
                ElseIf UCase(PropertyToChange) = "PRIMARYKEY" Then
                    tempControl.ControlPrimaryKey = ValueToChange
                ElseIf UCase(PropertyToChange) = "DBTABLE" Then
                    tempControl.ControlDBTable = ValueToChange
                ElseIf UCase(PropertyToChange) = "TAG" Then
                    tempControl.ControlTAG = ValueToChange
                ElseIf UCase(PropertyToChange) = "TOTALROWS" Then
                    tempControl.ControlTotalRows = ValueToChange
                ElseIf UCase(PropertyToChange) = "ROWNUMBER" Then
                    tempControl.ControlFrameRowNumber = ValueToChange
                ElseIf UCase(PropertyToChange) = "ERRMESSAGE" Then
                    tempControl.ControlErrMessage = ValueToChange
                ElseIf UCase(PropertyToChange) = "TOP" Then
                    tempControl.ControlTop = ValueToChange
                ElseIf UCase(PropertyToChange) = "LEFT" Then
                    tempControl.ControlLeft = ValueToChange
                ElseIf UCase(PropertyToChange) = "WIDTH" Then
                    tempControl.ControlWidth = ValueToChange
                ElseIf UCase(PropertyToChange) = "HEIGHT" Then
                    tempControl.ControlHeight = ValueToChange
                ElseIf UCase(PropertyToChange) = "LEFTMARGIN" Then 'TRUE OR FALSE TO SET A LEFT MARGIN OR NOT
                    tempControl.ControlLeftMargin = ValueToChange
                ElseIf UCase(PropertyToChange) = "THECONTROL" Then
                    If tempControl IsNot Nothing Then
                        tempControl.TheControl = ValueToChange
                    End If
                Else
                    tempControl.ControlValue = ValueToChange
                End If
                'coll.Remove(NewKey)
                'coll.Add(tempControl, NewKey)
                coll(NewKey) = tempControl
            Else
                MsgBox("Property Key is empty - cannot update Value: " & ValueToChange.ToString)
            End If
        End If

        UpdateCollection = coll
    End Function

    Sub Clear_Entry_Controls(Formname As String, LowerTAG As Long, UpperTAG As Long)
        Dim Clearctrl As Control
        Dim TagNumber As Long

        For TagNumber = LowerTAG To UpperTAG

            Clearctrl = frmMainGIForm.FindControls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, "", TagNumber)
            If TypeOf Clearctrl Is ComboBox Or TypeOf Clearctrl Is TextBox Then
                Clearctrl.Text = ""
            End If
        Next


    End Sub

End Module
