Imports System.Linq
Public Class frmReferenceProgress
    Public StrDeliveryRef As String
    Public ASNNo As String
    Public Criteria As String
    Public TodayOnly As Boolean = False
    Private Sub frmReferenceProgress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SETUP VIEW GRID:
        Dim SqlStr As String
        Dim NumRows As Integer
        Dim ErrMessage As String
        Dim ProgressFields As String

        Me.Timer_UpdateProgressGrid.Interval = 5000
        Me.Timer_UpdateProgressGrid.Enabled = True
        'rbALL.Parent = gbProgressFilter
        'rbRolled.Parent = gbProgressFilter
        'rbInProgress.Parent = gbProgressFilter
        'rbCompleted.Parent = gbProgressFilter

        ProgressFields = "STATUS,DeliveryReference AS 'REF',ASN_Number AS 'ASN',DeliveryDate AS 'Delivery',Last_Saved AS 'Last Saved',UpdatedByName AS 'Updated By'"
        ProgressFields = ProgressFields & ",Grid_No AS 'Grid No',Total_Operatives AS 'Total Ops',TotalShortRows AS 'Short',TotalExtraRows AS 'Extra',Supplier_Code AS 'Supplier Code',Supplier_Name AS 'Supplier'"
        ProgressFields = ProgressFields & ",SHIFT,Due_Time AS 'Due Time',Calc_Hours AS 'Calc Hrs',Total_Hours AS 'Total Hrs',Pallets_Due AS 'Pallets Due',Cartons_Due AS 'Cartons Due',Origin"
        ProgressFields = ProgressFields & ",Expected_Lines AS 'Expected Lines',Expected_Cases AS 'Expected Cases',Actual_Cases AS 'Actual Cases'"
        ProgressFields = ProgressFields & ",Estimated_Totes AS 'Estimated Totes',Estimated_Cages AS 'Estimated Cages',Estimated_Pallets AS 'Estimated Pallets',ArrivalTime AS 'Arrival Time',GRN_Number AS 'GRN'"


        SqlStr = "SELECT " & ProgressFields & " FROM tblDeliveryInfo order by Last_Saved desc,SHIFT"
        Call PopulateMyDataSource(dgvDeliveryProgress.DataSource, frmMainGIForm.myConnString, SqlStr, NumRows, ErrMessage)
        Criteria = ""
        dgvDeliveryProgress.Columns(5).Width = 200
    End Sub

    Sub Check_If_Reference_Exists(RowIndex As Integer, ColumnIndex As Integer, Optional ByRef FormTAG As String = "")
        'OK user clicks on a reference.
        'Get that reference.
        'Check if a TIMESHEET child form is already open ?
        'Check WHAT reference it is displaying ???
        'If it does NOT match the current reference just clicked - then open up a NEW TIMESHEET child form and display the NEW reference.
        Dim value As Object
        Dim RefValue As Object
        Dim cf As New frmGI_RP_Userform
        Dim i As Integer
        Dim ColName As String
        Dim RefCol As Integer
        Dim ASNCol As Integer
        Dim ChildForm As String
        Dim MDIChildIndex As Integer
        Dim FormUnique As String
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim SortFields As String
        Dim Reversed As Boolean = False
        Dim ErrMessage As String
        Dim SearchCriteria As String
        Dim SearchField As String
        Dim FieldType As String
        Dim SearchText As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim FoundASN As Boolean = False
        Dim FoundREF As Boolean = False
        Dim DBInfoTable As String = "tblDeliveryInfo"
        Dim strDeliveryDate As String
        Dim strDeliveryRef As String
        Dim strSupplier As String
        Dim strGrid As String
        Dim strLastSaved As String
        Dim strID As String
        Dim strCurrentStatus As String
        Dim controlBACKCOLOR As System.Drawing.Color
        Dim controlFORECOLOR As System.Drawing.Color



        'Need to test if the form is already open - with THAT reference ???
        If RowIndex < 0 Or ColumnIndex < 0 Then
            Exit Sub
        Else
            'Find out which columnIndex holds the reference number in the GRID ?
            RefCol = -1
            ASNCol = -1
            StrDeliveryRef = ""
            ASNNo = ""
            For i = 1 To dgvDeliveryProgress.Columns.Count - 1
                ColName = dgvDeliveryProgress.Columns(i).HeaderText
                If InStr(UCase(ColName), "REF") > 0 Then
                    RefCol = i
                End If
                If InStr(UCase(ColName), "ASN") > 0 Then
                    ASNCol = i
                End If
            Next
            value = dgvDeliveryProgress.Rows(RowIndex).Cells(ColumnIndex).Value
            RefValue = dgvDeliveryProgress.Rows(RowIndex).Cells(RefCol).Value
            StrDeliveryRef = CType(RefValue, String)
            If ColumnIndex = RefCol Then
                StrDeliveryRef = CType(value, String)
                frmMainGIForm.SelectedDeliveryRef = StrDeliveryRef

            End If
            If ColumnIndex = ASNCol Then
                ASNNo = CType(value, String)
                frmMainGIForm.SelectedASN = ASNNo
            End If
            'MsgBox("value= " & value)

            'Just open up a NEW CHILD FORM for the ControlPanel:
            'txtSelectDeliveryDate.Text
            If RefCol = ColumnIndex Or ASNCol = ColumnIndex Then
                'If Form cannot be found then create new child form instance:
                If frmMainGIForm.FindMyChild(StrDeliveryRef, ASNNo, False, FormUnique, MDIChildIndex, ChildForm) = False Then
                    FormID = frmMainGIForm.Get_GUID()
                    frmMainGIForm.ControlPanelIdx = FormID
                    frmMainGIForm.SelectedDeliveryRef = StrDeliveryRef
                    cf.MdiParent = frmMainGIForm
                    cf.StartPosition = FormStartPosition.Manual
                    AllValues = Nothing
                    AllFields = Nothing
                    SortFields = "DeliveryDate"
                    Reversed = True
                    ErrMessage = ""
                    SearchCriteria = ""
                    SearchField = "DELIVERYREFERENCE"
                    SearchText = StrDeliveryRef
                    FieldType = "String"
                    ReturnField = "ID"
                    ReturnValue = ""
                    If Len(StrDeliveryRef) > 0 Then
                        FoundASN = Find_myQuery(frmMainGIForm.myConnString, DBInfoTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                        ASNNo = GetMYValuebyFieldname(AllValues, AllFields, "ASN_Number")
                        cf.Text = "Goods In Control Panel_" & "REF:" & StrDeliveryRef
                    ElseIf Len(ASNNo) > 0 Then
                        'Find THE DELIVERY REFERENCE !!!
                        SearchField = "ASN_Number"
                        SearchText = ASNNo
                        FoundASN = Find_myQuery(frmMainGIForm.myConnString, DBInfoTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                       AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                        If FoundASN Then
                            StrDeliveryRef = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryReference")
                            cf.Text = "Goods In Control Panel_" & "REF:" & StrDeliveryRef
                        End If

                    Else
                        cf.Text = "Goods In Control Panel_" & Now().ToString("dd/MM/yyyy HH:mm:ss")
                    End If

                    cf.Name = CPFormName & FormID
                    cf.Tag = FormID
                    cf.txtDeliveryRef.Text = StrDeliveryRef
                    cf.comDeliveryRef.Text = StrDeliveryRef
                    cf.comASNNo.Text = ASNNo
                    cf.txtASNNum.Text = ASNNo
                    cf.Show()
                    'cf.ShowDialog() - gives error
                    'Clear_Entry_Controls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, 1, 40)
                    'Clear_Entry_Controls(frmMainGIForm.ControlPanelFormName & frmMainGIForm.ControlPanelIdx, 801, 807)
                    'frmMainGIForm.InsertValueIntoForm("GI_TIMESHEET_" & frmMainGIForm.ControlPanelIdx, "comDeliveryRef", StrDeliveryRef)
                    Application.DoEvents()
                Else
                    Application.OpenForms.Item(ChildForm).Activate()
                End If
            Else
                AllValues = Nothing
                AllFields = Nothing
                SortFields = "DeliveryDate"
                Reversed = True
                ErrMessage = ""
                SearchCriteria = ""
                SearchField = "DELIVERYREFERENCE"
                SearchText = StrDeliveryRef
                FieldType = "String"
                ReturnField = "ID"
                ReturnValue = ""
                If Len(StrDeliveryRef) > 0 Then
                    FoundREF = Find_myQuery(frmMainGIForm.myConnString, DBInfoTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
                    If FoundREF Then
                        strDeliveryDate = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryDate")
                        strSupplier = GetMYValuebyFieldname(AllValues, AllFields, "Supplier_Name")
                        strGrid = GetMYValuebyFieldname(AllValues, AllFields, "Grid_No")
                        strLastSaved = GetMYValuebyFieldname(AllValues, AllFields, "Last_Saved")
                        strID = GetMYValuebyFieldname(AllValues, AllFields, "ID")
                        strCurrentStatus = GetMYValuebyFieldname(AllValues, AllFields, "STATUS")

                        Me.txtDeliveryDate.Text = strDeliveryDate
                        Me.txtDeliveryReference.Text = strDeliveryRef
                        Me.txtSupplier.Text = strSupplier
                        Me.txtGrid.Text = strGrid
                        Me.txtLastSaved.Text = strLastSaved
                        Me.txtID.Text = strID
                        Me.txtCurrentStatus.Text = strCurrentStatus
                        controlBACKCOLOR = Get_STATUS_BACKCOLOR(strCurrentStatus, controlFORECOLOR)
                        Me.txtCurrentStatus.BackColor = controlBACKCOLOR
                        Me.txtCurrentStatus.ForeColor = controlFORECOLOR
                    End If

                End If
            End If
            'OK we have StrDeliveryRef - find record and fill in the TEXT BOXES:

        End If
    End Sub



    Sub Check_If_Switch_To_NO_SHOW(RowIndex As Integer, ColumnIndex As Integer, Optional ByRef FormTAG As String = "")
        If RowIndex < 0 Or ColumnIndex < 0 Then
            Exit Sub
        Else

        End If
    End Sub

    Private Sub dgvDeliveryProgress_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDeliveryProgress.CellClick
        Call Check_If_Reference_Exists(e.RowIndex, e.ColumnIndex)
        'CALL to check if STATUS needs to be NO SHOW,CANCELLED or back to ROLLED ?

    End Sub

    Private Sub dgvDeliveryProgress_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDeliveryProgress.CellContentClick
        Call Check_If_Reference_Exists(e.RowIndex, e.ColumnIndex)

    End Sub

    Private Sub btnSearchEmpNo_Click(sender As Object, e As EventArgs) Handles btnSearchEmpNo.Click
        Dim strSQL As String
        Dim ProgressFields As String
        Dim NumRows As Long
        Dim ErrMessage As String

        strSQL = "SELECT * FROM tblDeliveryInfo ORDER BY Last_Saved DESC"


        'ProgressFields = "STATUS,DeliveryReference,ASN_Number,DeliveryDate,Last_Saved,UpdatedByName"
        'ProgressFields = ProgressFields & ",Supplier_Code,Supplier_Name"
        'ProgressFields = ProgressFields & ",SHIFT,Due_Time,Calc_Hours,Total_Hours,Pallets_Due,Cartons_Due,Origin"
        'ProgressFields = ProgressFields & ",Expected_Lines,Expected_Cases,Actual_Cases"
        'ProgressFields = ProgressFields & ",Estimated_Totes,Estimated_Cages,Estimated_Pallets,ArrivalTime,GRN_Number"

        ProgressFields = "STATUS,DeliveryReference,ASN_Number,DeliveryDate,Last_Saved,UpdatedByName"
        ProgressFields = ProgressFields & ",Grid_No,Total_Operatives,TotalShortRows,TotalExtraRows,Supplier_Code,Supplier_Name"
        ProgressFields = ProgressFields & ",SHIFT,Due_Time,Calc_Hours,Total_Hours,Pallets_Due,Cartons_Due,Origin"
        ProgressFields = ProgressFields & ",Expected_Lines,Expected_Cases,Actual_Cases"
        ProgressFields = ProgressFields & ",Estimated_Totes,Estimated_Cages,Estimated_Pallets,ArrivalTime,GRN_Number"

        If Len(Me.txtSearchASN.Text) = 0 And Len(Me.txtSearchRef.Text) = 0 Then
            strSQL = "SELECT " & ProgressFields & " FROM tblDeliveryInfo ORDER BY LastSaved DESC,SHIFT"
        End If
        If Len(Me.txtSearchRef.Text) > 0 Then
            strSQL = "SELECT " & ProgressFields & " FROM tblDeliveryInfo WHERE DeliveryReference = " & Me.txtSearchRef.Text
            strSQL = strSQL & " ORDER BY DeliveryDate DESC,SHIFT"
        End If
        If Len(Me.txtSearchASN.Text) > 0 Then
            strSQL = "SELECT " & ProgressFields & " FROM tblDeliveryInfo WHERE ASN_NUMBER = " & Me.txtSearchASN.Text
            strSQL = strSQL & " ORDER BY Last_Saved DESC,SHIFT"
        End If

        'strSQL = "SELECT " & ProgressFields & " FROM tblDeliveryInfo order by deliverydate desc,SHIFT"
        Call PopulateMyDataSource(dgvDeliveryProgress.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)


    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer_UpdateProgressGrid.Tick
        Dim SqlStr As String
        Dim NumRows As Integer
        Dim ErrMessage As String
        Dim ProgressFields As String
        Dim cbSelected As String
        Dim curRow As Integer
        Dim curCol As Integer
        Dim CellPoint As Point

        'MsgBox("HEY!")
        'message works
        Static iVertBar As Integer
        Static iHorizBar As Integer
        iVertBar = dgvDeliveryProgress.FirstDisplayedScrollingRowIndex '// get Top row.
        iHorizBar = dgvDeliveryProgress.FirstDisplayedScrollingColumnIndex '//get left column.
        'Form2.ShowDialog() '// Form that edits record.
        CellPoint = dgvDeliveryProgress.CurrentCellAddress
        curRow = CellPoint.Y
        curCol = CellPoint.X
        ProgressFields = "STATUS,DeliveryReference,ASN_Number,DeliveryDate,Last_Saved,UpdatedByName"
        ProgressFields = ProgressFields & ",Grid_No,Total_Operatives,TotalShortRows,TotalExtraRows,Supplier_Code,Supplier_Name"
        ProgressFields = ProgressFields & ",SHIFT,Due_Time,Calc_Hours,Total_Hours,Pallets_Due,Cartons_Due,Origin"
        ProgressFields = ProgressFields & ",Expected_Lines,Expected_Cases,Actual_Cases"
        ProgressFields = ProgressFields & ",Estimated_Totes,Estimated_Cages,Estimated_Pallets,ArrivalTime,GRN_Number"

        'cbSelected = WhatCBIsSelected()
        'If UCase(cbSelected) = "CBROLLED" Then
        'Criteria = "STATUS = " & Chr(39) & "ROLLED" & Chr(39)
        'ElseIf UCase(cbSelected) = "CBINPROGRESS" Then
        'Criteria = "STATUS = " & Chr(39) & "IN PROGRESS" & Chr(39)
        'ElseIf UCase(cbSelected) = "CBCOMPLETED" Then
        'Criteria = "STATUS = " & Chr(39) & "COMPLETED" & Chr(39)
        'Else
        'Criteria = ""
        'End If
        'If UCase(cbSelected) = "CBALL" Then
        'Criteria = ""
        'End If
        If Len(Criteria) = 0 Then
            SqlStr = "SELECT " & ProgressFields & " FROM tblDeliveryInfo order by Last_Saved desc,SHIFT"
            If TodayOnly Then
                SqlStr = "SELECT " & ProgressFields & " FROM tblDeliveryInfo "
                SqlStr = SqlStr & " WHERE last_saved >= TIMESTAMP(CURDATE()," & Chr(39) & "06:00:00" & Chr(39) & ")"
                'SqlStr = SqlStr & " AND last_saved < TIMESTAMP(CURDATE()," & Chr(39) & "06:00:00" & Chr(39) & ")"
                SqlStr = SqlStr & " order by Last_Saved desc"
            End If
        Else
            frmMainGIForm.txtMessages.Text = "CRITERIA CHANGED: " & Criteria
            SqlStr = "SELECT " & ProgressFields & " FROM tblDeliveryInfo WHERE " & Criteria
            If TodayOnly Then
                SqlStr = SqlStr & " AND last_saved >= TIMESTAMP(CURDATE()," & Chr(39) & "06:00:00" & Chr(39) & ")"
                'SqlStr = SqlStr & " AND last_saved < TIMESTAMP(CURDATE()," & Chr(39) & "06:00:00" & Chr(39) & ")"
            End If
            SqlStr = SqlStr & " order by Last_Saved desc"
        End If
        'WHERE STATUS = Rolled / In Progress / Completed
        'Invoked by user selecting dropdown filter.
        ErrMessage = ""
        'Error generated if no records are returned:
        Call PopulateMyDataSource(dgvDeliveryProgress.DataSource, frmMainGIForm.myConnString, SqlStr, NumRows, ErrMessage)
        Me.txtNumRows.Text = CStr(NumRows)
        If iVertBar > -1 And iVertBar <= NumRows Then
            dgvDeliveryProgress.FirstDisplayedScrollingRowIndex = iVertBar '// set Top row.
        End If
        If iHorizBar > -1 Then
            dgvDeliveryProgress.FirstDisplayedScrollingColumnIndex = iHorizBar '// set Top row.
        End If
        If curRow >= 0 And curRow <= NumRows Then
            dgvDeliveryProgress.CurrentCell = dgvDeliveryProgress.Rows(curRow).Cells(curCol)
        End If

    End Sub

    Function WhatCBIsSelected(Optional ByVal grp As GroupBox = Nothing) As String
        Dim cbox As CheckBox
        Dim cboxName As String = String.Empty
        Try
            Dim ctl As Control
            For Each ctl In pnlTitle.Controls
                If TypeOf ctl Is CheckBox Then
                    cbox = DirectCast(ctl, CheckBox)
                    If cbox.Checked Then
                        cboxName = cbox.Name
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            Dim stackframe As New Diagnostics.StackFrame(1)
            Throw New Exception("An error occurred in WhatCBIsSelected, '" & stackframe.GetMethod.ReflectedType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "'." & Environment.NewLine & "  Message was: '" & ex.Message & "'")
        End Try
        Return cboxName
    End Function

    Private Sub dgvDeliveryProgress_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvDeliveryProgress.CellFormatting
        If Not e.RowIndex = dgvDeliveryProgress.NewRowIndex Then
            If e.ColumnIndex = dgvDeliveryProgress.Columns("Status").Index Then
                If UCase(e.Value.ToString) = "ROLLED" Then
                    e.CellStyle.BackColor = Color.Red
                End If
                If UCase(e.Value.ToString) = "ROLLED OVER" Then
                    e.CellStyle.BackColor = Color.Red
                End If
                If UCase(e.Value.ToString) = "IN PROGRESS" Then
                    e.CellStyle.BackColor = Color.Orange
                End If
                If UCase(e.Value.ToString) = "COMPLETED" Then
                    e.CellStyle.BackColor = Color.LimeGreen
                End If
                If UCase(e.Value.ToString) = "CANCELLED" Then
                    e.CellStyle.BackColor = Color.Maroon
                End If
                If UCase(e.Value.ToString) = "NO SHOW" Then
                    e.CellStyle.BackColor = Color.Crimson
                End If
            End If
        End If
    End Sub

    Private Sub rbALL_CheckedChanged(sender As Object, e As EventArgs) Handles rbALL.Click, rbALL.CheckedChanged
        'No Filter
        frmMainGIForm.txtMessages.Text = "ALL"
        Criteria = ""
    End Sub

    Private Sub rbRolled_CheckedChanged(sender As Object, e As EventArgs) Handles rbRolled.Click, rbRolled.CheckedChanged
        frmMainGIForm.txtMessages.Text = "ROLLED"
        Criteria = "STATUS = " & Chr(39) & "ROLLED" & Chr(39)
    End Sub

    Private Sub rbInProgress_CheckedChanged(sender As Object, e As EventArgs) Handles rbInProgress.Click, rbInProgress.CheckedChanged
        frmMainGIForm.txtMessages.Text = "IN PROGRESS"
        Criteria = "STATUS = " & Chr(39) & "IN PROGRESS" & Chr(39)
    End Sub

    Private Sub rbCompleted_CheckedChanged(sender As Object, e As EventArgs) Handles rbCompleted.Click, rbCompleted.CheckedChanged
        frmMainGIForm.txtMessages.Text = "COMPLETED"
        Criteria = "STATUS = " & Chr(39) & "COMPLETED" & Chr(39)
    End Sub

    Private Sub cbALL_CheckedChanged(sender As Object, e As EventArgs)
        Criteria = ""
        'cbALL.CheckState = CheckState.Checked
        'cbRolled.CheckState = CheckState.Unchecked
        'cbInProgress.CheckState = CheckState.Unchecked
        'cbCompleted.CheckState = CheckState.Unchecked
    End Sub

    Private Sub cbRolled_CheckedChanged(sender As Object, e As EventArgs)
        Criteria = "STATUS = " & Chr(39) & "ROLLED" & Chr(39)
        'cbRolled.CheckState = CheckState.Checked
        'cbALL.CheckState = CheckState.Unchecked
        'cbInProgress.CheckState = CheckState.Unchecked
        'cbCompleted.CheckState = CheckState.Unchecked
    End Sub

    Private Sub cbInProgress_CheckedChanged(sender As Object, e As EventArgs)
        Criteria = "STATUS = " & Chr(39) & "IN PROGRESS" & Chr(39)
        'cbInProgress.CheckState = CheckState.Checked
        'cbRolled.CheckState = CheckState.Unchecked
        'cbALL.CheckState = CheckState.Unchecked
        'cbCompleted.CheckState = CheckState.Unchecked
    End Sub

    Private Sub cbCompleted_CheckedChanged(sender As Object, e As EventArgs)
        Criteria = "STATUS = " & Chr(39) & "COMPLETED" & Chr(39)
        'cbCompleted.CheckState = CheckState.Checked
        'cbInProgress.CheckState = CheckState.Unchecked
        'cbRolled.CheckState = CheckState.Unchecked
        'cbALL.CheckState = CheckState.Unchecked
    End Sub

    Private Sub cbRolled_Click(sender As Object, e As EventArgs)
        Criteria = "STATUS = " & Chr(39) & "ROLLED" & Chr(39)
        'cbRolled.CheckState = CheckState.Checked
        'cbALL.CheckState = CheckState.Unchecked
        'cbInProgress.CheckState = CheckState.Unchecked
        'cbCompleted.CheckState = CheckState.Unchecked
    End Sub

    Private Sub rbALL_CheckedChanged_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub cbToday_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnExportToCSV_Click(sender As Object, e As EventArgs) Handles btnExportToCSV.Click
        'EXPORT GRID CONTENTS TO CSV FILE:
        Dim ExportOK As Boolean = False
        Dim TodayDate As String = ""

        TodayDate = Now().ToString("dd_MMM_yyyy HHMM")
        ExportOK = Module_DanG_MySQL_Tools.ExportDGVToCSVFile(Me.dgvDeliveryProgress, "VB_TIMESHEET_References_" & TodayDate)
    End Sub

    Private Sub btnTodayOnly_Click(sender As Object, e As EventArgs) Handles btnTodayOnly.Click
        If UCase(Me.btnTodayOnly.Text) = "TODAY ONLY" Then
            TodayOnly = True
            Me.btnTodayOnly.Text = "SHOW ALL"
        Else
            TodayOnly = False
            Me.btnTodayOnly.Text = "TODAY ONLY"
        End If
    End Sub

    Private Sub frmReferenceProgress_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim RadioButtonBackColor As System.Drawing.Color
        Dim RadioButtonForeColor As System.Drawing.Color
        Dim ListArray() As String
        Dim strSQL As String
        Dim Criteria As String
        Dim ErrMessage As String
        Dim ImportButtonColor As String
        Dim UpdatedToday As Boolean

        Me.btnTodayOnly.Text = "SHOW ALL"
        TodayOnly = True
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("ALL", RadioButtonForeColor)
        Me.rbALL.BackColor = RadioButtonBackColor
        Me.rbALL.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("ROLLED", RadioButtonForeColor)
        Me.rbRolled.BackColor = RadioButtonBackColor
        Me.rbRolled.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("ROLLED OVER", RadioButtonForeColor)
        Me.rbRolledOver.BackColor = RadioButtonBackColor
        Me.rbRolledOver.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("NO SHOW", RadioButtonForeColor)
        Me.rbNoShow.BackColor = RadioButtonBackColor
        Me.rbNoShow.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("CANCELLED", RadioButtonForeColor)
        Me.rbCancelled.BackColor = RadioButtonBackColor
        Me.rbCancelled.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("IN PROGRESS", RadioButtonForeColor)
        Me.rbInProgress.BackColor = RadioButtonBackColor
        Me.rbInProgress.ForeColor = RadioButtonForeColor
        RadioButtonBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR("COMPLETED", RadioButtonForeColor)
        Me.rbCompleted.BackColor = RadioButtonBackColor
        Me.rbCompleted.ForeColor = RadioButtonForeColor
        ReDim ListArray(1)
        strSQL = "SELECT * FROM tblStatusList ORDER BY STATUS"
        Criteria = ""
        ErrMessage = ""
        Get_DropdownItems(frmMainGIForm.myConnString, strSQL, Criteria, ErrMessage, ListArray, 1)
        Me.comSTATUS.Items.AddRange(ListArray)

        UpdatedToday = Check_ImportDate_IsToday("tblDeliveryInfo", "DeliveryDate", frmMainGIForm.myConnString, ImportButtonColor, ErrMessage)
        If UCase(ImportButtonColor) = "RED" Then
            'SELECT UNSELECTED INDEX appropriate for RED from IMAGE LIST:
            Me.btnImportData.BackgroundImage = imglistImportData.Images.Item(6)
        ElseIf UCase(ImportButtonColor) = "AMBER" Then
            'SELECT UNSELECTED INDEX appropriate for RED from IMAGE LIST:
            Me.btnImportData.BackgroundImage = imglistImportData.Images.Item(3)
        ElseIf UCase(ImportButtonColor) = "GREEN" Then
            'SELECT UNSELECTED INDEX appropriate for RED from IMAGE LIST:
            Me.btnImportData.BackgroundImage = imglistImportData.Images.Item(0)
        Else
            'UNKNOWN COLOUR:
            Me.btnImportData.BackgroundImage = imglistImportData.Images.Item(8)
        End If

    End Sub

    Private Sub NOSHOWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NOSHOWToolStripMenuItem.Click
        'SENDER should be GRID.
        'Need to get which Row that was selected - find the REFERENCE it will belong to.
        'NO SHOW needs inserting into appropriate DB field and record containing that ID (after SEARCH QUERY).
        Dim REF As String
        Dim MenuItem As String
        Dim CellPoint As Point
        Dim curRow As Integer
        Dim curCol As Integer

        CellPoint = dgvDeliveryProgress.CurrentCellAddress
        curRow = CellPoint.Y
        curCol = CellPoint.X
        REF = dgvDeliveryProgress.Rows(curRow).Cells(curCol).ToString
        MsgBox("Insert NO SHOW: " & REF)

    End Sub

    Private Sub CANCELLEDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CANCELLEDToolStripMenuItem.Click
        'SENDER should be GRID.
        'Need to get which Row that was selected - find the REFERENCE it will belong to.
        'CANCELLED needs inserting into appropriate DB field and record containing that ID (after SEARCH QUERY).
        Me.txtNewStatus.Text = "CANCELLED"
        Call SetStatus()

    End Sub

    Private Sub dgvDeliveryProgress_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDeliveryProgress.CellMouseClick
        If UCase(e.Button) = "RIGHT" And e.RowIndex >= 0 Then
            dgvDeliveryProgress.Rows(e.RowIndex).Selected = True
        End If
    End Sub

    Private Sub dgvDeliveryProgress_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDeliveryProgress.CellMouseDown
        If UCase(e.Button) = "RIGHT" And e.RowIndex >= 0 Then
            dgvDeliveryProgress.Rows(e.RowIndex).Selected = True
        End If
    End Sub

    Private Sub btnSetNewStatus_Click(sender As Object, e As EventArgs)
        'ASSIGN NEW STATUS in txtNewStatus box to DB STATUS field : i.e. perform UPDATE on found record in GRID after user has clicked on desired record:



    End Sub

    Private Sub btnSetNewStatus_Click_1(sender As Object, e As EventArgs) Handles btnSetNewStatus.Click

        Call SetStatus()

    End Sub

    Sub SetStatus()
        Dim strNewStatus As String
        Dim strDeliveryRef As String
        Dim ErrMessages As String

        ErrMessages = ""
        strNewStatus = Me.txtNewStatus.Text
        strDeliveryRef = Me.txtDeliveryReference.Text
        If Len(strNewStatus) = 0 Then
            MsgBox("No Status Specified")
            Exit Sub
        End If
        If Len(strDeliveryRef) = 0 Then
            MsgBox("Delivery Ref is EMPTY")
            Exit Sub
        End If
        Call Change_STATUS(strDeliveryRef, strNewStatus, ErrMessages)


    End Sub

    Private Sub STATUSMenuStrip_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles STATUSMenuStrip.Opening

    End Sub

    Private Sub ROLLEDOVERToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ROLLEDOVERToolStripMenuItem1.Click
        '
        Me.txtNewStatus.Text = "ROLLED OVER"
        Call SetStatus()
    End Sub

    Private Sub rbRolledOver_Click(sender As Object, e As EventArgs) Handles rbRolledOver.Click
        Criteria = "STATUS = " & Chr(39) & "ROLLED OVER" & Chr(39)
    End Sub

    Private Sub rbNoShow_Click(sender As Object, e As EventArgs) Handles rbNoShow.Click
        Criteria = "STATUS = " & Chr(39) & "NO SHOW" & Chr(39)
    End Sub

    Private Sub rbCancelled_Click(sender As Object, e As EventArgs) Handles rbCancelled.Click
        Criteria = "STATUS = " & Chr(39) & "CANCELLED" & Chr(39)
    End Sub

    Private Sub comSTATUS_SelectedValueChanged(sender As Object, e As EventArgs) Handles comSTATUS.SelectedValueChanged
        Dim txtBackColor As System.Drawing.Color
        Dim txtForeColor As System.Drawing.Color
        Dim ItemSelected As String

        ItemSelected = comSTATUS.Items(comSTATUS.SelectedIndex)
        Me.txtNewStatus.Text = ItemSelected
        txtBackColor = Module_TIMESHEET.Get_STATUS_BACKCOLOR(ItemSelected, txtForeColor)
        Me.txtNewStatus.BackColor = txtBackColor
        Me.txtNewStatus.ForeColor = txtForeColor

    End Sub

    Private Sub btnImportData_MouseHover(sender As Object, e As EventArgs) Handles btnImportData.MouseHover
        Dim ReturnColour As String
        Dim ErrMessage As String

        ReturnColour = Module_TIMESHEET.Check_ImportDate_IsToday("tblDeliveryInfo", "DeliveryDate", frmMainGIForm.myConnString, ReturnColour, ErrMessage)
        If UCase(ReturnColour) = "RED" Then
            'RED: INDEX 6-8 need pushed image
            btnImportData.BackgroundImage = imglistImportData.Images.Item(7)
        ElseIf UCase(ReturnColour) = "AMBER" Then
            'AMBER: INDEX 3-5
            btnImportData.BackgroundImage = imglistImportData.Images.Item(4)
        ElseIf UCase(ReturnColour) = "GREEN" Then
            'GREEN: INDEX 0-2
            btnImportData.BackgroundImage = imglistImportData.Images.Item(1)
        Else
            'unknown
            btnImportData.BackgroundImage = imglistImportData.Images.Item(7)
        End If

    End Sub

    Private Sub btnImportData_Click(sender As Object, e As EventArgs) Handles btnImportData.Click

        Call ImportReferenceData()

    End Sub

    Sub ImportReferenceData()
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

        'dtDateSelected = CDate(Me.txtSelectDeliveryDate.Text)
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

    Private Sub btnImportData_MouseLeave(sender As Object, e As EventArgs) Handles btnImportData.MouseLeave
        Dim ReturnColour As String
        Dim ErrMessage As String

        ReturnColour = Module_TIMESHEET.Check_ImportDate_IsToday("tblDeliveryInfo", "DeliveryDate", frmMainGIForm.myConnString, ReturnColour, ErrMessage)
        If UCase(ReturnColour) = "RED" Then
            'RED: INDEX 6-8 need pushed image
            btnImportData.BackgroundImage = imglistImportData.Images.Item(6)
        ElseIf UCase(ReturnColour) = "AMBER" Then
            'AMBER: INDEX 3-5
            btnImportData.BackgroundImage = imglistImportData.Images.Item(3)
        ElseIf UCase(ReturnColour) = "GREEN" Then
            'GREEN: INDEX 0-2
            btnImportData.BackgroundImage = imglistImportData.Images.Item(0)
        Else
            'unknown
            btnImportData.BackgroundImage = imglistImportData.Images.Item(7)
        End If
    End Sub
End Class