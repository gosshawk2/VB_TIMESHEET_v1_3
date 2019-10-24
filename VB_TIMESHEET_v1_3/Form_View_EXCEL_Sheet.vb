Public Class frmViewExcelSheet
    Dim ExcelFilename As String
    Private Sub btnGetExcelSheet_Click(sender As Object, e As EventArgs) Handles btnGetExcelSheet.Click
        'Get EXCEL DATA from selected Workbook:

        Dim SheetNames() As String
        Dim StartCell As String
        Dim EndCell As String
        Dim IDX As Long
        'Dim dt As DataTable
        'Dim ExtractArr(,) As Object
        Me.txtMessages.Visible = True
        Me.btnFinish.Visible = False
        Me.txtSelectedTab.Visible = False
        Me.lblSelectedTab.Visible = False
        Me.txtMessages.Text = "Please Wait ... Getting Excel TABS"
        ReDim SheetNames(1)
        ExcelFilename = BrowseFilename(1)
        Me.txtFilename.Text = ExcelFilename
        If Len(ExcelFilename) > 0 Then
            Me.lstExcelTabs.Visible = True
            Call GetExcelSheets(ExcelFilename, SheetNames)
            IDX = 0
            Do While IDX < UBound(SheetNames)
                If Not IsNothing(SheetNames(IDX)) Then
                    lstExcelTabs.Items.Add(SheetNames(IDX))
                End If
                IDX = IDX + 1
            Loop
        End If
        Me.txtMessages.Text = "Total Excel TABS = " & CStr(IDX - 1)
        'PLEASE WAIT ...
        'SheetName = "Daily"
        'StartCell = "A1"
        'EndCell = "U2979"
        'ExtractArr = ExtractExcelRows(ExcelFilename, SheetName, 1, 1, 0, 0)
        'Me.dgvExcelSheet.DataSource = dt
    End Sub

    Private Sub frmViewExcelSheet_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.txtTITLE.Width = Me.pnlTop.Width

    End Sub

    Private Sub lstExcelTabs_Click(sender As Object, e As EventArgs) Handles lstExcelTabs.Click
        Dim ItemChosen As String

        ItemChosen = lstExcelTabs.SelectedItem.ToString
        Me.lblSelectedTab.Visible = True
        Me.txtSelectedTab.Visible = True
        Me.txtSelectedTab.Text = ItemChosen
        Me.btnFinish.Visible = True
        ' frmMainGIForm
        frmMainGIForm.ListSelectedItem = Me.txtSelectedTab.Text
        frmMainGIForm.WorkbookPath = Me.txtFilename.Text
    End Sub

    Private Sub btnFinish_Click(sender As Object, e As EventArgs) Handles btnFinish.Click
        Me.Close()
    End Sub

    Private Sub lstExcelTabs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstExcelTabs.SelectedIndexChanged
        Dim ItemChosen As String
        'This works too
        ItemChosen = lstExcelTabs.SelectedItem.ToString
        Me.lblSelectedTab.Visible = True
        Me.txtSelectedTab.Visible = True
        Me.txtSelectedTab.Text = ItemChosen
        Me.btnFinish.Visible = True
        ' frmMainGIForm
        frmMainGIForm.ListSelectedItem = Me.txtSelectedTab.Text
        frmMainGIForm.WorkbookPath = Me.txtFilename.Text
    End Sub

    Private Sub lstExcelTabs_SelectedValueChanged(sender As Object, e As EventArgs) Handles lstExcelTabs.SelectedValueChanged
        Dim ItemChosen As String

        ItemChosen = lstExcelTabs.SelectedItem.ToString
        Me.lblSelectedTab.Visible = True
        Me.txtSelectedTab.Visible = True
        Me.txtSelectedTab.Text = ItemChosen
        Me.btnFinish.Visible = True
        ' frmMainGIForm
        frmMainGIForm.ListSelectedItem = Me.txtSelectedTab.Text
        frmMainGIForm.WorkbookPath = Me.txtFilename.Text
    End Sub

    Private Sub lstExcelTabs_MouseClick(sender As Object, e As MouseEventArgs) Handles lstExcelTabs.MouseClick
        Dim ItemChosen As String
        Dim dt As New DataTable
        Dim ExtractArr(,) As Object
        Dim Sheetname As String

        'try this: YES THIS WORKS ! 
        ReDim ExtractArr(1, 1)
        ItemChosen = lstExcelTabs.SelectedItem.ToString
        'MsgBox(ItemChosen)
        Me.lblSelectedTab.Visible = True
        Me.txtSelectedTab.Visible = True
        Me.txtMessages.Visible = True
        Me.txtSelectedTab.Text = ItemChosen
        Me.btnFinish.Visible = True
        ' frmMainGIForm
        frmMainGIForm.ListSelectedItem = Me.txtSelectedTab.Text
        frmMainGIForm.WorkbookPath = Me.txtFilename.Text
        Sheetname = ItemChosen
        If Len(ExcelFilename) > 0 Then
            ExtractArr = ExtractExcelRows(ExcelFilename, Sheetname, 1, 1, 0, 0, dt)
            '
            If Not IsNothing(ExtractArr) Then
                Me.txtMessages.Text = "Total Rows: " & CStr(UBound(ExtractArr))
                dgvExcelSheet.Rows.Clear()
                Me.TabPage1.Text = Sheetname
                If Not IsNothing(dt) Then

                    dgvExcelSheet.DataSource = dt
                    'dgvExcelSheet.databind() 'not a member aparently.
                End If
            Else
                Me.txtMessages.Text = "Total Rows: " & CStr(0)
            End If
        End If
    End Sub
End Class