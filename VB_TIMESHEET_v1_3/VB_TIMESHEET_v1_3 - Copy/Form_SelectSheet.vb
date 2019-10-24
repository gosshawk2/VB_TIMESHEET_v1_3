Public Class frmSelectSheet
    Private _WorksheetName As String
    Private _SheetList As New List(Of String)
    Public WorkbookFilename As String
    Public Property SelectedWorksheet() As String
        Get
            Return _WorksheetName
        End Get
        Set(value As String)
            _WorksheetName = value
        End Set
    End Property

    Public Property DropdownItems() As List(Of String)
        Get
            Return _SheetList
        End Get
        Set(value As List(Of String))
            _SheetList = value
        End Set
    End Property

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim Messages As String = ""
        'OK BUTTON
        'REturn item selected BACK to calling program:
        Me.SelectedWorksheet = Me.txtSelectedSheet.Text
        frmMainGIForm.SelectedSheet = Me.SelectedWorksheet
        frmMainGIForm.SelectedWorkbookFile = Me.txtWorkbook.Text
        Me.Close()
        Call UpdateEmployees(frmMainGIForm.ErrList)



    End Sub

    Private Sub frmSelectSheet_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim ListItem As String

        If Not Me.DropdownItems Is Nothing Then
            For Each ListItem In DropdownItems
                Me.lstSheets.Items.Add(ListItem)
            Next
        End If

        Me.txtWorkbook.Focus()


    End Sub

    Private Sub btnSelectWorkbook_Click(sender As Object, e As EventArgs) Handles btnSelectWorkbook.Click
        'Shower Browser and get workbook filename:
        Dim SheetList As New List(Of String)
        Dim SheetsArr() As String
        Dim SheetName As String
        Dim IDX As Long


        WorkbookFilename = BrowseFilename(1)
        If Len(WorkbookFilename) = 0 Then
            Exit Sub
        End If
        Me.txtWorkbook.Text = WorkbookFilename
        frmMainGIForm.txtMessages.Text = "Please Wait ... Opening Workbook and Extract Sheets"
        ReDim SheetsArr(1)
        Call DanG_EXCEL_Module.GetExcelSheets(WorkbookFilename, SheetsArr)
        For IDX = 0 To UBound(SheetsArr)
            SheetName = SheetsArr(IDX)
            If SheetName IsNot Nothing And Len(SheetName) > 1 Then
                lstSheets.Items.Add(SheetName)
            End If
        Next

    End Sub

    Private Sub lstSheets_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSheets.SelectedIndexChanged
        Me.txtSelectedSheet.Text = lstSheets.Text

    End Sub

    Private Sub frmSelectSheet_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'LOAD
    End Sub
End Class