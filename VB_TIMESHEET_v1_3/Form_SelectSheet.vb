Public Class frmSelectSheet
    Private _WorksheetName As String
    Private _SheetList As New List(Of String)
    Public WorkbookFilename As String
    Public DeleteAllRecords As Boolean = False
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
        Dim DeleteAllEmployeesFirst As Boolean = False
        Dim TotalUpdated As Long
        Dim TotalInserts As Long
        Dim AllowSaves As Boolean = True
        Dim TotalDescriptionsChanged As Long
        Dim IncludeTGWNames As Boolean = False

        'OK BUTTON
        'REturn item selected BACK to calling program:
        Me.SelectedWorksheet = Me.txtSelectedSheet.Text
        frmMainGIForm.SelectedSheet = Me.SelectedWorksheet
        frmMainGIForm.SelectedWorkbookFile = Me.txtWorkbook.Text
        If Me.lblConfirmUpdates.ForeColor = Color.Red Then
            DeleteAllRecords = True
            DeleteAllEmployeesFirst = True
        Else
            DeleteAllRecords = False
            DeleteAllEmployeesFirst = False
        End If

        'Me.Close()
        Call UpdateEmployees(frmMainGIForm.ErrList, DeleteAllEmployeesFirst, TotalUpdated, TotalInserts, AllowSaves, TotalDescriptionsChanged, IncludeTGWNames)
        Me.txtTotalUpdates.Text = "Total Updated: " & CStr(TotalUpdated) & ", Descriptions changed: " & CStr(TotalDescriptionsChanged)



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
        Dim ErrorList() As String
        Dim TotalUpdates As Long
        Dim TotalDescUpdates As Long
        Dim TotalInsertsNeeded As Long
        Dim IncludeTGW As Boolean = False

        TotalUpdates = 0
        TotalDescUpdates = 0
        ReDim ErrorList(1)
        Me.txtSelectedSheet.Text = lstSheets.Text

        frmMainGIForm.SelectedWorkbookFile = Me.txtWorkbook.Text
        frmMainGIForm.SelectedSheet = Me.txtSelectedSheet.Text

        'Call procedure to calc number of changes:
        'Call UpdateEmployees(ErrorList, False, TotalUpdates, False, TotalDescUpdates)
        Call UpdateNames_Using_TABLES(ErrorList, False, TotalUpdates, TotalInsertsNeeded, True, TotalDescUpdates, IncludeTGW)

        Me.txtTotalUpdates.Visible = True
        Me.lblTotalUpdates.Visible = True
        Me.txtTotalUpdates.Text = CStr(TotalUpdates) & " , Total INSERTS: " & TotalInsertsNeeded & ", #Desc= " & CStr(TotalDescUpdates)

    End Sub

    Private Sub frmSelectSheet_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'LOAD EVENT:
        'GET the DEFAULT EXCEL PATH to the Employee Spreadsheet location ?
        'OR NEEDS TO BE IN THE SHOWN() EVENT ???

    End Sub

    Private Sub lblConfirmUpdates_Click(sender As Object, e As EventArgs) Handles lblConfirmUpdates.Click
        If lblConfirmUpdates.ForeColor = Color.Red Then
            'Turn off DELETE ALL RECORDS:
            lblConfirmUpdates.ForeColor = Color.LawnGreen
        Else
            lblConfirmUpdates.ForeColor = Color.Red
        End If
    End Sub
End Class