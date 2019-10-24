Public Class frmViewDataInGrid
    Public DBTAble As String
    Public criteria As String
    Public DisplayFields As String
    Public SortFields As String
    Private Sub frmViewDataInGrid_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------------------------------***************** LOAD ***************************************
        comDBTables.Items.Add("Delivery Info")
        comDBTables.Items.Add("Labour Hours")
        comDBTables.Items.Add("Operatives")
        comDBTables.Items.Add("Daily")
        comDBTables.Items.Add("Short Parts")
        comDBTables.Items.Add("Extra Parts")
        comDBTables.Items.Add("Supplier Compliance")
        criteria = ""
    End Sub

    Private Sub comDBTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comDBTables.SelectedIndexChanged
        If comDBTables.Text = "Delivery Info" Then
            DBTAble = "tblDeliveryInfo"
            DisplayFields = ""
            SortFields = ""
            criteria = ""
        End If
        If comDBTables.Text = "Labour Hours" Then
            DBTAble = "tblLabourHours"
            DisplayFields = ""
            SortFields = ""
            criteria = ""
        End If
        If comDBTables.Text = "Operatives" Then
            DBTAble = "tblOperatives"
            DisplayFields = ""
            SortFields = ""
            criteria = ""
        End If
        If comDBTables.Text = "Daily" Then
            DBTAble = "tblDailySheet"
            DisplayFields = ""
            SortFields = ""
            criteria = ""
        End If
        If comDBTables.Text = "Short Parts" Then
            DBTAble = "tblShortsAndExtraParts"
            criteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
            DisplayFields = ""
            SortFields = ""
        End If
        If comDBTables.Text = "Extra Parts" Then
            DBTAble = "tblShortsAndExtraParts"
            criteria = "ShortOrExtra = " & Chr(34) & "EXTRA" & Chr(34)
            DisplayFields = ""
            SortFields = ""
        End If
        If comDBTables.Text = "Supplier Compliance" Then
            DBTAble = "tblSupplierCompliance"
            DisplayFields = ""
            SortFields = ""
            criteria = ""
        End If

    End Sub

    Private Sub btnViewDBTable_Click(sender As Object, e As EventArgs) Handles btnViewDBTable.Click
        'VIEW TABLE IN GRID: dgv
        Dim strSQL As String
        Dim NumRows As Long

        If Len(DBTAble) > 0 Then
            If Len(DisplayFields) = 0 Then
                strSQL = "SELECT * FROM " & DBTAble
            Else
                strSQL = "SELECT " & DisplayFields & " FROM " & DBTAble
            End If
            If Len(criteria) > 0 Then
                strSQL = strSQL & " WHERE " & criteria
            End If
            If Len(SortFields) > 0 Then
                strSQL = strSQL & " ORDER BY " & SortFields
            End If
            If Me.dgv.RowCount > 0 Then
                'Me.dgv.Rows.Clear()
            End If

            Call Module_DanG_MySQL_Tools.PopulateMyDataSource(dgv.DataSource, frmMainGIForm.myConnString, strSQL, NumRows)
            Me.txtTotalRecs.Text = CStr(NumRows)

        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnExportToCSV_Click(sender As Object, e As EventArgs) Handles btnExportToCSV.Click
        'EXPORT DATA IN GRID TO CSV FILE:
        Dim ExportOK As Boolean
        Dim CSVFilename As String
        Dim TodayDate As String = ""

        TodayDate = Now().ToString("dd_MMM_yyyy HHMM")
        ExportOK = Module_DanG_MySQL_Tools.ExportDGVToCSVFile(Me.dgv, comDBTables.Text & TodayDate)

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        'SEARCH entry - ASN or Delivery Reference:

        If Len(Me.txtSearchASN) > 0 Then
            'Search ASN No.

        End If
        If Len(Me.txtSearchRef) > 0 Then
            'Search Delivery Reference:

        End If
    End Sub

    Private Sub btnColSave_Click(sender As Object, e As EventArgs) Handles btnColSave.Click
        'SAVE COLUMN POSITION to tblPreferences:
        Dim Headings As Object
        Dim strHeadings As String

        ReDim Headings(1)
        strHeadings = ExtractHeadings(dgv, Headings)
        'Save to tblPreferences under the current username and employee number:


    End Sub
End Class