Public Class frmViewDataInGrid
    Public DBTAble As String
    Public criteria As String
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
        End If
        If comDBTables.Text = "Labout Hours" Then
            DBTAble = "tblLabourHours"
        End If
        If comDBTables.Text = "Operatives" Then
            DBTAble = "tblOperatives"
        End If
        If comDBTables.Text = "Daily" Then
            DBTAble = "tblDailySheet"
        End If
        If comDBTables.Text = "Short Parts" Then
            DBTAble = "tblShortsAndExtraParts"
            criteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
        End If
        If comDBTables.Text = "Extra Parts" Then
            DBTAble = "tblShortsAndExtraParts"
            criteria = "ShortOrExtra = " & Chr(34) & "SHORT" & Chr(34)
        End If
        If comDBTables.Text = "Supplier Compliance" Then
            DBTAble = "tblSupplierCompliance"
        End If
        Me.txtTable.Text = DBTAble

    End Sub

    Private Sub btnViewDBTable_Click(sender As Object, e As EventArgs) Handles btnViewDBTable.Click
        'VIEW TABLE IN GRID: dgv
        Dim strSQL As String
        Dim NumRows As Long

        If Len(DBTAble) > 0 Then
            strSQL = "SELECT * FROM " & DBTAble
            If Len(criteria) > 0 Then
                strSQL = strSQL & " WHERE " & criteria
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
End Class