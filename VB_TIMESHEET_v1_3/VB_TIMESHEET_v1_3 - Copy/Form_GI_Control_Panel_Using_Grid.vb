Public Class frmGI_ControlPanel_Using_Grid
    Private Sub btnAddOperative_Click(sender As Object, e As EventArgs) Handles btnAddOperative.Click
        'Add Operative row to DataGridView - dgvOperatives:
        ADD_ROW_TO_DATA_GRID_VIEW_OPERATIVES(Me.dgvOperatives, 5)

    End Sub

    Private Sub dgvOperatives_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOperatives.CellContentClick
        Dim colName As String = dgvOperatives.Columns(e.ColumnIndex).Name

        MsgBox(colName)

    End Sub
End Class