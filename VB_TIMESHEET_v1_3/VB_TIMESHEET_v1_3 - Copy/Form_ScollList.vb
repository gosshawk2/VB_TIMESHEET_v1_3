Public Class frmScrollList
    Private Sub frmScrollList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Messages As String
        Dim IDX As Long

        For IDX = 0 To UBound(frmMainGIForm.ErrList)
            Messages = frmMainGIForm.ErrList(IDX)
            Me.txtOutput.AppendText(Messages)
        Next
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class