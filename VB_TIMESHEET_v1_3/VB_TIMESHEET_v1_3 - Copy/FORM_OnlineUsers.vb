Public Class frmOnlineUsers
    Private Sub frmOnlineUsers_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim strSQL As String
        Dim NumRows As Long
        Dim ErrMessage As String

        strSQL = "SELECT Firstname,Lastname,LogInDatetime,EmpNo,ComputerName,Location FROM tblOperatorsOnline"
        Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)
        Me.dgvOnlineUsers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim strSQL As String
        Dim NumRows As Long
        Dim ErrMessage As String

        strSQL = "SELECT Firstname,Lastname,LogInDatetime,EmpNo,ComputerName,Location FROM tblOperatorsOnline"
        Call PopulateMyDataSource(Me.dgvOnlineUsers.DataSource, frmMainGIForm.myConnString, strSQL, NumRows, ErrMessage)
        Me.dgvOnlineUsers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
    End Sub
End Class