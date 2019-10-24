Public Class frmReferenceProgress
    Private Sub frmReferenceProgress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SETUP VIEW GRID:
        Dim SqlStr As String
        Dim NumRows As Integer
        Dim ErrMessage As String
        Dim ProgressFields As String

        ProgressFields = "STATUS,GRID_SHIFT,Supplier_Code,Supplier_Name,DeliveryDate,DeliveryReference,Pallets_Due,Cartons_Due,Origin"
        ProgressFields = ProgressFields & ",ASN_Number,SHIFT,Due_Time,Expected_Lines,Expected_Cases,Actual_Cases"
        ProgressFields = ProgressFields & ",Estimated_Totes,Estimated_Cages,Estimated_Pallets,Calc_Hours"

        SqlStr = "SELECT " & ProgressFields & " FROM tblDeliveryInfo order by deliverydate desc,SHIFT"
        Call PopulateMyDataSource(dgvDeliveryProgress.DataSource, frmMainGIForm.myConnString, SqlStr, NumRows, ErrMessage)

    End Sub

    Private Sub dgvDeliveryProgress_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDeliveryProgress.CellClick
        'OK user clicks on a reference.
        'Get that reference.
        'Check if a TIMESHEET child form is already open ?
        'Check WHAT reference it is displaying ???
        'If it does NOT match the current reference just clicked - then open up a NEW TIMESHEET child form and display the NEW reference.

    End Sub
End Class