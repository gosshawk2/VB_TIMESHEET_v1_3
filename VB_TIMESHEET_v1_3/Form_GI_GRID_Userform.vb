Public Class frm_GI_GRID
    Private Sub btnSaveAndContinue_Click(sender As Object, e As EventArgs) Handles btnSaveAndContinue.Click
        'SAVE AND CONTINUE

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'RESET CONTROLS

    End Sub

    Private Sub btnLoadDeliveryRef_Click(sender As Object, e As EventArgs) Handles btnLoadDeliveryRef.Click
        'LOAD DELIVERY REF:

    End Sub

    Private Sub btnCreateTimesheet_Click(sender As Object, e As EventArgs) Handles btnCreateTimesheet.Click
        'CREATE TIMESHEET:

    End Sub

    Private Sub btnImportData_Click(sender As Object, e As EventArgs) Handles btnImportData.Click
        'IMPORT DATA:

    End Sub

    Private Sub btnUpdateEmployees_Click(sender As Object, e As EventArgs) Handles btnUpdateEmployees.Click
        'UPDATE EMPLOYEES:

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'EXIT:

    End Sub

    Private Sub imgCalendar_SelectDeliveryDate_Click(sender As Object, e As EventArgs) Handles imgCalendar_SelectDeliveryDate.Click
        'CALENDER HAS BEEN SELECTED:
        'Bring up calendar control to allow user to select date:
        Dim Monthlycal = New frmMonthly

        For Each frm As Form In Application.OpenForms
            If frm Is frmMonthly Then
                frm.Focus()
                Return
            End If
        Next
        Monthlycal.MdiParent = frmMainGIForm 'frmGI_RP_Userform is NOT an MDI Container.
        Monthlycal.StartPosition = FormStartPosition.CenterParent
        Monthlycal.FormTitle = "alt:Select Delivery Date"
        Monthlycal.Name = "MonthlyCalView"
        Monthlycal.Show()
    End Sub

    Private Sub comDeliveryRef_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comDeliveryRef.SelectedIndexChanged
        'Selected Delivery REference:
        Dim SelectedRef As String
        Dim strDeliveryDate As String

        strDeliveryDate = Me.txtSelectDeliveryDate.Text
        If Len(strDeliveryDate) = 0 Then
            'MsgBox("No Delivery Date")
            frmMainGIForm.txtMessages.Text = "No Delivery Date"
            Exit Sub
        End If
        SelectedRef = Me.comDeliveryRef.Text
        If Len(SelectedRef) = 0 Then
            'MsgBox("No Delivery Ref")
            frmMainGIForm.txtMessages.Text = "No Delivery Ref"
            Exit Sub
        End If
        Call PopulateStaticControls(strDeliveryDate, SelectedRef, "")
        'MsgBox("Finished Populating from Ref = " & SelectedRef)
        frmMainGIForm.txtMessages.Text = "OK. Finished Populating from Ref = " & SelectedRef
    End Sub

    Private Sub dgvOperatives_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOperatives.CellContentClick

    End Sub

    Private Sub btnAddOperative_Click(sender As Object, e As EventArgs) Handles btnAddOperative.Click

    End Sub
End Class