Public Class frmPreferences
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub

    Private Sub btnSAVE_Click(sender As Object, e As EventArgs) Handles btnSAVE.Click
        'SAVE PREFERENCES:
        Dim EmpNo As String

        EmpNo = frmMainGIForm.myEmpNo
        Call Save_Preference(EmpNo)


    End Sub

    Sub Save_Preference(EmpNo As String)
        'SAVE PREFERENCE to tblPreferences
        'Need to check if user has account already ?
        'USE EMPLOYEE NO.
        Dim PrefTABLE As String = "tblPreferences"
        Dim ExcludeFields As String = "ID"
        Dim SearchCriteria As String = ""
        Dim SearchField As String
        Dim SearchValue As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim SearchType As String = "STRING"
        Dim AllFieldValues() As Object
        Dim AllFieldnames() As String
        Dim FoundRecord As Boolean = False
        Dim FieldNames As String
        Dim FieldValues As String
        Dim Username As String
        Dim Firstname As String
        Dim Lastname As String
        Dim LastSaved As DateTime
        Dim DeliveryGridFields As String
        Dim StartForm As String
        Dim UpdatedOK As Boolean = False
        Dim UpdateCriteria As String = ""
        Dim ErrMessage As String = ""

        ExcludeFields = "ID"
        SearchCriteria = ""
        SearchValue = EmpNo
        SearchField = "EmpNo"
        ReturnField = "ID"
        'ONLY WORKS FOR SINGLE RECORDS:
        'test for IsUpdate: Search Database first for Delivery Date and Reference:
        FoundRecord = Find_myQuery(frmMainGIForm.myConnString, PrefTABLE, SearchField, SearchValue, "STRING",
                                                      ReturnField, ReturnValue, AllFieldValues, AllFieldnames, SearchCriteria)

        StartForm = Me.txtStartForm.Text
        If FoundRecord Then
            FieldNames = "StartForm"
            FieldValues = StartForm
            UpdateCriteria = "ID = " & ReturnValue
        Else
            'EMP NO NOT FOUND:
            FieldNames = "Username,EmpNo,Firstname,Lastname,LastSaved,StartForm"
            Username = frmMainGIForm.myUsername
            Firstname = frmMainGIForm.myFirstname
            Lastname = frmMainGIForm.myLastname
            LastSaved = Now()

            FieldValues = Username
            FieldValues = FieldValues & "," & EmpNo
            FieldValues = FieldValues & "," & Firstname
            FieldValues = FieldValues & "," & Lastname
            FieldValues = FieldValues & "," & LastSaved
            FieldValues = FieldValues & "," & StartForm

        End If
        UpdatedOK = InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, FoundRecord, "", PrefTABLE,
                                                                FieldNames, FieldValues, UpdateCriteria, ExcludeFields, ErrMessage, False)
        If UpdatedOK Then
            MsgBox("OK START FORM CHANGED")
        Else
            MsgBox("Record did NOT save")
        End If


    End Sub

    Private Sub rbTimesheetEntry_CheckedChanged(sender As Object, e As EventArgs) Handles rbTimesheetEntry.CheckedChanged
        Me.txtStartForm.Text = "Timesheet Entry"
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)
        Me.txtStartForm.Text = "Status View"
    End Sub

    Private Sub rbTimesheetEntry_Click(sender As Object, e As EventArgs) Handles rbTimesheetEntry.Click
        Me.txtStartForm.Text = "Timesheet Entry"
    End Sub

    Private Sub rbStatusView_CheckedChanged(sender As Object, e As EventArgs) Handles rbStatusView.CheckedChanged
        Me.txtStartForm.Text = "Status View"
    End Sub

    Private Sub rbStatusView_Click(sender As Object, e As EventArgs) Handles rbStatusView.Click
        Me.txtStartForm.Text = "Status View"
    End Sub
End Class