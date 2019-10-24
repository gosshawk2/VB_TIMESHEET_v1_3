Public Class frmInboundOfficeEntry
    Private Sub btnSubmitReference_Click(sender As Object, e As EventArgs) Handles btnSubmitReference.Click
        'CLERK Enters Reference Number:
        'Search For Reference (and latest date ?) and populate the other text boxes with returned info.
        Dim SearchText As String
        Dim SearchField As String
        Dim ReturnField As String
        Dim ReturnValue As String
        Dim SearchCriteria = ""
        Dim FieldType As String
        Dim Reversed As Boolean = True
        Dim SortFields As String
        Dim FoundRef As Boolean = False
        Dim AllValues() As Object
        Dim AllFields() As String
        Dim ErrMessage As String = ""
        Dim DBInfoTable As String = "tblDeliveryInfo"
        Dim strDeliveryRef As String
        Dim strDeliveryDate As String
        Dim dtDeliveryDate As DateTime
        Dim SupplierCode As String
        Dim SupplierName As String
        Dim DueTime As String


        strDeliveryRef = Me.txtDeliveryRefEntry.Text
        If Len(strDeliveryRef) = 0 Then
            MsgBox("Please enter a Delivery Reference first")
            Exit Sub
        End If

        SearchText = strDeliveryRef
        SearchField = "DeliveryReference"
        ReturnField = "ID"
        ReturnValue = ""
        FieldType = "STRING"
        Reversed = True
        SortFields = "DeliveryDate"
        FoundRef = Find_myQuery(frmMainGIForm.myConnString, DBInfoTable, SearchField, SearchText, FieldType, ReturnField, ReturnValue, AllValues,
                                   AllFields, SearchCriteria, SortFields, Reversed, ErrMessage)
        If FoundRef Then
            Me.pnlDeliveryInfo.Visible = True
            Me.pnlArrivalTime.Visible = True
            Me.txtDeliveryDate.Visible = True
            Me.txtSupplierCode.Visible = True
            Me.txtSupplierName.Visible = True
            Me.txtDueTime.Visible = True

            strDeliveryDate = GetMYValuebyFieldname(AllValues, AllFields, "DeliveryDate")
            SupplierCode = GetMYValuebyFieldname(AllValues, AllFields, "Supplier_Code")
            SupplierName = GetMYValuebyFieldname(AllValues, AllFields, "Supplier_Name")
            DueTime = GetMYValuebyFieldname(AllValues, AllFields, "Due_Time")
            Me.txtDeliveryDate.Text = CDate(strDeliveryDate).ToString("dd/MM/yyyy")
            Me.txtSupplierCode.Text = SupplierCode
            Me.txtSupplierName.Text = SupplierName
            Me.txtDueTime.Text = DueTime
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'CLOSE FORM:
        Me.Close()
    End Sub

    Private Sub btnSubmitTime_Click(sender As Object, e As EventArgs) Handles btnSubmitTime.Click
        'Accept Arrival Time and put into database - in tblDeliveryInfo.
        Dim strArrivalTime As String
        Dim FieldNames As String
        Dim FieldValues As String
        Dim SavedOK As Boolean = False
        Dim DBInfoTable As String = "tblDeliveryInfo"
        Dim UpdateCriteria As String
        Dim ExcludeFields As String
        Dim ErrMessages As String = ""
        Dim strDeliveryRef As String = ""

        strDeliveryRef = Me.txtDeliveryRefEntry.Text
        If Len(strDeliveryRef) = 0 Then
            MsgBox("Please enter a Delivery Reference first")
            Exit Sub
        End If
        strArrivalTime = Me.msktxtArrivalDate.Text & " " & Me.msktxtArrivalTime.Text
        If Len(strArrivalTime) = 0 Then
            MsgBox("No Arrival Time Entered")
            Exit Sub
        End If

        'Need to validate the Arrival Time (and date)
        'Needs a TIMESTAMP button to help the clerk. They press it. Auto-insert current date and time.
        'the submit button is used to confirm the date and time - as this may be different slightly at the time etc.

        FieldNames = "ArrivalTime"
        FieldValues = strArrivalTime
        ExcludeFields = "ID"
        'If JUST using the Delivery Reference as the update search criteria - WILL UPDATE ALL matching REFERENCES !
        UpdateCriteria = "DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34)
        SavedOK = Module_DanG_MySQL_Tools.InsertUpdateRecords_Using_Parameters(frmMainGIForm.myConnString, True, "",
                                                                               DBInfoTable, FieldNames, FieldValues,
                                                                               UpdateCriteria, ExcludeFields,
                                                                               ErrMessages, False)

        If SavedOK Then
            MsgBox("OK Arrival Time Entered")
        Else
            MsgBox("Problem: Arrival Time NOT SAVE")
        End If

    End Sub

    Private Sub btnSubmitGRNNumber_Click(sender As Object, e As EventArgs) Handles btnSubmitGRNNumber.Click
        'INSERT GRN NUMBER (when REFERENCE has been completed):

    End Sub
End Class