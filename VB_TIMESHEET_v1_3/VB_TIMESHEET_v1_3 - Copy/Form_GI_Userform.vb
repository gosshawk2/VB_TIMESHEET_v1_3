Public Class frmGI_Userform
    Private Sub ScrollableControl1_Click(sender As Object, e As EventArgs) Handles Frame_Operatives.Click
        Me.HorizontalScroll.Minimum = 50
        Me.HorizontalScroll.Maximum = 100
        Me.VerticalScroll.Minimum = 50
        Me.VerticalScroll.Maximum = 100
        Me.AutoScroll = False
        Me.HScroll = True
        Me.VScroll = True
        Me.VerticalScroll.LargeChange = 5
        Me.HorizontalScroll.LargeChange = 5
        Me.VerticalScroll.Value = 50
        Me.HorizontalScroll.Value = 100

    End Sub

    Private Sub frmGI_Userform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '*************************************************************** LOAD PROCEDURE ********************************************************************
        Me.Frame_Operatives.HorizontalScroll.Minimum = 50
        Me.Frame_Operatives.HorizontalScroll.Maximum = 200
        Me.Frame_Operatives.VerticalScroll.Minimum = 50
        Me.Frame_Operatives.VerticalScroll.Maximum = 200
        Me.Frame_Operatives.AutoScroll = True
        Me.Frame_Operatives.Controls.Clear()
        Me.Left = 10
        Me.Top = 100
    End Sub

    Private Sub btnAddInboundData_Click(sender As Object, e As EventArgs)
        'Create INBOUND DATA:

    End Sub

    Private Sub btnResetInboundData_Click(sender As Object, e As EventArgs)
        'RESET INBOUND DATA:

    End Sub

    Private Sub btnAddOperative_Click(sender As Object, e As EventArgs) Handles btnAddOperative.Click
        'ADD OPERATIVE:

    End Sub

    Private Sub btnAddShort_Click(sender As Object, e As EventArgs) Handles btnAddShort.Click
        'ADD SHORT:

    End Sub

    Private Sub btnAddExtra_Click(sender As Object, e As EventArgs) Handles btnAddExtra.Click
        'ADD EXTRA:

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles txtPalletsArrived.TextChanged

    End Sub

    Private Sub imgCalendar_DeliveryDate_Click(sender As Object, e As EventArgs) Handles imgCalendar_ImportDate.Click
        'SELECT DELIVERY DATE:

    End Sub

    Private Sub btnSelectDatabase_Click(sender As Object, e As EventArgs)
        'SELECT DATABASE:

    End Sub

    Private Sub btnImportData_Click(sender As Object, e As EventArgs) Handles btnImportData.Click
        'IMPORT DATA:
        'IMPORT ALL OF CHETANs SPREADSHEET TO SEPARATE TABLE ALSO - FOR VIEWING IN SEP FORM USING VIEW GRID:
        Dim ExtractArr(,) As Object
        Dim Sheetname As String
        Dim EndRow As Long
        Dim EndCol As Long
        Dim ExcelFilename As String
        Dim TotalRows As Long
        Dim RowIDX As Long


        ExcelFilename = BrowseFilename(1)
        If Len(ExcelFilename) > 0 Then
            EndRow = 0
            EndCol = 48
            Sheetname = "Daily"
            ExtractArr = ExtractExcelRows(ExcelFilename, Sheetname, 1, 1, EndRow, EndCol)



        End If


    End Sub

    Private Sub btnUpdateEmployees_Click(sender As Object, e As EventArgs) Handles btnUpdateEmployees.Click
        'UPDATE EMPLOYEE TABLE:

    End Sub

    Private Sub rbASNNo_Click(sender As Object, e As EventArgs) Handles rbASNNo.Click
        comASNNo.Visible = True
        comDeliveryRef.Visible = False
        lblSelectRef.Text = "Select ASN"
    End Sub

    Private Sub rbDeliveryRef_Click(sender As Object, e As EventArgs) Handles rbDeliveryRef.Click
        comASNNo.Visible = False
        comDeliveryRef.Visible = True
        lblSelectRef.Text = "Select Delivery Ref"
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

    End Sub
End Class