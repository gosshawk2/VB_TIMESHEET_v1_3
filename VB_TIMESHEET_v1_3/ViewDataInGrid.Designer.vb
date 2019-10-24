<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewDataInGrid
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewDataInGrid))
        Me.pnlViewGrid = New System.Windows.Forms.Panel()
        Me.dgv = New System.Windows.Forms.DataGridView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnExportToCSV = New System.Windows.Forms.Button()
        Me.btnViewDBTable = New System.Windows.Forms.Button()
        Me.comDBTables = New System.Windows.Forms.ComboBox()
        Me.lblTotalRecords = New System.Windows.Forms.Label()
        Me.txtTotalRecs = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlSearch = New System.Windows.Forms.Panel()
        Me.lblSearchRef = New System.Windows.Forms.Label()
        Me.txtSearchRef = New System.Windows.Forms.TextBox()
        Me.lblSearchASN = New System.Windows.Forms.Label()
        Me.txtSearchASN = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.lblDBTables = New System.Windows.Forms.Label()
        Me.lblDateFrom = New System.Windows.Forms.Label()
        Me.txtDateFrom = New System.Windows.Forms.TextBox()
        Me.lblDateTo = New System.Windows.Forms.Label()
        Me.txtDateTo = New System.Windows.Forms.TextBox()
        Me.btnColSave = New System.Windows.Forms.Button()
        Me.pnlViewGrid.SuspendLayout()
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlSearch.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlViewGrid
        '
        Me.pnlViewGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlViewGrid.Controls.Add(Me.dgv)
        Me.pnlViewGrid.Location = New System.Drawing.Point(56, 89)
        Me.pnlViewGrid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnlViewGrid.Name = "pnlViewGrid"
        Me.pnlViewGrid.Size = New System.Drawing.Size(1034, 386)
        Me.pnlViewGrid.TabIndex = 3
        '
        'dgv
        '
        Me.dgv.AllowUserToOrderColumns = True
        Me.dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgv.BackgroundColor = System.Drawing.Color.LightBlue
        Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv.Location = New System.Drawing.Point(3, 6)
        Me.dgv.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgv.Name = "dgv"
        Me.dgv.Size = New System.Drawing.Size(1028, 374)
        Me.dgv.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel1.Controls.Add(Me.lblDBTables)
        Me.Panel1.Controls.Add(Me.btnExportToCSV)
        Me.Panel1.Controls.Add(Me.btnViewDBTable)
        Me.Panel1.Controls.Add(Me.comDBTables)
        Me.Panel1.Controls.Add(Me.lblTotalRecords)
        Me.Panel1.Controls.Add(Me.txtTotalRecs)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Location = New System.Drawing.Point(56, 4)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1034, 43)
        Me.Panel1.TabIndex = 2
        '
        'btnExportToCSV
        '
        Me.btnExportToCSV.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnExportToCSV.BackgroundImage = CType(resources.GetObject("btnExportToCSV.BackgroundImage"), System.Drawing.Image)
        Me.btnExportToCSV.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnExportToCSV.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExportToCSV.Location = New System.Drawing.Point(707, 2)
        Me.btnExportToCSV.Name = "btnExportToCSV"
        Me.btnExportToCSV.Size = New System.Drawing.Size(161, 38)
        Me.btnExportToCSV.TabIndex = 6
        Me.btnExportToCSV.UseVisualStyleBackColor = False
        '
        'btnViewDBTable
        '
        Me.btnViewDBTable.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnViewDBTable.BackgroundImage = CType(resources.GetObject("btnViewDBTable.BackgroundImage"), System.Drawing.Image)
        Me.btnViewDBTable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnViewDBTable.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnViewDBTable.Location = New System.Drawing.Point(588, 2)
        Me.btnViewDBTable.Name = "btnViewDBTable"
        Me.btnViewDBTable.Size = New System.Drawing.Size(96, 38)
        Me.btnViewDBTable.TabIndex = 4
        Me.btnViewDBTable.UseVisualStyleBackColor = False
        '
        'comDBTables
        '
        Me.comDBTables.BackColor = System.Drawing.Color.AliceBlue
        Me.comDBTables.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comDBTables.FormattingEnabled = True
        Me.comDBTables.Location = New System.Drawing.Point(384, 10)
        Me.comDBTables.Name = "comDBTables"
        Me.comDBTables.Size = New System.Drawing.Size(170, 23)
        Me.comDBTables.TabIndex = 3
        '
        'lblTotalRecords
        '
        Me.lblTotalRecords.AutoSize = True
        Me.lblTotalRecords.BackColor = System.Drawing.Color.White
        Me.lblTotalRecords.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalRecords.ForeColor = System.Drawing.Color.Blue
        Me.lblTotalRecords.Location = New System.Drawing.Point(11, 14)
        Me.lblTotalRecords.Name = "lblTotalRecords"
        Me.lblTotalRecords.Size = New System.Drawing.Size(86, 15)
        Me.lblTotalRecords.TabIndex = 2
        Me.lblTotalRecords.Text = "Total Records:"
        '
        'txtTotalRecs
        '
        Me.txtTotalRecs.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalRecs.Location = New System.Drawing.Point(108, 11)
        Me.txtTotalRecs.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtTotalRecs.Name = "txtTotalRecs"
        Me.txtTotalRecs.Size = New System.Drawing.Size(64, 23)
        Me.txtTotalRecs.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.Red
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnClose.Location = New System.Drawing.Point(922, 2)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(96, 38)
        Me.btnClose.TabIndex = 0
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'pnlSearch
        '
        Me.pnlSearch.BackColor = System.Drawing.Color.DodgerBlue
        Me.pnlSearch.Controls.Add(Me.btnColSave)
        Me.pnlSearch.Controls.Add(Me.lblDateTo)
        Me.pnlSearch.Controls.Add(Me.txtDateTo)
        Me.pnlSearch.Controls.Add(Me.lblDateFrom)
        Me.pnlSearch.Controls.Add(Me.txtDateFrom)
        Me.pnlSearch.Controls.Add(Me.btnSearch)
        Me.pnlSearch.Controls.Add(Me.lblSearchRef)
        Me.pnlSearch.Controls.Add(Me.txtSearchRef)
        Me.pnlSearch.Controls.Add(Me.lblSearchASN)
        Me.pnlSearch.Controls.Add(Me.txtSearchASN)
        Me.pnlSearch.Location = New System.Drawing.Point(56, 48)
        Me.pnlSearch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnlSearch.Name = "pnlSearch"
        Me.pnlSearch.Size = New System.Drawing.Size(1034, 43)
        Me.pnlSearch.TabIndex = 8
        '
        'lblSearchRef
        '
        Me.lblSearchRef.AutoSize = True
        Me.lblSearchRef.BackColor = System.Drawing.Color.White
        Me.lblSearchRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchRef.ForeColor = System.Drawing.Color.Blue
        Me.lblSearchRef.Location = New System.Drawing.Point(309, 14)
        Me.lblSearchRef.Name = "lblSearchRef"
        Me.lblSearchRef.Size = New System.Drawing.Size(67, 15)
        Me.lblSearchRef.TabIndex = 7
        Me.lblSearchRef.Text = "Search Ref:"
        '
        'txtSearchRef
        '
        Me.txtSearchRef.BackColor = System.Drawing.Color.Azure
        Me.txtSearchRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchRef.ForeColor = System.Drawing.Color.Black
        Me.txtSearchRef.Location = New System.Drawing.Point(384, 10)
        Me.txtSearchRef.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtSearchRef.Name = "txtSearchRef"
        Me.txtSearchRef.Size = New System.Drawing.Size(100, 23)
        Me.txtSearchRef.TabIndex = 5
        '
        'lblSearchASN
        '
        Me.lblSearchASN.AutoSize = True
        Me.lblSearchASN.BackColor = System.Drawing.Color.White
        Me.lblSearchASN.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchASN.ForeColor = System.Drawing.Color.Blue
        Me.lblSearchASN.Location = New System.Drawing.Point(25, 14)
        Me.lblSearchASN.Name = "lblSearchASN"
        Me.lblSearchASN.Size = New System.Drawing.Size(72, 15)
        Me.lblSearchASN.TabIndex = 2
        Me.lblSearchASN.Text = "Search ASN:"
        '
        'txtSearchASN
        '
        Me.txtSearchASN.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchASN.Location = New System.Drawing.Point(108, 11)
        Me.txtSearchASN.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtSearchASN.Name = "txtSearchASN"
        Me.txtSearchASN.Size = New System.Drawing.Size(100, 23)
        Me.txtSearchASN.TabIndex = 1
        '
        'btnSearch
        '
        Me.btnSearch.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnSearch.BackgroundImage = CType(resources.GetObject("btnSearch.BackgroundImage"), System.Drawing.Image)
        Me.btnSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSearch.Location = New System.Drawing.Point(923, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(96, 38)
        Me.btnSearch.TabIndex = 8
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'lblDBTables
        '
        Me.lblDBTables.AutoSize = True
        Me.lblDBTables.BackColor = System.Drawing.Color.White
        Me.lblDBTables.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBTables.ForeColor = System.Drawing.Color.Blue
        Me.lblDBTables.Location = New System.Drawing.Point(309, 13)
        Me.lblDBTables.Name = "lblDBTables"
        Me.lblDBTables.Size = New System.Drawing.Size(66, 15)
        Me.lblDBTables.TabIndex = 8
        Me.lblDBTables.Text = "DB Tables:"
        '
        'lblDateFrom
        '
        Me.lblDateFrom.AutoSize = True
        Me.lblDateFrom.BackColor = System.Drawing.Color.White
        Me.lblDateFrom.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateFrom.ForeColor = System.Drawing.Color.Blue
        Me.lblDateFrom.Location = New System.Drawing.Point(512, 14)
        Me.lblDateFrom.Name = "lblDateFrom"
        Me.lblDateFrom.Size = New System.Drawing.Size(68, 15)
        Me.lblDateFrom.TabIndex = 10
        Me.lblDateFrom.Text = "Date From:"
        '
        'txtDateFrom
        '
        Me.txtDateFrom.BackColor = System.Drawing.Color.Azure
        Me.txtDateFrom.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateFrom.ForeColor = System.Drawing.Color.Black
        Me.txtDateFrom.Location = New System.Drawing.Point(586, 10)
        Me.txtDateFrom.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtDateFrom.Name = "txtDateFrom"
        Me.txtDateFrom.Size = New System.Drawing.Size(100, 26)
        Me.txtDateFrom.TabIndex = 9
        '
        'lblDateTo
        '
        Me.lblDateTo.AutoSize = True
        Me.lblDateTo.BackColor = System.Drawing.Color.White
        Me.lblDateTo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateTo.ForeColor = System.Drawing.Color.Blue
        Me.lblDateTo.Location = New System.Drawing.Point(706, 14)
        Me.lblDateTo.Name = "lblDateTo"
        Me.lblDateTo.Size = New System.Drawing.Size(24, 15)
        Me.lblDateTo.TabIndex = 12
        Me.lblDateTo.Text = "To:"
        '
        'txtDateTo
        '
        Me.txtDateTo.BackColor = System.Drawing.Color.Azure
        Me.txtDateTo.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateTo.ForeColor = System.Drawing.Color.Black
        Me.txtDateTo.Location = New System.Drawing.Point(736, 10)
        Me.txtDateTo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtDateTo.Name = "txtDateTo"
        Me.txtDateTo.Size = New System.Drawing.Size(100, 26)
        Me.txtDateTo.TabIndex = 11
        '
        'btnColSave
        '
        Me.btnColSave.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnColSave.BackgroundImage = CType(resources.GetObject("btnColSave.BackgroundImage"), System.Drawing.Image)
        Me.btnColSave.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnColSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnColSave.Location = New System.Drawing.Point(865, 3)
        Me.btnColSave.Name = "btnColSave"
        Me.btnColSave.Size = New System.Drawing.Size(31, 38)
        Me.btnColSave.TabIndex = 13
        Me.btnColSave.Text = "&s"
        Me.btnColSave.UseVisualStyleBackColor = False
        '
        'frmViewDataInGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1117, 480)
        Me.Controls.Add(Me.pnlSearch)
        Me.Controls.Add(Me.pnlViewGrid)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "frmViewDataInGrid"
        Me.Text = "View Data from Database"
        Me.pnlViewGrid.ResumeLayout(False)
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnlSearch.ResumeLayout(False)
        Me.pnlSearch.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlViewGrid As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lblTotalRecords As Label
    Friend WithEvents txtTotalRecs As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents dgv As DataGridView
    Friend WithEvents btnViewDBTable As Button
    Friend WithEvents comDBTables As ComboBox
    Friend WithEvents btnExportToCSV As Button
    Friend WithEvents lblDBTables As Label
    Friend WithEvents pnlSearch As Panel
    Friend WithEvents btnSearch As Button
    Friend WithEvents lblSearchRef As Label
    Friend WithEvents txtSearchRef As TextBox
    Friend WithEvents lblSearchASN As Label
    Friend WithEvents txtSearchASN As TextBox
    Friend WithEvents lblDateTo As Label
    Friend WithEvents txtDateTo As TextBox
    Friend WithEvents lblDateFrom As Label
    Friend WithEvents txtDateFrom As TextBox
    Friend WithEvents btnColSave As Button
End Class
