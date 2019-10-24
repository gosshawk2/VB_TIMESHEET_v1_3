<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewExcelSheet
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
        Me.pnlTabs = New System.Windows.Forms.Panel()
        Me.tcExcelTabs = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.btnFinish = New System.Windows.Forms.Button()
        Me.txtSelectedTab = New System.Windows.Forms.TextBox()
        Me.lblSelectedTab = New System.Windows.Forms.Label()
        Me.txtFilename = New System.Windows.Forms.TextBox()
        Me.lstExcelTabs = New System.Windows.Forms.ListBox()
        Me.txtTITLE = New System.Windows.Forms.TextBox()
        Me.btnGetExcelSheet = New System.Windows.Forms.Button()
        Me.txtMessages = New System.Windows.Forms.TextBox()
        Me.dgvExcelSheet = New System.Windows.Forms.DataGridView()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.dgvExcelSheet2 = New System.Windows.Forms.DataGridView()
        Me.pnlTabs.SuspendLayout()
        Me.tcExcelTabs.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        CType(Me.dgvExcelSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvExcelSheet2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTabs
        '
        Me.pnlTabs.BackColor = System.Drawing.Color.Transparent
        Me.pnlTabs.Controls.Add(Me.tcExcelTabs)
        Me.pnlTabs.Location = New System.Drawing.Point(0, 183)
        Me.pnlTabs.Name = "pnlTabs"
        Me.pnlTabs.Size = New System.Drawing.Size(970, 367)
        Me.pnlTabs.TabIndex = 0
        '
        'tcExcelTabs
        '
        Me.tcExcelTabs.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.tcExcelTabs.Controls.Add(Me.TabPage1)
        Me.tcExcelTabs.Controls.Add(Me.TabPage2)
        Me.tcExcelTabs.Location = New System.Drawing.Point(3, 27)
        Me.tcExcelTabs.Name = "tcExcelTabs"
        Me.tcExcelTabs.SelectedIndex = 0
        Me.tcExcelTabs.Size = New System.Drawing.Size(964, 333)
        Me.tcExcelTabs.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Lavender
        Me.TabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TabPage1.Controls.Add(Me.dgvExcelSheet)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(956, 307)
        Me.TabPage1.TabIndex = 1
        Me.TabPage1.Text = "Excel Tab 1"
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.Color.AliceBlue
        Me.pnlTop.Controls.Add(Me.btnFinish)
        Me.pnlTop.Controls.Add(Me.txtSelectedTab)
        Me.pnlTop.Controls.Add(Me.lblSelectedTab)
        Me.pnlTop.Controls.Add(Me.txtFilename)
        Me.pnlTop.Controls.Add(Me.lstExcelTabs)
        Me.pnlTop.Controls.Add(Me.txtTITLE)
        Me.pnlTop.Controls.Add(Me.btnGetExcelSheet)
        Me.pnlTop.Location = New System.Drawing.Point(16, 5)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(954, 152)
        Me.pnlTop.TabIndex = 1
        '
        'btnFinish
        '
        Me.btnFinish.BackColor = System.Drawing.Color.Red
        Me.btnFinish.Font = New System.Drawing.Font("Cambria", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinish.Location = New System.Drawing.Point(775, 70)
        Me.btnFinish.Name = "btnFinish"
        Me.btnFinish.Size = New System.Drawing.Size(164, 37)
        Me.btnFinish.TabIndex = 6
        Me.btnFinish.Text = "Finish"
        Me.btnFinish.UseVisualStyleBackColor = False
        Me.btnFinish.Visible = False
        '
        'txtSelectedTab
        '
        Me.txtSelectedTab.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectedTab.Location = New System.Drawing.Point(438, 120)
        Me.txtSelectedTab.Name = "txtSelectedTab"
        Me.txtSelectedTab.Size = New System.Drawing.Size(501, 23)
        Me.txtSelectedTab.TabIndex = 5
        Me.txtSelectedTab.Visible = False
        '
        'lblSelectedTab
        '
        Me.lblSelectedTab.AutoSize = True
        Me.lblSelectedTab.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblSelectedTab.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectedTab.ForeColor = System.Drawing.Color.Black
        Me.lblSelectedTab.Location = New System.Drawing.Point(319, 123)
        Me.lblSelectedTab.Name = "lblSelectedTab"
        Me.lblSelectedTab.Size = New System.Drawing.Size(100, 17)
        Me.lblSelectedTab.TabIndex = 4
        Me.lblSelectedTab.Text = "Selected Tab = "
        Me.lblSelectedTab.Visible = False
        '
        'txtFilename
        '
        Me.txtFilename.BackColor = System.Drawing.Color.AliceBlue
        Me.txtFilename.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilename.Location = New System.Drawing.Point(3, 38)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(936, 26)
        Me.txtFilename.TabIndex = 3
        Me.txtFilename.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lstExcelTabs
        '
        Me.lstExcelTabs.BackColor = System.Drawing.Color.White
        Me.lstExcelTabs.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstExcelTabs.ForeColor = System.Drawing.Color.Black
        Me.lstExcelTabs.FormattingEnabled = True
        Me.lstExcelTabs.ItemHeight = 15
        Me.lstExcelTabs.Location = New System.Drawing.Point(70, 70)
        Me.lstExcelTabs.Name = "lstExcelTabs"
        Me.lstExcelTabs.Size = New System.Drawing.Size(229, 79)
        Me.lstExcelTabs.TabIndex = 2
        Me.lstExcelTabs.Visible = False
        '
        'txtTITLE
        '
        Me.txtTITLE.BackColor = System.Drawing.Color.SkyBlue
        Me.txtTITLE.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTITLE.Location = New System.Drawing.Point(3, 7)
        Me.txtTITLE.Name = "txtTITLE"
        Me.txtTITLE.Size = New System.Drawing.Size(936, 26)
        Me.txtTITLE.TabIndex = 1
        Me.txtTITLE.Text = "SELECT AND VIEW EXCEL SHEETS"
        Me.txtTITLE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnGetExcelSheet
        '
        Me.btnGetExcelSheet.BackColor = System.Drawing.Color.Gold
        Me.btnGetExcelSheet.Font = New System.Drawing.Font("Cambria", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetExcelSheet.Location = New System.Drawing.Point(322, 70)
        Me.btnGetExcelSheet.Name = "btnGetExcelSheet"
        Me.btnGetExcelSheet.Size = New System.Drawing.Size(285, 37)
        Me.btnGetExcelSheet.TabIndex = 0
        Me.btnGetExcelSheet.Text = "Select EXCEL Workbook"
        Me.btnGetExcelSheet.UseVisualStyleBackColor = False
        '
        'txtMessages
        '
        Me.txtMessages.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMessages.Location = New System.Drawing.Point(16, 160)
        Me.txtMessages.Name = "txtMessages"
        Me.txtMessages.Size = New System.Drawing.Size(954, 23)
        Me.txtMessages.TabIndex = 6
        Me.txtMessages.Visible = False
        '
        'dgvExcelSheet
        '
        Me.dgvExcelSheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExcelSheet.Location = New System.Drawing.Point(6, 6)
        Me.dgvExcelSheet.Name = "dgvExcelSheet"
        Me.dgvExcelSheet.Size = New System.Drawing.Size(940, 291)
        Me.dgvExcelSheet.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvExcelSheet2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(956, 307)
        Me.TabPage2.TabIndex = 2
        Me.TabPage2.Text = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvExcelSheet2
        '
        Me.dgvExcelSheet2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExcelSheet2.Location = New System.Drawing.Point(8, 8)
        Me.dgvExcelSheet2.Name = "dgvExcelSheet2"
        Me.dgvExcelSheet2.Size = New System.Drawing.Size(940, 291)
        Me.dgvExcelSheet2.TabIndex = 1
        '
        'frmViewExcelSheet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 556)
        Me.Controls.Add(Me.txtMessages)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlTabs)
        Me.Name = "frmViewExcelSheet"
        Me.Text = "View EXCEL Sheet"
        Me.pnlTabs.ResumeLayout(False)
        Me.tcExcelTabs.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        CType(Me.dgvExcelSheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvExcelSheet2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pnlTabs As Panel
    Friend WithEvents pnlTop As Panel
    Friend WithEvents txtTITLE As TextBox
    Friend WithEvents btnGetExcelSheet As Button
    Friend WithEvents tcExcelTabs As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents lstExcelTabs As ListBox
    Friend WithEvents btnFinish As Button
    Friend WithEvents txtSelectedTab As TextBox
    Friend WithEvents lblSelectedTab As Label
    Friend WithEvents txtFilename As TextBox
    Friend WithEvents txtMessages As TextBox
    Friend WithEvents dgvExcelSheet As DataGridView
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents dgvExcelSheet2 As DataGridView
End Class
