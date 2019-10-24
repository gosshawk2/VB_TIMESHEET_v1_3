<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectSheet
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectSheet))
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.pnlSelection = New System.Windows.Forms.Panel()
        Me.lblSelectedSheet = New System.Windows.Forms.Label()
        Me.lblTotalUpdates = New System.Windows.Forms.Label()
        Me.txtTotalUpdates = New System.Windows.Forms.TextBox()
        Me.txtSelectedSheet = New System.Windows.Forms.TextBox()
        Me.lstSheets = New System.Windows.Forms.ListBox()
        Me.btnSelectWorkbook = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtWorkbook = New System.Windows.Forms.TextBox()
        Me.lblSelectSheet = New System.Windows.Forms.Label()
        Me.lblSelectedWorkbook = New System.Windows.Forms.Label()
        Me.txtCONFIRM_ASK_UPDATES = New System.Windows.Forms.TextBox()
        Me.lblConfirmUpdates = New System.Windows.Forms.Label()
        Me.pnlTop.SuspendLayout()
        Me.pnlSelection.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTop.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.pnlTop.Controls.Add(Me.txtTitle)
        Me.pnlTop.Location = New System.Drawing.Point(3, 3)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(577, 41)
        Me.pnlTop.TabIndex = 0
        '
        'txtTitle
        '
        Me.txtTitle.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.Black
        Me.txtTitle.Location = New System.Drawing.Point(3, 7)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(646, 26)
        Me.txtTitle.TabIndex = 0
        Me.txtTitle.Text = "SELECT SHEET from WORKBBOOK"
        Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlSelection
        '
        Me.pnlSelection.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlSelection.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.pnlSelection.Controls.Add(Me.lblConfirmUpdates)
        Me.pnlSelection.Controls.Add(Me.txtCONFIRM_ASK_UPDATES)
        Me.pnlSelection.Controls.Add(Me.lblSelectedSheet)
        Me.pnlSelection.Controls.Add(Me.lblTotalUpdates)
        Me.pnlSelection.Controls.Add(Me.txtTotalUpdates)
        Me.pnlSelection.Controls.Add(Me.txtSelectedSheet)
        Me.pnlSelection.Controls.Add(Me.lstSheets)
        Me.pnlSelection.Controls.Add(Me.btnSelectWorkbook)
        Me.pnlSelection.Controls.Add(Me.btnOK)
        Me.pnlSelection.Controls.Add(Me.btnCancel)
        Me.pnlSelection.Controls.Add(Me.txtWorkbook)
        Me.pnlSelection.Controls.Add(Me.lblSelectSheet)
        Me.pnlSelection.Controls.Add(Me.lblSelectedWorkbook)
        Me.pnlSelection.Location = New System.Drawing.Point(3, 45)
        Me.pnlSelection.Name = "pnlSelection"
        Me.pnlSelection.Size = New System.Drawing.Size(577, 404)
        Me.pnlSelection.TabIndex = 1
        '
        'lblSelectedSheet
        '
        Me.lblSelectedSheet.AutoSize = True
        Me.lblSelectedSheet.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectedSheet.ForeColor = System.Drawing.Color.White
        Me.lblSelectedSheet.Location = New System.Drawing.Point(10, 301)
        Me.lblSelectedSheet.Name = "lblSelectedSheet"
        Me.lblSelectedSheet.Size = New System.Drawing.Size(100, 19)
        Me.lblSelectedSheet.TabIndex = 10
        Me.lblSelectedSheet.Text = "Select Sheet:"
        Me.lblSelectedSheet.Visible = False
        '
        'lblTotalUpdates
        '
        Me.lblTotalUpdates.AutoSize = True
        Me.lblTotalUpdates.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalUpdates.ForeColor = System.Drawing.Color.White
        Me.lblTotalUpdates.Location = New System.Drawing.Point(10, 264)
        Me.lblTotalUpdates.Name = "lblTotalUpdates"
        Me.lblTotalUpdates.Size = New System.Drawing.Size(116, 19)
        Me.lblTotalUpdates.TabIndex = 9
        Me.lblTotalUpdates.Text = "Total Updates:"
        Me.lblTotalUpdates.Visible = False
        '
        'txtTotalUpdates
        '
        Me.txtTotalUpdates.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTotalUpdates.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalUpdates.ForeColor = System.Drawing.Color.Black
        Me.txtTotalUpdates.Location = New System.Drawing.Point(170, 264)
        Me.txtTotalUpdates.Name = "txtTotalUpdates"
        Me.txtTotalUpdates.Size = New System.Drawing.Size(280, 23)
        Me.txtTotalUpdates.TabIndex = 8
        Me.txtTotalUpdates.Visible = False
        '
        'txtSelectedSheet
        '
        Me.txtSelectedSheet.BackColor = System.Drawing.Color.AliceBlue
        Me.txtSelectedSheet.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectedSheet.ForeColor = System.Drawing.Color.Black
        Me.txtSelectedSheet.Location = New System.Drawing.Point(170, 301)
        Me.txtSelectedSheet.Name = "txtSelectedSheet"
        Me.txtSelectedSheet.Size = New System.Drawing.Size(280, 23)
        Me.txtSelectedSheet.TabIndex = 7
        '
        'lstSheets
        '
        Me.lstSheets.BackColor = System.Drawing.Color.AliceBlue
        Me.lstSheets.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstSheets.ForeColor = System.Drawing.Color.Black
        Me.lstSheets.FormattingEnabled = True
        Me.lstSheets.ItemHeight = 19
        Me.lstSheets.Location = New System.Drawing.Point(170, 171)
        Me.lstSheets.Name = "lstSheets"
        Me.lstSheets.ScrollAlwaysVisible = True
        Me.lstSheets.Size = New System.Drawing.Size(280, 80)
        Me.lstSheets.TabIndex = 6
        '
        'btnSelectWorkbook
        '
        Me.btnSelectWorkbook.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnSelectWorkbook.BackColor = System.Drawing.Color.Transparent
        Me.btnSelectWorkbook.BackgroundImage = CType(resources.GetObject("btnSelectWorkbook.BackgroundImage"), System.Drawing.Image)
        Me.btnSelectWorkbook.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSelectWorkbook.FlatAppearance.BorderColor = System.Drawing.Color.Blue
        Me.btnSelectWorkbook.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnSelectWorkbook.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnSelectWorkbook.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSelectWorkbook.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectWorkbook.ForeColor = System.Drawing.Color.Black
        Me.btnSelectWorkbook.Location = New System.Drawing.Point(126, 14)
        Me.btnSelectWorkbook.Name = "btnSelectWorkbook"
        Me.btnSelectWorkbook.Size = New System.Drawing.Size(380, 30)
        Me.btnSelectWorkbook.TabIndex = 5
        Me.btnSelectWorkbook.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.Lime
        Me.btnOK.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.Color.Black
        Me.btnOK.Location = New System.Drawing.Point(365, 341)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(85, 53)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "UPDATE"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.Red
        Me.btnCancel.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Black
        Me.btnCancel.Location = New System.Drawing.Point(170, 341)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(85, 53)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "CANCEL"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'txtWorkbook
        '
        Me.txtWorkbook.BackColor = System.Drawing.Color.AliceBlue
        Me.txtWorkbook.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkbook.ForeColor = System.Drawing.Color.Black
        Me.txtWorkbook.Location = New System.Drawing.Point(170, 125)
        Me.txtWorkbook.Name = "txtWorkbook"
        Me.txtWorkbook.Size = New System.Drawing.Size(280, 23)
        Me.txtWorkbook.TabIndex = 1
        '
        'lblSelectSheet
        '
        Me.lblSelectSheet.AutoSize = True
        Me.lblSelectSheet.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectSheet.ForeColor = System.Drawing.Color.White
        Me.lblSelectSheet.Location = New System.Drawing.Point(10, 171)
        Me.lblSelectSheet.Name = "lblSelectSheet"
        Me.lblSelectSheet.Size = New System.Drawing.Size(100, 19)
        Me.lblSelectSheet.TabIndex = 1
        Me.lblSelectSheet.Text = "Select Sheet:"
        '
        'lblSelectedWorkbook
        '
        Me.lblSelectedWorkbook.AutoSize = True
        Me.lblSelectedWorkbook.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectedWorkbook.ForeColor = System.Drawing.Color.White
        Me.lblSelectedWorkbook.Location = New System.Drawing.Point(10, 125)
        Me.lblSelectedWorkbook.Name = "lblSelectedWorkbook"
        Me.lblSelectedWorkbook.Size = New System.Drawing.Size(153, 19)
        Me.lblSelectedWorkbook.TabIndex = 0
        Me.lblSelectedWorkbook.Text = "Selected Workbook:"
        '
        'txtCONFIRM_ASK_UPDATES
        '
        Me.txtCONFIRM_ASK_UPDATES.BackColor = System.Drawing.Color.AliceBlue
        Me.txtCONFIRM_ASK_UPDATES.Font = New System.Drawing.Font("Cambria", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCONFIRM_ASK_UPDATES.Location = New System.Drawing.Point(261, 57)
        Me.txtCONFIRM_ASK_UPDATES.Name = "txtCONFIRM_ASK_UPDATES"
        Me.txtCONFIRM_ASK_UPDATES.Size = New System.Drawing.Size(27, 39)
        Me.txtCONFIRM_ASK_UPDATES.TabIndex = 11
        Me.txtCONFIRM_ASK_UPDATES.Text = "X"
        Me.txtCONFIRM_ASK_UPDATES.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblConfirmUpdates
        '
        Me.lblConfirmUpdates.AutoSize = True
        Me.lblConfirmUpdates.Font = New System.Drawing.Font("Cambria", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirmUpdates.ForeColor = System.Drawing.Color.LawnGreen
        Me.lblConfirmUpdates.Location = New System.Drawing.Point(9, 57)
        Me.lblConfirmUpdates.Name = "lblConfirmUpdates"
        Me.lblConfirmUpdates.Size = New System.Drawing.Size(246, 32)
        Me.lblConfirmUpdates.TabIndex = 12
        Me.lblConfirmUpdates.Text = "Confirm Updates ?:"
        '
        'frmSelectSheet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.MidnightBlue
        Me.ClientSize = New System.Drawing.Size(584, 461)
        Me.Controls.Add(Me.pnlSelection)
        Me.Controls.Add(Me.pnlTop)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(600, 500)
        Me.MinimumSize = New System.Drawing.Size(500, 400)
        Me.Name = "frmSelectSheet"
        Me.Text = "Select Spreadsheet"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlSelection.ResumeLayout(False)
        Me.pnlSelection.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTop As Panel
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents pnlSelection As Panel
    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents txtWorkbook As TextBox
    Friend WithEvents lblSelectSheet As Label
    Friend WithEvents lblSelectedWorkbook As Label
    Friend WithEvents btnSelectWorkbook As Button
    Friend WithEvents lstSheets As ListBox
    Friend WithEvents txtSelectedSheet As TextBox
    Friend WithEvents lblTotalUpdates As Label
    Friend WithEvents txtTotalUpdates As TextBox
    Friend WithEvents lblSelectedSheet As Label
    Friend WithEvents lblConfirmUpdates As Label
    Friend WithEvents txtCONFIRM_ASK_UPDATES As TextBox
End Class
