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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.pnlSelection = New System.Windows.Forms.Panel()
        Me.txtSelectedSheet = New System.Windows.Forms.TextBox()
        Me.lstSheets = New System.Windows.Forms.ListBox()
        Me.btnSelectWorkbook = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtWorkbook = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlTop.SuspendLayout()
        Me.pnlSelection.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTop.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.pnlTop.Controls.Add(Me.TextBox1)
        Me.pnlTop.Location = New System.Drawing.Point(12, 12)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(652, 41)
        Me.pnlTop.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.AliceBlue
        Me.TextBox1.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(3, 7)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(646, 26)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "SELECT SHEET from WORKBBOOK"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlSelection
        '
        Me.pnlSelection.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlSelection.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.pnlSelection.Controls.Add(Me.txtSelectedSheet)
        Me.pnlSelection.Controls.Add(Me.lstSheets)
        Me.pnlSelection.Controls.Add(Me.btnSelectWorkbook)
        Me.pnlSelection.Controls.Add(Me.btnOK)
        Me.pnlSelection.Controls.Add(Me.btnCancel)
        Me.pnlSelection.Controls.Add(Me.txtWorkbook)
        Me.pnlSelection.Controls.Add(Me.Label2)
        Me.pnlSelection.Controls.Add(Me.Label1)
        Me.pnlSelection.Location = New System.Drawing.Point(12, 54)
        Me.pnlSelection.Name = "pnlSelection"
        Me.pnlSelection.Size = New System.Drawing.Size(652, 347)
        Me.pnlSelection.TabIndex = 1
        '
        'txtSelectedSheet
        '
        Me.txtSelectedSheet.BackColor = System.Drawing.Color.AliceBlue
        Me.txtSelectedSheet.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectedSheet.ForeColor = System.Drawing.Color.Black
        Me.txtSelectedSheet.Location = New System.Drawing.Point(170, 212)
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
        Me.lstSheets.Location = New System.Drawing.Point(170, 116)
        Me.lstSheets.Name = "lstSheets"
        Me.lstSheets.ScrollAlwaysVisible = True
        Me.lstSheets.Size = New System.Drawing.Size(280, 80)
        Me.lstSheets.TabIndex = 6
        '
        'btnSelectWorkbook
        '
        Me.btnSelectWorkbook.BackColor = System.Drawing.Color.Gold
        Me.btnSelectWorkbook.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectWorkbook.ForeColor = System.Drawing.Color.Black
        Me.btnSelectWorkbook.Location = New System.Drawing.Point(170, 15)
        Me.btnSelectWorkbook.Name = "btnSelectWorkbook"
        Me.btnSelectWorkbook.Size = New System.Drawing.Size(280, 30)
        Me.btnSelectWorkbook.TabIndex = 5
        Me.btnSelectWorkbook.Text = "Select EXCEL Workbook"
        Me.btnSelectWorkbook.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.Lime
        Me.btnOK.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.ForeColor = System.Drawing.Color.Black
        Me.btnOK.Location = New System.Drawing.Point(375, 252)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 53)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.Red
        Me.btnCancel.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Black
        Me.btnCancel.Location = New System.Drawing.Point(170, 252)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 53)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "CANCEL"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'txtWorkbook
        '
        Me.txtWorkbook.BackColor = System.Drawing.Color.AliceBlue
        Me.txtWorkbook.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkbook.ForeColor = System.Drawing.Color.Black
        Me.txtWorkbook.Location = New System.Drawing.Point(170, 70)
        Me.txtWorkbook.Name = "txtWorkbook"
        Me.txtWorkbook.Size = New System.Drawing.Size(479, 23)
        Me.txtWorkbook.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(10, 116)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 19)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Select Sheet:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(10, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(153, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Selected Workbook:"
        '
        'frmSelectSheet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MidnightBlue
        Me.ClientSize = New System.Drawing.Size(676, 413)
        Me.Controls.Add(Me.pnlSelection)
        Me.Controls.Add(Me.pnlTop)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectSheet"
        Me.Text = "Select Spreadsheet"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlSelection.ResumeLayout(False)
        Me.pnlSelection.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTop As Panel
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents pnlSelection As Panel
    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents txtWorkbook As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSelectWorkbook As Button
    Friend WithEvents lstSheets As ListBox
    Friend WithEvents txtSelectedSheet As TextBox
End Class
