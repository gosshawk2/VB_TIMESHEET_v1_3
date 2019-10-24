<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReferenceProgress
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReferenceProgress))
        Me.pnlTitle = New System.Windows.Forms.Panel()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.pnlbuttons = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSearchRef = New System.Windows.Forms.TextBox()
        Me.btnSearchEmpNo = New System.Windows.Forms.Button()
        Me.lblSearchASN = New System.Windows.Forms.Label()
        Me.txtSearchASN = New System.Windows.Forms.TextBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.dgvDeliveryProgress = New System.Windows.Forms.DataGridView()
        Me.pnlTitle.SuspendLayout()
        Me.pnlbuttons.SuspendLayout()
        CType(Me.dgvDeliveryProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTitle
        '
        Me.pnlTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTitle.BackColor = System.Drawing.Color.Blue
        Me.pnlTitle.Controls.Add(Me.txtTitle)
        Me.pnlTitle.ForeColor = System.Drawing.Color.Black
        Me.pnlTitle.Location = New System.Drawing.Point(3, 0)
        Me.pnlTitle.Name = "pnlTitle"
        Me.pnlTitle.Size = New System.Drawing.Size(995, 47)
        Me.pnlTitle.TabIndex = 41
        '
        'txtTitle
        '
        Me.txtTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTitle.BackColor = System.Drawing.Color.DarkBlue
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.AliceBlue
        Me.txtTitle.Location = New System.Drawing.Point(1, 5)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(990, 36)
        Me.txtTitle.TabIndex = 0
        Me.txtTitle.Text = "Delivery Status"
        Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlbuttons
        '
        Me.pnlbuttons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlbuttons.BackColor = System.Drawing.Color.MidnightBlue
        Me.pnlbuttons.Controls.Add(Me.Label1)
        Me.pnlbuttons.Controls.Add(Me.txtSearchRef)
        Me.pnlbuttons.Controls.Add(Me.btnSearchEmpNo)
        Me.pnlbuttons.Controls.Add(Me.lblSearchASN)
        Me.pnlbuttons.Controls.Add(Me.txtSearchASN)
        Me.pnlbuttons.Controls.Add(Me.btnSave)
        Me.pnlbuttons.Controls.Add(Me.btnClose)
        Me.pnlbuttons.Location = New System.Drawing.Point(3, 47)
        Me.pnlbuttons.Name = "pnlbuttons"
        Me.pnlbuttons.Size = New System.Drawing.Size(990, 50)
        Me.pnlbuttons.TabIndex = 42
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.Label1.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(445, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 19)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Search Reference:"
        '
        'txtSearchRef
        '
        Me.txtSearchRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchRef.Location = New System.Drawing.Point(595, 15)
        Me.txtSearchRef.Name = "txtSearchRef"
        Me.txtSearchRef.Size = New System.Drawing.Size(125, 23)
        Me.txtSearchRef.TabIndex = 8
        '
        'btnSearchEmpNo
        '
        Me.btnSearchEmpNo.BackColor = System.Drawing.Color.SpringGreen
        Me.btnSearchEmpNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchEmpNo.Location = New System.Drawing.Point(729, 10)
        Me.btnSearchEmpNo.Name = "btnSearchEmpNo"
        Me.btnSearchEmpNo.Size = New System.Drawing.Size(60, 30)
        Me.btnSearchEmpNo.TabIndex = 7
        Me.btnSearchEmpNo.Text = "Go"
        Me.btnSearchEmpNo.UseVisualStyleBackColor = False
        '
        'lblSearchASN
        '
        Me.lblSearchASN.AutoSize = True
        Me.lblSearchASN.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblSearchASN.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchASN.Location = New System.Drawing.Point(169, 15)
        Me.lblSearchASN.Name = "lblSearchASN"
        Me.lblSearchASN.Size = New System.Drawing.Size(125, 19)
        Me.lblSearchASN.TabIndex = 6
        Me.lblSearchASN.Text = "Search ASN/PO:"
        '
        'txtSearchASN
        '
        Me.txtSearchASN.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchASN.Location = New System.Drawing.Point(300, 15)
        Me.txtSearchASN.Name = "txtSearchASN"
        Me.txtSearchASN.Size = New System.Drawing.Size(125, 23)
        Me.txtSearchASN.TabIndex = 5
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.SpringGreen
        Me.btnSave.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(18, 10)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(140, 30)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        Me.btnSave.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.Red
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnClose.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(873, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(101, 44)
        Me.btnClose.TabIndex = 3
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'dgvDeliveryProgress
        '
        Me.dgvDeliveryProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDeliveryProgress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDeliveryProgress.Location = New System.Drawing.Point(4, 103)
        Me.dgvDeliveryProgress.Name = "dgvDeliveryProgress"
        Me.dgvDeliveryProgress.Size = New System.Drawing.Size(990, 371)
        Me.dgvDeliveryProgress.TabIndex = 43
        '
        'frmReferenceProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Blue
        Me.ClientSize = New System.Drawing.Size(1004, 483)
        Me.Controls.Add(Me.dgvDeliveryProgress)
        Me.Controls.Add(Me.pnlbuttons)
        Me.Controls.Add(Me.pnlTitle)
        Me.MinimumSize = New System.Drawing.Size(1020, 39)
        Me.Name = "frmReferenceProgress"
        Me.Text = "Reference Progress"
        Me.pnlTitle.ResumeLayout(False)
        Me.pnlTitle.PerformLayout()
        Me.pnlbuttons.ResumeLayout(False)
        Me.pnlbuttons.PerformLayout()
        CType(Me.dgvDeliveryProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTitle As Panel
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents pnlbuttons As Panel
    Friend WithEvents btnSearchEmpNo As Button
    Friend WithEvents lblSearchASN As Label
    Friend WithEvents txtSearchASN As TextBox
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtSearchRef As TextBox
    Friend WithEvents dgvDeliveryProgress As DataGridView
End Class
