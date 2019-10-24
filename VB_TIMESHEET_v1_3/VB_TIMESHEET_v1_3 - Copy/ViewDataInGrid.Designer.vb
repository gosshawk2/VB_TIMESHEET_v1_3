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
        Me.btnViewDBTable = New System.Windows.Forms.Button()
        Me.comDBTables = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtTotalRecs = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtTable = New System.Windows.Forms.TextBox()
        Me.pnlViewGrid.SuspendLayout()
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlViewGrid
        '
        Me.pnlViewGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlViewGrid.Controls.Add(Me.dgv)
        Me.pnlViewGrid.Location = New System.Drawing.Point(56, 70)
        Me.pnlViewGrid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnlViewGrid.Name = "pnlViewGrid"
        Me.pnlViewGrid.Size = New System.Drawing.Size(1034, 391)
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
        Me.dgv.Location = New System.Drawing.Point(3, 2)
        Me.dgv.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dgv.Name = "dgv"
        Me.dgv.Size = New System.Drawing.Size(1028, 386)
        Me.dgv.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel1.Controls.Add(Me.txtTable)
        Me.Panel1.Controls.Add(Me.btnViewDBTable)
        Me.Panel1.Controls.Add(Me.comDBTables)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtTotalRecs)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Location = New System.Drawing.Point(56, 20)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1034, 43)
        Me.Panel1.TabIndex = 2
        '
        'btnViewDBTable
        '
        Me.btnViewDBTable.BackColor = System.Drawing.Color.Lavender
        Me.btnViewDBTable.BackgroundImage = CType(resources.GetObject("btnViewDBTable.BackgroundImage"), System.Drawing.Image)
        Me.btnViewDBTable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnViewDBTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewDBTable.Location = New System.Drawing.Point(622, 2)
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
        Me.comDBTables.Location = New System.Drawing.Point(426, 10)
        Me.comDBTables.Name = "comDBTables"
        Me.comDBTables.Size = New System.Drawing.Size(170, 23)
        Me.comDBTables.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(107, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Total Records:"
        '
        'txtTotalRecs
        '
        Me.txtTotalRecs.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalRecs.Location = New System.Drawing.Point(204, 11)
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
        'txtTable
        '
        Me.txtTable.BackColor = System.Drawing.Color.Azure
        Me.txtTable.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTable.ForeColor = System.Drawing.Color.Black
        Me.txtTable.Location = New System.Drawing.Point(285, 11)
        Me.txtTable.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtTable.Name = "txtTable"
        Me.txtTable.Size = New System.Drawing.Size(117, 23)
        Me.txtTable.TabIndex = 5
        '
        'frmViewDataInGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1117, 480)
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
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlViewGrid As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents txtTotalRecs As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents dgv As DataGridView
    Friend WithEvents btnViewDBTable As Button
    Friend WithEvents comDBTables As ComboBox
    Friend WithEvents txtTable As TextBox
End Class
