<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMysqlLogin
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMysqlLogin))
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnLocal = New System.Windows.Forms.Button()
        Me.btnRemote = New System.Windows.Forms.Button()
        Me.rb3307 = New System.Windows.Forms.RadioButton()
        Me.rb3306 = New System.Windows.Forms.RadioButton()
        Me.rb6080 = New System.Windows.Forms.RadioButton()
        Me.lblPort = New System.Windows.Forms.Label()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblDatabase = New System.Windows.Forms.Label()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnAdvanced = New System.Windows.Forms.Button()
        Me.btnLOGIN = New System.Windows.Forms.Button()
        Me.btnCANCEL = New System.Windows.Forms.Button()
        Me.btnAutoLOGIN = New System.Windows.Forms.Button()
        Me.pnlButtons.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlButtons
        '
        Me.pnlButtons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlButtons.BackColor = System.Drawing.Color.Blue
        Me.pnlButtons.Controls.Add(Me.btnClose)
        Me.pnlButtons.Controls.Add(Me.txtTitle)
        Me.pnlButtons.Location = New System.Drawing.Point(14, 7)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(686, 58)
        Me.pnlButtons.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.Red
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnClose.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(582, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(92, 36)
        Me.btnClose.TabIndex = 1
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'txtTitle
        '
        Me.txtTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.Location = New System.Drawing.Point(9, 6)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.ReadOnly = True
        Me.txtTitle.Size = New System.Drawing.Size(562, 36)
        Me.txtTitle.TabIndex = 0
        Me.txtTitle.TabStop = False
        Me.txtTitle.Text = "Please Login to Goods In Timesheet Database"
        Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel1.Controls.Add(Me.btnAutoLOGIN)
        Me.Panel1.Controls.Add(Me.btnLocal)
        Me.Panel1.Controls.Add(Me.btnRemote)
        Me.Panel1.Controls.Add(Me.rb3307)
        Me.Panel1.Controls.Add(Me.rb3306)
        Me.Panel1.Controls.Add(Me.rb6080)
        Me.Panel1.Controls.Add(Me.lblPort)
        Me.Panel1.Controls.Add(Me.txtPort)
        Me.Panel1.Controls.Add(Me.lblPassword)
        Me.Panel1.Controls.Add(Me.txtPassword)
        Me.Panel1.Controls.Add(Me.lblUsername)
        Me.Panel1.Controls.Add(Me.txtUsername)
        Me.Panel1.Controls.Add(Me.lblDatabase)
        Me.Panel1.Controls.Add(Me.txtDatabase)
        Me.Panel1.Controls.Add(Me.lblServer)
        Me.Panel1.Controls.Add(Me.txtServer)
        Me.Panel1.Location = New System.Drawing.Point(14, 73)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(686, 225)
        Me.Panel1.TabIndex = 1
        '
        'btnLocal
        '
        Me.btnLocal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLocal.BackColor = System.Drawing.Color.MidnightBlue
        Me.btnLocal.BackgroundImage = CType(resources.GetObject("btnLocal.BackgroundImage"), System.Drawing.Image)
        Me.btnLocal.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnLocal.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnLocal.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLocal.Location = New System.Drawing.Point(573, 58)
        Me.btnLocal.Name = "btnLocal"
        Me.btnLocal.Size = New System.Drawing.Size(103, 36)
        Me.btnLocal.TabIndex = 12
        Me.btnLocal.UseVisualStyleBackColor = False
        Me.btnLocal.Visible = False
        '
        'btnRemote
        '
        Me.btnRemote.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRemote.BackColor = System.Drawing.Color.MidnightBlue
        Me.btnRemote.BackgroundImage = CType(resources.GetObject("btnRemote.BackgroundImage"), System.Drawing.Image)
        Me.btnRemote.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnRemote.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRemote.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemote.Location = New System.Drawing.Point(573, 14)
        Me.btnRemote.Name = "btnRemote"
        Me.btnRemote.Size = New System.Drawing.Size(103, 36)
        Me.btnRemote.TabIndex = 2
        Me.btnRemote.UseVisualStyleBackColor = False
        Me.btnRemote.Visible = False
        '
        'rb3307
        '
        Me.rb3307.AutoSize = True
        Me.rb3307.BackColor = System.Drawing.Color.LightSkyBlue
        Me.rb3307.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb3307.Location = New System.Drawing.Point(491, 166)
        Me.rb3307.Name = "rb3307"
        Me.rb3307.Size = New System.Drawing.Size(63, 23)
        Me.rb3307.TabIndex = 11
        Me.rb3307.TabStop = True
        Me.rb3307.Text = "3307"
        Me.rb3307.UseVisualStyleBackColor = False
        '
        'rb3306
        '
        Me.rb3306.AutoSize = True
        Me.rb3306.BackColor = System.Drawing.Color.LightSkyBlue
        Me.rb3306.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb3306.Location = New System.Drawing.Point(400, 166)
        Me.rb3306.Name = "rb3306"
        Me.rb3306.Size = New System.Drawing.Size(63, 23)
        Me.rb3306.TabIndex = 10
        Me.rb3306.TabStop = True
        Me.rb3306.Text = "3306"
        Me.rb3306.UseVisualStyleBackColor = False
        '
        'rb6080
        '
        Me.rb6080.AutoSize = True
        Me.rb6080.BackColor = System.Drawing.Color.LightSkyBlue
        Me.rb6080.Checked = True
        Me.rb6080.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb6080.Location = New System.Drawing.Point(309, 166)
        Me.rb6080.Name = "rb6080"
        Me.rb6080.Size = New System.Drawing.Size(63, 23)
        Me.rb6080.TabIndex = 9
        Me.rb6080.TabStop = True
        Me.rb6080.Text = "6080"
        Me.rb6080.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rb6080.UseVisualStyleBackColor = False
        '
        'lblPort
        '
        Me.lblPort.AutoSize = True
        Me.lblPort.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblPort.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPort.Location = New System.Drawing.Point(16, 166)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(45, 19)
        Me.lblPort.TabIndex = 9
        Me.lblPort.Text = "Port:"
        '
        'txtPort
        '
        Me.txtPort.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPort.Location = New System.Drawing.Point(175, 163)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.ReadOnly = True
        Me.txtPort.Size = New System.Drawing.Size(116, 26)
        Me.txtPort.TabIndex = 8
        Me.txtPort.Text = "3306"
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblPassword.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.Location = New System.Drawing.Point(16, 129)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(85, 19)
        Me.lblPassword.TabIndex = 7
        Me.lblPassword.Text = "Password:"
        '
        'txtPassword
        '
        Me.txtPassword.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(175, 126)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(389, 26)
        Me.txtPassword.TabIndex = 7
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblUsername.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsername.Location = New System.Drawing.Point(16, 92)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(87, 19)
        Me.lblUsername.TabIndex = 5
        Me.lblUsername.Text = "Username:"
        '
        'txtUsername
        '
        Me.txtUsername.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.Location = New System.Drawing.Point(175, 89)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(389, 26)
        Me.txtUsername.TabIndex = 6
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblDatabase.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabase.Location = New System.Drawing.Point(16, 55)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(81, 19)
        Me.lblDatabase.TabIndex = 3
        Me.lblDatabase.Text = "Database:"
        '
        'txtDatabase
        '
        Me.txtDatabase.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabase.Location = New System.Drawing.Point(175, 52)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.ReadOnly = True
        Me.txtDatabase.Size = New System.Drawing.Size(389, 26)
        Me.txtDatabase.TabIndex = 5
        Me.txtDatabase.Text = "timesheet"
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.BackColor = System.Drawing.Color.LightSkyBlue
        Me.lblServer.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServer.Location = New System.Drawing.Point(16, 18)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(59, 19)
        Me.lblServer.TabIndex = 1
        Me.lblServer.Text = "Server:"
        '
        'txtServer
        '
        Me.txtServer.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServer.Location = New System.Drawing.Point(175, 15)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.ReadOnly = True
        Me.txtServer.Size = New System.Drawing.Size(389, 26)
        Me.txtServer.TabIndex = 4
        Me.txtServer.Text = "localhost"
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Panel2.Controls.Add(Me.btnAdvanced)
        Me.Panel2.Controls.Add(Me.btnLOGIN)
        Me.Panel2.Controls.Add(Me.btnCANCEL)
        Me.Panel2.Location = New System.Drawing.Point(14, 305)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(686, 75)
        Me.Panel2.TabIndex = 2
        '
        'btnAdvanced
        '
        Me.btnAdvanced.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAdvanced.BackColor = System.Drawing.Color.Blue
        Me.btnAdvanced.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdvanced.ForeColor = System.Drawing.Color.Lime
        Me.btnAdvanced.Location = New System.Drawing.Point(21, 9)
        Me.btnAdvanced.Name = "btnAdvanced"
        Me.btnAdvanced.Size = New System.Drawing.Size(190, 57)
        Me.btnAdvanced.TabIndex = 2
        Me.btnAdvanced.Text = "Advanced"
        Me.btnAdvanced.UseVisualStyleBackColor = False
        '
        'btnLOGIN
        '
        Me.btnLOGIN.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLOGIN.BackColor = System.Drawing.Color.MidnightBlue
        Me.btnLOGIN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnLOGIN.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLOGIN.ForeColor = System.Drawing.Color.Lime
        Me.btnLOGIN.Location = New System.Drawing.Point(243, 9)
        Me.btnLOGIN.Name = "btnLOGIN"
        Me.btnLOGIN.Size = New System.Drawing.Size(190, 57)
        Me.btnLOGIN.TabIndex = 1
        Me.btnLOGIN.Text = "Login"
        Me.btnLOGIN.UseVisualStyleBackColor = False
        '
        'btnCANCEL
        '
        Me.btnCANCEL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCANCEL.BackColor = System.Drawing.Color.Red
        Me.btnCANCEL.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCANCEL.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCANCEL.ForeColor = System.Drawing.Color.Black
        Me.btnCANCEL.Location = New System.Drawing.Point(476, 9)
        Me.btnCANCEL.Name = "btnCANCEL"
        Me.btnCANCEL.Size = New System.Drawing.Size(189, 57)
        Me.btnCANCEL.TabIndex = 3
        Me.btnCANCEL.Text = "Cancel"
        Me.btnCANCEL.UseVisualStyleBackColor = False
        '
        'btnAutoLOGIN
        '
        Me.btnAutoLOGIN.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAutoLOGIN.BackColor = System.Drawing.Color.MidnightBlue
        Me.btnAutoLOGIN.BackgroundImage = CType(resources.GetObject("btnAutoLOGIN.BackgroundImage"), System.Drawing.Image)
        Me.btnAutoLOGIN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnAutoLOGIN.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAutoLOGIN.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAutoLOGIN.Location = New System.Drawing.Point(573, 112)
        Me.btnAutoLOGIN.Name = "btnAutoLOGIN"
        Me.btnAutoLOGIN.Size = New System.Drawing.Size(103, 36)
        Me.btnAutoLOGIN.TabIndex = 13
        Me.btnAutoLOGIN.UseVisualStyleBackColor = False
        Me.btnAutoLOGIN.Visible = False
        '
        'frmMysqlLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.Indigo
        Me.ClientSize = New System.Drawing.Size(714, 389)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnlButtons)
        Me.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMysqlLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MySQL Database Login"
        Me.TopMost = True
        Me.pnlButtons.ResumeLayout(False)
        Me.pnlButtons.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlButtons As Panel
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lblPort As Label
    Friend WithEvents txtPort As TextBox
    Friend WithEvents lblPassword As Label
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents lblUsername As Label
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents lblDatabase As Label
    Friend WithEvents txtDatabase As TextBox
    Friend WithEvents lblServer As Label
    Friend WithEvents txtServer As TextBox
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnLOGIN As Button
    Friend WithEvents btnCANCEL As Button
    Friend WithEvents rb3307 As RadioButton
    Friend WithEvents rb3306 As RadioButton
    Friend WithEvents rb6080 As RadioButton
    Friend WithEvents btnClose As Button
    Friend WithEvents btnAdvanced As Button
    Friend WithEvents btnRemote As Button
    Friend WithEvents btnLocal As Button
    Friend WithEvents btnAutoLOGIN As Button
End Class
