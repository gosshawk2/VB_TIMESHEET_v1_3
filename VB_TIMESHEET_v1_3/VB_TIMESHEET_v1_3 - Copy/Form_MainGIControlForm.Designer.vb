<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainGIForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMainGIForm))
        Me.MainGITimesheet_MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimesheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportErrorListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DatabaseTablesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReferenceProgressToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimerOnMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimerOffMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateNewUsersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimesheetToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowMainControlToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowAltControlToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtMessages = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtMainClock = New System.Windows.Forms.TextBox()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.pbarMain = New System.Windows.Forms.ProgressBar()
        Me.txtConnected = New System.Windows.Forms.TextBox()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.OnlineUsersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ViewOnlineUsersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SendMessageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MainGITimesheet_MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainGITimesheet_MenuStrip1
        '
        Me.MainGITimesheet_MenuStrip1.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainGITimesheet_MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ViewToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.TimesheetToolStripMenuItem1})
        Me.MainGITimesheet_MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MainGITimesheet_MenuStrip1.Name = "MainGITimesheet_MenuStrip1"
        Me.MainGITimesheet_MenuStrip1.Size = New System.Drawing.Size(1344, 24)
        Me.MainGITimesheet_MenuStrip1.TabIndex = 0
        Me.MainGITimesheet_MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(41, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(101, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TimesheetToolStripMenuItem, Me.ImportErrorListToolStripMenuItem, Me.DatabaseTablesToolStripMenuItem, Me.ToolStripMenuItem1, Me.ReferenceProgressToolStripMenuItem, Me.ToolStripMenuItem2, Me.OnlineUsersToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ViewToolStripMenuItem.Text = "View"
        '
        'TimesheetToolStripMenuItem
        '
        Me.TimesheetToolStripMenuItem.Name = "TimesheetToolStripMenuItem"
        Me.TimesheetToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.TimesheetToolStripMenuItem.Text = "Daily Excel Sheet"
        '
        'ImportErrorListToolStripMenuItem
        '
        Me.ImportErrorListToolStripMenuItem.Name = "ImportErrorListToolStripMenuItem"
        Me.ImportErrorListToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.ImportErrorListToolStripMenuItem.Text = "Import Error List"
        '
        'DatabaseTablesToolStripMenuItem
        '
        Me.DatabaseTablesToolStripMenuItem.Name = "DatabaseTablesToolStripMenuItem"
        Me.DatabaseTablesToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.DatabaseTablesToolStripMenuItem.Text = "Database Tables"
        '
        'ReferenceProgressToolStripMenuItem
        '
        Me.ReferenceProgressToolStripMenuItem.Name = "ReferenceProgressToolStripMenuItem"
        Me.ReferenceProgressToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.ReferenceProgressToolStripMenuItem.Text = "Reference Progress"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TimersToolStripMenuItem, Me.CreateNewUsersToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'TimersToolStripMenuItem
        '
        Me.TimersToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TimerOnMenuItem, Me.TimerOffMenuItem})
        Me.TimersToolStripMenuItem.Name = "TimersToolStripMenuItem"
        Me.TimersToolStripMenuItem.Size = New System.Drawing.Size(182, 22)
        Me.TimersToolStripMenuItem.Text = "Timers"
        '
        'TimerOnMenuItem
        '
        Me.TimerOnMenuItem.Name = "TimerOnMenuItem"
        Me.TimerOnMenuItem.Size = New System.Drawing.Size(94, 22)
        Me.TimerOnMenuItem.Text = "On"
        '
        'TimerOffMenuItem
        '
        Me.TimerOffMenuItem.Name = "TimerOffMenuItem"
        Me.TimerOffMenuItem.Size = New System.Drawing.Size(94, 22)
        Me.TimerOffMenuItem.Text = "Off"
        '
        'CreateNewUsersToolStripMenuItem
        '
        Me.CreateNewUsersToolStripMenuItem.Name = "CreateNewUsersToolStripMenuItem"
        Me.CreateNewUsersToolStripMenuItem.Size = New System.Drawing.Size(182, 22)
        Me.CreateNewUsersToolStripMenuItem.Text = "Create New Users"
        '
        'TimesheetToolStripMenuItem1
        '
        Me.TimesheetToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ShowMainControlToolStripMenuItem, Me.ShowAltControlToolStripMenuItem})
        Me.TimesheetToolStripMenuItem1.Name = "TimesheetToolStripMenuItem1"
        Me.TimesheetToolStripMenuItem1.Size = New System.Drawing.Size(83, 20)
        Me.TimesheetToolStripMenuItem1.Text = "Timesheet"
        '
        'ShowMainControlToolStripMenuItem
        '
        Me.ShowMainControlToolStripMenuItem.BackColor = System.Drawing.Color.RoyalBlue
        Me.ShowMainControlToolStripMenuItem.Name = "ShowMainControlToolStripMenuItem"
        Me.ShowMainControlToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.ShowMainControlToolStripMenuItem.Text = "Show Main Control"
        '
        'ShowAltControlToolStripMenuItem
        '
        Me.ShowAltControlToolStripMenuItem.Enabled = False
        Me.ShowAltControlToolStripMenuItem.Name = "ShowAltControlToolStripMenuItem"
        Me.ShowAltControlToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.ShowAltControlToolStripMenuItem.Text = "Show Alt Control"
        '
        'txtMessages
        '
        Me.txtMessages.BackColor = System.Drawing.Color.AliceBlue
        Me.txtMessages.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMessages.Location = New System.Drawing.Point(310, 0)
        Me.txtMessages.Multiline = True
        Me.txtMessages.Name = "txtMessages"
        Me.txtMessages.ReadOnly = True
        Me.txtMessages.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMessages.Size = New System.Drawing.Size(388, 24)
        Me.txtMessages.TabIndex = 1
        '
        'Timer1
        '
        '
        'txtMainClock
        '
        Me.txtMainClock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMainClock.BackColor = System.Drawing.Color.Black
        Me.txtMainClock.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMainClock.ForeColor = System.Drawing.Color.Red
        Me.txtMainClock.Location = New System.Drawing.Point(1028, -4)
        Me.txtMainClock.Name = "txtMainClock"
        Me.txtMainClock.Size = New System.Drawing.Size(270, 26)
        Me.txtMainClock.TabIndex = 2
        Me.txtMainClock.Text = "21/06/2018 01:00:00"
        Me.txtMainClock.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Timer2
        '
        '
        'pbarMain
        '
        Me.pbarMain.BackColor = System.Drawing.Color.Red
        Me.pbarMain.ForeColor = System.Drawing.Color.Lime
        Me.pbarMain.Location = New System.Drawing.Point(800, 1)
        Me.pbarMain.Name = "pbarMain"
        Me.pbarMain.Size = New System.Drawing.Size(200, 23)
        Me.pbarMain.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pbarMain.TabIndex = 4
        '
        'txtConnected
        '
        Me.txtConnected.BackColor = System.Drawing.Color.Lime
        Me.txtConnected.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConnected.Location = New System.Drawing.Point(701, 3)
        Me.txtConnected.Name = "txtConnected"
        Me.txtConnected.Size = New System.Drawing.Size(11, 20)
        Me.txtConnected.TabIndex = 5
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(191, 6)
        '
        'OnlineUsersToolStripMenuItem
        '
        Me.OnlineUsersToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewOnlineUsersToolStripMenuItem, Me.SendMessageToolStripMenuItem})
        Me.OnlineUsersToolStripMenuItem.Name = "OnlineUsersToolStripMenuItem"
        Me.OnlineUsersToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.OnlineUsersToolStripMenuItem.Text = "Online Users"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(191, 6)
        '
        'ViewOnlineUsersToolStripMenuItem
        '
        Me.ViewOnlineUsersToolStripMenuItem.BackColor = System.Drawing.Color.AliceBlue
        Me.ViewOnlineUsersToolStripMenuItem.Name = "ViewOnlineUsersToolStripMenuItem"
        Me.ViewOnlineUsersToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.U), System.Windows.Forms.Keys)
        Me.ViewOnlineUsersToolStripMenuItem.Size = New System.Drawing.Size(268, 22)
        Me.ViewOnlineUsersToolStripMenuItem.Text = "View Online Users"
        '
        'SendMessageToolStripMenuItem
        '
        Me.SendMessageToolStripMenuItem.Name = "SendMessageToolStripMenuItem"
        Me.SendMessageToolStripMenuItem.Size = New System.Drawing.Size(268, 22)
        Me.SendMessageToolStripMenuItem.Text = "Send Message"
        Me.SendMessageToolStripMenuItem.Visible = False
        '
        'frmMainGIForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.MidnightBlue
        Me.ClientSize = New System.Drawing.Size(1344, 761)
        Me.Controls.Add(Me.txtConnected)
        Me.Controls.Add(Me.pbarMain)
        Me.Controls.Add(Me.txtMainClock)
        Me.Controls.Add(Me.txtMessages)
        Me.Controls.Add(Me.MainGITimesheet_MenuStrip1)
        Me.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(20, 20)
        Me.MainMenuStrip = Me.MainGITimesheet_MenuStrip1
        Me.MaximumSize = New System.Drawing.Size(1360, 800)
        Me.MinimumSize = New System.Drawing.Size(1270, 700)
        Me.Name = "frmMainGIForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Goods Inwards TIMESHEET Application by D.Goss 2018 v1.1a"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.MainGITimesheet_MenuStrip1.ResumeLayout(False)
        Me.MainGITimesheet_MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MainGITimesheet_MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimesheetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimesheetToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ShowMainControlToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents txtMessages As TextBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents txtMainClock As TextBox
    Friend WithEvents Timer2 As Timer
    Friend WithEvents ImportErrorListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents pbarMain As ProgressBar
    Friend WithEvents DatabaseTablesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents txtConnected As TextBox
    Friend WithEvents ShowAltControlToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TimerOnMenuItem As ToolStripMenuItem
    Friend WithEvents TimerOffMenuItem As ToolStripMenuItem
    Friend WithEvents CreateNewUsersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReferenceProgressToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents OnlineUsersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewOnlineUsersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SendMessageToolStripMenuItem As ToolStripMenuItem
End Class
