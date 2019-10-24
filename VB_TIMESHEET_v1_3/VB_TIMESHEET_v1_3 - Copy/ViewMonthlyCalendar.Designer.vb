<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMonthly
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMonthly))
        Me.monthly = New System.Windows.Forms.MonthCalendar()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtSelectedDate = New System.Windows.Forms.TextBox()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.pnlDateTo = New System.Windows.Forms.Panel()
        Me.MonthlyDateTo = New System.Windows.Forms.MonthCalendar()
        Me.txtLastSelectedDate = New System.Windows.Forms.TextBox()
        Me.pnlLayout = New System.Windows.Forms.Panel()
        Me.img_4x4 = New System.Windows.Forms.PictureBox()
        Me.img_Empty_Square = New System.Windows.Forms.PictureBox()
        Me.img_Horiz_1x3 = New System.Windows.Forms.PictureBox()
        Me.img_Vert_3x1 = New System.Windows.Forms.PictureBox()
        Me.pnlCalendars = New System.Windows.Forms.Panel()
        Me.pnlMain.SuspendLayout()
        Me.pnlDateTo.SuspendLayout()
        Me.pnlLayout.SuspendLayout()
        CType(Me.img_4x4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_Empty_Square, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_Horiz_1x3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.img_Vert_3x1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCalendars.SuspendLayout()
        Me.SuspendLayout()
        '
        'monthly
        '
        Me.monthly.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.monthly.BackColor = System.Drawing.Color.AliceBlue
        Me.monthly.FirstDayOfWeek = System.Windows.Forms.Day.Monday
        Me.monthly.Font = New System.Drawing.Font("Cambria", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.monthly.ForeColor = System.Drawing.Color.Black
        Me.monthly.Location = New System.Drawing.Point(0, 0)
        Me.monthly.MaxSelectionCount = 1
        Me.monthly.MinDate = New Date(1970, 1, 1, 0, 0, 0, 0)
        Me.monthly.MinimumSize = New System.Drawing.Size(250, 170)
        Me.monthly.Name = "monthly"
        Me.monthly.ShowWeekNumbers = True
        Me.monthly.TabIndex = 0
        Me.monthly.TitleBackColor = System.Drawing.Color.DeepSkyBlue
        Me.monthly.TitleForeColor = System.Drawing.Color.Red
        Me.monthly.TrailingForeColor = System.Drawing.Color.CornflowerBlue
        '
        'btnClose
        '
        Me.btnClose.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnClose.BackColor = System.Drawing.Color.Red
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnClose.ForeColor = System.Drawing.Color.Black
        Me.btnClose.Location = New System.Drawing.Point(135, 315)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(90, 47)
        Me.btnClose.TabIndex = 1
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtSelectedDate
        '
        Me.txtSelectedDate.BackColor = System.Drawing.Color.AliceBlue
        Me.txtSelectedDate.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectedDate.ForeColor = System.Drawing.Color.DodgerBlue
        Me.txtSelectedDate.Location = New System.Drawing.Point(50, 20)
        Me.txtSelectedDate.Name = "txtSelectedDate"
        Me.txtSelectedDate.Size = New System.Drawing.Size(250, 26)
        Me.txtSelectedDate.TabIndex = 2
        Me.txtSelectedDate.Text = "21/06/2018"
        Me.txtSelectedDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlMain
        '
        Me.pnlMain.AutoSize = True
        Me.pnlMain.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlMain.BackColor = System.Drawing.Color.SteelBlue
        Me.pnlMain.Controls.Add(Me.pnlDateTo)
        Me.pnlMain.Controls.Add(Me.txtLastSelectedDate)
        Me.pnlMain.Controls.Add(Me.txtSelectedDate)
        Me.pnlMain.Controls.Add(Me.btnClose)
        Me.pnlMain.Controls.Add(Me.pnlLayout)
        Me.pnlMain.Controls.Add(Me.pnlCalendars)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 0)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(324, 411)
        Me.pnlMain.TabIndex = 3
        '
        'pnlDateTo
        '
        Me.pnlDateTo.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlDateTo.Controls.Add(Me.MonthlyDateTo)
        Me.pnlDateTo.Location = New System.Drawing.Point(345, 60)
        Me.pnlDateTo.Name = "pnlDateTo"
        Me.pnlDateTo.Size = New System.Drawing.Size(264, 236)
        Me.pnlDateTo.TabIndex = 7
        '
        'MonthlyDateTo
        '
        Me.MonthlyDateTo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MonthlyDateTo.BackColor = System.Drawing.Color.AliceBlue
        Me.MonthlyDateTo.FirstDayOfWeek = System.Windows.Forms.Day.Monday
        Me.MonthlyDateTo.Font = New System.Drawing.Font("Cambria", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MonthlyDateTo.ForeColor = System.Drawing.Color.Black
        Me.MonthlyDateTo.Location = New System.Drawing.Point(0, 0)
        Me.MonthlyDateTo.MaxSelectionCount = 1
        Me.MonthlyDateTo.MinDate = New Date(1970, 1, 1, 0, 0, 0, 0)
        Me.MonthlyDateTo.MinimumSize = New System.Drawing.Size(250, 170)
        Me.MonthlyDateTo.Name = "MonthlyDateTo"
        Me.MonthlyDateTo.ShowWeekNumbers = True
        Me.MonthlyDateTo.TabIndex = 0
        Me.MonthlyDateTo.TitleBackColor = System.Drawing.Color.DeepSkyBlue
        Me.MonthlyDateTo.TitleForeColor = System.Drawing.Color.Red
        Me.MonthlyDateTo.TrailingForeColor = System.Drawing.Color.CornflowerBlue
        '
        'txtLastSelectedDate
        '
        Me.txtLastSelectedDate.BackColor = System.Drawing.Color.AliceBlue
        Me.txtLastSelectedDate.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastSelectedDate.ForeColor = System.Drawing.Color.DodgerBlue
        Me.txtLastSelectedDate.Location = New System.Drawing.Point(345, 20)
        Me.txtLastSelectedDate.Name = "txtLastSelectedDate"
        Me.txtLastSelectedDate.Size = New System.Drawing.Size(250, 26)
        Me.txtLastSelectedDate.TabIndex = 7
        Me.txtLastSelectedDate.Text = "21/06/2018"
        Me.txtLastSelectedDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtLastSelectedDate.Visible = False
        '
        'pnlLayout
        '
        Me.pnlLayout.BackColor = System.Drawing.Color.LightBlue
        Me.pnlLayout.Controls.Add(Me.img_4x4)
        Me.pnlLayout.Controls.Add(Me.img_Empty_Square)
        Me.pnlLayout.Controls.Add(Me.img_Horiz_1x3)
        Me.pnlLayout.Controls.Add(Me.img_Vert_3x1)
        Me.pnlLayout.Location = New System.Drawing.Point(10, 60)
        Me.pnlLayout.Name = "pnlLayout"
        Me.pnlLayout.Size = New System.Drawing.Size(30, 150)
        Me.pnlLayout.TabIndex = 5
        '
        'img_4x4
        '
        Me.img_4x4.BackColor = System.Drawing.Color.AliceBlue
        Me.img_4x4.BackgroundImage = CType(resources.GetObject("img_4x4.BackgroundImage"), System.Drawing.Image)
        Me.img_4x4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_4x4.Location = New System.Drawing.Point(0, 120)
        Me.img_4x4.Name = "img_4x4"
        Me.img_4x4.Size = New System.Drawing.Size(30, 30)
        Me.img_4x4.TabIndex = 3
        Me.img_4x4.TabStop = False
        '
        'img_Empty_Square
        '
        Me.img_Empty_Square.BackColor = System.Drawing.Color.AliceBlue
        Me.img_Empty_Square.BackgroundImage = CType(resources.GetObject("img_Empty_Square.BackgroundImage"), System.Drawing.Image)
        Me.img_Empty_Square.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_Empty_Square.Location = New System.Drawing.Point(0, 80)
        Me.img_Empty_Square.Name = "img_Empty_Square"
        Me.img_Empty_Square.Size = New System.Drawing.Size(30, 30)
        Me.img_Empty_Square.TabIndex = 4
        Me.img_Empty_Square.TabStop = False
        '
        'img_Horiz_1x3
        '
        Me.img_Horiz_1x3.BackColor = System.Drawing.Color.AliceBlue
        Me.img_Horiz_1x3.BackgroundImage = CType(resources.GetObject("img_Horiz_1x3.BackgroundImage"), System.Drawing.Image)
        Me.img_Horiz_1x3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_Horiz_1x3.Location = New System.Drawing.Point(0, 0)
        Me.img_Horiz_1x3.Name = "img_Horiz_1x3"
        Me.img_Horiz_1x3.Size = New System.Drawing.Size(30, 30)
        Me.img_Horiz_1x3.TabIndex = 0
        Me.img_Horiz_1x3.TabStop = False
        Me.img_Horiz_1x3.Visible = False
        '
        'img_Vert_3x1
        '
        Me.img_Vert_3x1.BackColor = System.Drawing.Color.AliceBlue
        Me.img_Vert_3x1.BackgroundImage = CType(resources.GetObject("img_Vert_3x1.BackgroundImage"), System.Drawing.Image)
        Me.img_Vert_3x1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.img_Vert_3x1.Location = New System.Drawing.Point(0, 40)
        Me.img_Vert_3x1.Name = "img_Vert_3x1"
        Me.img_Vert_3x1.Size = New System.Drawing.Size(30, 30)
        Me.img_Vert_3x1.TabIndex = 1
        Me.img_Vert_3x1.TabStop = False
        Me.img_Vert_3x1.Visible = False
        '
        'pnlCalendars
        '
        Me.pnlCalendars.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlCalendars.Controls.Add(Me.monthly)
        Me.pnlCalendars.Location = New System.Drawing.Point(50, 60)
        Me.pnlCalendars.Name = "pnlCalendars"
        Me.pnlCalendars.Size = New System.Drawing.Size(264, 236)
        Me.pnlCalendars.TabIndex = 6
        '
        'frmMonthly
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.MidnightBlue
        Me.ClientSize = New System.Drawing.Size(324, 411)
        Me.Controls.Add(Me.pnlMain)
        Me.Name = "frmMonthly"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "View  Monthly Calendar"
        Me.pnlMain.ResumeLayout(False)
        Me.pnlMain.PerformLayout()
        Me.pnlDateTo.ResumeLayout(False)
        Me.pnlLayout.ResumeLayout(False)
        CType(Me.img_4x4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_Empty_Square, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_Horiz_1x3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.img_Vert_3x1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCalendars.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents monthly As MonthCalendar
    Friend WithEvents btnClose As Button
    Friend WithEvents txtSelectedDate As TextBox
    Friend WithEvents pnlMain As Panel
    Friend WithEvents img_Empty_Square As PictureBox
    Friend WithEvents img_4x4 As PictureBox
    Friend WithEvents img_Vert_3x1 As PictureBox
    Friend WithEvents img_Horiz_1x3 As PictureBox
    Friend WithEvents pnlLayout As Panel
    Friend WithEvents pnlCalendars As Panel
    Friend WithEvents txtLastSelectedDate As TextBox
    Friend WithEvents pnlDateTo As Panel
    Friend WithEvents MonthlyDateTo As MonthCalendar
End Class
