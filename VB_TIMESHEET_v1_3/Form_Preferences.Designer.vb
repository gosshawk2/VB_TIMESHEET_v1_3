<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPreferences
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPreferences))
        Me.Frame_TOP = New System.Windows.Forms.ScrollableControl()
        Me.imgLineDivider7 = New System.Windows.Forms.PictureBox()
        Me.imgECP_ThumbUpLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPCarLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPLogo = New System.Windows.Forms.PictureBox()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.ScrollableControl1 = New System.Windows.Forms.ScrollableControl()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.txtStartForm = New System.Windows.Forms.TextBox()
        Me.lblDeliveryDate = New System.Windows.Forms.Label()
        Me.btnSAVE = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbTimesheetEntry = New System.Windows.Forms.RadioButton()
        Me.rbStatusView = New System.Windows.Forms.RadioButton()
        Me.Frame_TOP.SuspendLayout()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ScrollableControl1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame_TOP
        '
        Me.Frame_TOP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Frame_TOP.AutoScroll = True
        Me.Frame_TOP.BackColor = System.Drawing.Color.Azure
        Me.Frame_TOP.Controls.Add(Me.imgLineDivider7)
        Me.Frame_TOP.Controls.Add(Me.imgECP_ThumbUpLogo)
        Me.Frame_TOP.Controls.Add(Me.imgECPCarLogo)
        Me.Frame_TOP.Controls.Add(Me.imgECPLogo)
        Me.Frame_TOP.Controls.Add(Me.txtTitle)
        Me.Frame_TOP.Location = New System.Drawing.Point(0, 0)
        Me.Frame_TOP.Name = "Frame_TOP"
        Me.Frame_TOP.Size = New System.Drawing.Size(519, 103)
        Me.Frame_TOP.TabIndex = 116
        Me.Frame_TOP.Tag = "frmOpInput"
        Me.Frame_TOP.Text = "ScrollableControl3"
        '
        'imgLineDivider7
        '
        Me.imgLineDivider7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.imgLineDivider7.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider7.Image = CType(resources.GetObject("imgLineDivider7.Image"), System.Drawing.Image)
        Me.imgLineDivider7.Location = New System.Drawing.Point(0, 56)
        Me.imgLineDivider7.Name = "imgLineDivider7"
        Me.imgLineDivider7.Size = New System.Drawing.Size(519, 10)
        Me.imgLineDivider7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider7.TabIndex = 121
        Me.imgLineDivider7.TabStop = False
        Me.imgLineDivider7.Tag = "img6"
        '
        'imgECP_ThumbUpLogo
        '
        Me.imgECP_ThumbUpLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.imgECP_ThumbUpLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECP_ThumbUpLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECP_ThumbUpLogo.Image = CType(resources.GetObject("imgECP_ThumbUpLogo.Image"), System.Drawing.Image)
        Me.imgECP_ThumbUpLogo.Location = New System.Drawing.Point(368, 2)
        Me.imgECP_ThumbUpLogo.Name = "imgECP_ThumbUpLogo"
        Me.imgECP_ThumbUpLogo.Size = New System.Drawing.Size(150, 50)
        Me.imgECP_ThumbUpLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECP_ThumbUpLogo.TabIndex = 120
        Me.imgECP_ThumbUpLogo.TabStop = False
        Me.imgECP_ThumbUpLogo.Tag = "img3"
        '
        'imgECPCarLogo
        '
        Me.imgECPCarLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.imgECPCarLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPCarLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPCarLogo.Image = CType(resources.GetObject("imgECPCarLogo.Image"), System.Drawing.Image)
        Me.imgECPCarLogo.Location = New System.Drawing.Point(184, 2)
        Me.imgECPCarLogo.Name = "imgECPCarLogo"
        Me.imgECPCarLogo.Size = New System.Drawing.Size(150, 50)
        Me.imgECPCarLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPCarLogo.TabIndex = 119
        Me.imgECPCarLogo.TabStop = False
        Me.imgECPCarLogo.Tag = "img2"
        '
        'imgECPLogo
        '
        Me.imgECPLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPLogo.Image = CType(resources.GetObject("imgECPLogo.Image"), System.Drawing.Image)
        Me.imgECPLogo.Location = New System.Drawing.Point(0, 2)
        Me.imgECPLogo.Name = "imgECPLogo"
        Me.imgECPLogo.Size = New System.Drawing.Size(150, 50)
        Me.imgECPLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPLogo.TabIndex = 118
        Me.imgECPLogo.TabStop = False
        Me.imgECPLogo.Tag = "img1"
        '
        'txtTitle
        '
        Me.txtTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTitle.BackColor = System.Drawing.Color.LightGray
        Me.txtTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.Black
        Me.txtTitle.Location = New System.Drawing.Point(0, 72)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(519, 26)
        Me.txtTitle.TabIndex = 26
        Me.txtTitle.Tag = "21"
        Me.txtTitle.Text = "SET PREFERENCES"
        Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ScrollableControl1
        '
        Me.ScrollableControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScrollableControl1.AutoScroll = True
        Me.ScrollableControl1.BackColor = System.Drawing.Color.LightCyan
        Me.ScrollableControl1.Controls.Add(Me.GroupBox1)
        Me.ScrollableControl1.Controls.Add(Me.btnCancel)
        Me.ScrollableControl1.Controls.Add(Me.btnSAVE)
        Me.ScrollableControl1.Controls.Add(Me.txtStartForm)
        Me.ScrollableControl1.Controls.Add(Me.lblDeliveryDate)
        Me.ScrollableControl1.Controls.Add(Me.PictureBox1)
        Me.ScrollableControl1.Location = New System.Drawing.Point(3, 95)
        Me.ScrollableControl1.Name = "ScrollableControl1"
        Me.ScrollableControl1.Size = New System.Drawing.Size(519, 186)
        Me.ScrollableControl1.TabIndex = 117
        Me.ScrollableControl1.Tag = "frmOpInput"
        Me.ScrollableControl1.Text = "ScrollableControl3"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-3, 9)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(519, 10)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 121
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Tag = "img6"
        '
        'txtStartForm
        '
        Me.txtStartForm.BackColor = System.Drawing.Color.LightGray
        Me.txtStartForm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtStartForm.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartForm.ForeColor = System.Drawing.Color.Black
        Me.txtStartForm.Location = New System.Drawing.Point(199, 94)
        Me.txtStartForm.Name = "txtStartForm"
        Me.txtStartForm.Size = New System.Drawing.Size(209, 26)
        Me.txtStartForm.TabIndex = 122
        Me.txtStartForm.Tag = "1"
        Me.txtStartForm.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblDeliveryDate
        '
        Me.lblDeliveryDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDeliveryDate.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeliveryDate.ForeColor = System.Drawing.Color.Yellow
        Me.lblDeliveryDate.Location = New System.Drawing.Point(77, 95)
        Me.lblDeliveryDate.Name = "lblDeliveryDate"
        Me.lblDeliveryDate.Size = New System.Drawing.Size(114, 22)
        Me.lblDeliveryDate.TabIndex = 123
        Me.lblDeliveryDate.Text = "Starting Form:"
        Me.lblDeliveryDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSAVE
        '
        Me.btnSAVE.BackColor = System.Drawing.Color.Transparent
        Me.btnSAVE.BackgroundImage = CType(resources.GetObject("btnSAVE.BackgroundImage"), System.Drawing.Image)
        Me.btnSAVE.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSAVE.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSAVE.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSAVE.Location = New System.Drawing.Point(77, 141)
        Me.btnSAVE.Name = "btnSAVE"
        Me.btnSAVE.Size = New System.Drawing.Size(100, 45)
        Me.btnSAVE.TabIndex = 124
        Me.btnSAVE.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.Transparent
        Me.btnCancel.BackgroundImage = CType(resources.GetObject("btnCancel.BackgroundImage"), System.Drawing.Image)
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(308, 141)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 45)
        Me.btnCancel.TabIndex = 125
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.GroupBox1.Controls.Add(Me.rbStatusView)
        Me.GroupBox1.Controls.Add(Me.rbTimesheetEntry)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox1.Location = New System.Drawing.Point(77, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(331, 56)
        Me.GroupBox1.TabIndex = 126
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Startup Form:"
        '
        'rbTimesheetEntry
        '
        Me.rbTimesheetEntry.AutoSize = True
        Me.rbTimesheetEntry.BackColor = System.Drawing.Color.Aquamarine
        Me.rbTimesheetEntry.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbTimesheetEntry.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbTimesheetEntry.Location = New System.Drawing.Point(6, 21)
        Me.rbTimesheetEntry.Name = "rbTimesheetEntry"
        Me.rbTimesheetEntry.Size = New System.Drawing.Size(146, 23)
        Me.rbTimesheetEntry.TabIndex = 0
        Me.rbTimesheetEntry.TabStop = True
        Me.rbTimesheetEntry.Text = "Timesheet Entry"
        Me.rbTimesheetEntry.UseVisualStyleBackColor = False
        '
        'rbStatusView
        '
        Me.rbStatusView.AutoSize = True
        Me.rbStatusView.BackColor = System.Drawing.Color.Aquamarine
        Me.rbStatusView.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbStatusView.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbStatusView.Location = New System.Drawing.Point(210, 21)
        Me.rbStatusView.Name = "rbStatusView"
        Me.rbStatusView.Size = New System.Drawing.Size(112, 23)
        Me.rbStatusView.TabIndex = 1
        Me.rbStatusView.TabStop = True
        Me.rbStatusView.Text = "Status View"
        Me.rbStatusView.UseVisualStyleBackColor = False
        '
        'frmPreferences
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Azure
        Me.ClientSize = New System.Drawing.Size(524, 293)
        Me.Controls.Add(Me.ScrollableControl1)
        Me.Controls.Add(Me.Frame_TOP)
        Me.MinimumSize = New System.Drawing.Size(540, 140)
        Me.Name = "frmPreferences"
        Me.Text = "TIMESHEET PREFERENCES"
        Me.Frame_TOP.ResumeLayout(False)
        Me.Frame_TOP.PerformLayout()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ScrollableControl1.ResumeLayout(False)
        Me.ScrollableControl1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Frame_TOP As ScrollableControl
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents imgLineDivider7 As PictureBox
    Friend WithEvents imgECP_ThumbUpLogo As PictureBox
    Friend WithEvents imgECPCarLogo As PictureBox
    Friend WithEvents imgECPLogo As PictureBox
    Friend WithEvents ScrollableControl1 As ScrollableControl
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents txtStartForm As TextBox
    Friend WithEvents lblDeliveryDate As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnSAVE As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rbStatusView As RadioButton
    Friend WithEvents rbTimesheetEntry As RadioButton
End Class
