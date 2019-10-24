<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.pnlTOP = New System.Windows.Forms.Panel()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.imgLineDivider7 = New System.Windows.Forms.PictureBox()
        Me.imgECP_ThumbUpLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPCarLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPLogo = New System.Windows.Forms.PictureBox()
        Me.pnlTOP.SuspendLayout()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTOP
        '
        Me.pnlTOP.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlTOP.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlTOP.BackColor = System.Drawing.Color.Azure
        Me.pnlTOP.Controls.Add(Me.btnExit)
        Me.pnlTOP.Controls.Add(Me.imgLineDivider7)
        Me.pnlTOP.Controls.Add(Me.imgECP_ThumbUpLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPCarLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPLogo)
        Me.pnlTOP.Location = New System.Drawing.Point(0, 2)
        Me.pnlTOP.Name = "pnlTOP"
        Me.pnlTOP.Size = New System.Drawing.Size(786, 97)
        Me.pnlTOP.TabIndex = 1
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnExit.BackColor = System.Drawing.Color.Red
        Me.btnExit.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.Black
        Me.btnExit.Location = New System.Drawing.Point(638, 10)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(136, 28)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Tag = "btn5"
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'imgLineDivider7
        '
        Me.imgLineDivider7.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider7.Image = CType(resources.GetObject("imgLineDivider7.Image"), System.Drawing.Image)
        Me.imgLineDivider7.Location = New System.Drawing.Point(0, 49)
        Me.imgLineDivider7.Name = "imgLineDivider7"
        Me.imgLineDivider7.Size = New System.Drawing.Size(1277, 7)
        Me.imgLineDivider7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider7.TabIndex = 117
        Me.imgLineDivider7.TabStop = False
        Me.imgLineDivider7.Tag = "img6"
        '
        'imgECP_ThumbUpLogo
        '
        Me.imgECP_ThumbUpLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECP_ThumbUpLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECP_ThumbUpLogo.Image = CType(resources.GetObject("imgECP_ThumbUpLogo.Image"), System.Drawing.Image)
        Me.imgECP_ThumbUpLogo.Location = New System.Drawing.Point(391, 1)
        Me.imgECP_ThumbUpLogo.Name = "imgECP_ThumbUpLogo"
        Me.imgECP_ThumbUpLogo.Size = New System.Drawing.Size(74, 48)
        Me.imgECP_ThumbUpLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECP_ThumbUpLogo.TabIndex = 13
        Me.imgECP_ThumbUpLogo.TabStop = False
        Me.imgECP_ThumbUpLogo.Tag = "img3"
        '
        'imgECPCarLogo
        '
        Me.imgECPCarLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPCarLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPCarLogo.Image = CType(resources.GetObject("imgECPCarLogo.Image"), System.Drawing.Image)
        Me.imgECPCarLogo.Location = New System.Drawing.Point(272, 1)
        Me.imgECPCarLogo.Name = "imgECPCarLogo"
        Me.imgECPCarLogo.Size = New System.Drawing.Size(113, 48)
        Me.imgECPCarLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPCarLogo.TabIndex = 12
        Me.imgECPCarLogo.TabStop = False
        Me.imgECPCarLogo.Tag = "img2"
        '
        'imgECPLogo
        '
        Me.imgECPLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPLogo.Image = CType(resources.GetObject("imgECPLogo.Image"), System.Drawing.Image)
        Me.imgECPLogo.Location = New System.Drawing.Point(127, 1)
        Me.imgECPLogo.Name = "imgECPLogo"
        Me.imgECPLogo.Size = New System.Drawing.Size(139, 48)
        Me.imgECPLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPLogo.TabIndex = 11
        Me.imgECPLogo.TabStop = False
        Me.imgECPLogo.Tag = "img1"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 261)
        Me.Controls.Add(Me.pnlTOP)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.pnlTOP.ResumeLayout(False)
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTOP As Panel
    Friend WithEvents btnExit As Button
    Friend WithEvents imgLineDivider7 As PictureBox
    Friend WithEvents imgECP_ThumbUpLogo As PictureBox
    Friend WithEvents imgECPCarLogo As PictureBox
    Friend WithEvents imgECPLogo As PictureBox
End Class
