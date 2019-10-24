<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInboundOfficeEntry
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInboundOfficeEntry))
        Me.pnlTOP = New System.Windows.Forms.Panel()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.imgLineDivider7 = New System.Windows.Forms.PictureBox()
        Me.imgECP_ThumbUpLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPCarLogo = New System.Windows.Forms.PictureBox()
        Me.imgECPLogo = New System.Windows.Forms.PictureBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.pnlReferenceEntry = New System.Windows.Forms.Panel()
        Me.btnSubmitReference = New System.Windows.Forms.Button()
        Me.txtDeliveryRefEntry = New System.Windows.Forms.TextBox()
        Me.lblDeliveryDateEntry = New System.Windows.Forms.Label()
        Me.pnlDeliveryInfo = New System.Windows.Forms.Panel()
        Me.txtDueTime = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDeliveryDate = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSupplierName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSupplierCode = New System.Windows.Forms.TextBox()
        Me.lblSupplierCode = New System.Windows.Forms.Label()
        Me.pnlArrivalTime = New System.Windows.Forms.Panel()
        Me.btnSubmitTime = New System.Windows.Forms.Button()
        Me.lblArrivalTime = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnSubmitGRNNumber = New System.Windows.Forms.Button()
        Me.txtGRNNumberEntry = New System.Windows.Forms.TextBox()
        Me.lblGRNNumber = New System.Windows.Forms.Label()
        Me.msktxtArrivalDate = New System.Windows.Forms.MaskedTextBox()
        Me.msktxtArrivalTime = New System.Windows.Forms.MaskedTextBox()
        Me.pnlTOP.SuspendLayout()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlReferenceEntry.SuspendLayout()
        Me.pnlDeliveryInfo.SuspendLayout()
        Me.pnlArrivalTime.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTOP
        '
        Me.pnlTOP.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.pnlTOP.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlTOP.BackColor = System.Drawing.Color.RoyalBlue
        Me.pnlTOP.Controls.Add(Me.TextBox1)
        Me.pnlTOP.Controls.Add(Me.imgLineDivider7)
        Me.pnlTOP.Controls.Add(Me.imgECP_ThumbUpLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPCarLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPLogo)
        Me.pnlTOP.Location = New System.Drawing.Point(4, 1)
        Me.pnlTOP.Name = "pnlTOP"
        Me.pnlTOP.Size = New System.Drawing.Size(620, 97)
        Me.pnlTOP.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Cambria", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(180, 60)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(246, 32)
        Me.TextBox1.TabIndex = 118
        Me.TextBox1.Text = "INBOUND OFFICE ENTRY"
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
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnExit.BackColor = System.Drawing.Color.Red
        Me.btnExit.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.Black
        Me.btnExit.Location = New System.Drawing.Point(229, 451)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(136, 28)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Tag = "btn5"
        Me.btnExit.Text = "Close"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'pnlReferenceEntry
        '
        Me.pnlReferenceEntry.BackColor = System.Drawing.Color.AliceBlue
        Me.pnlReferenceEntry.Controls.Add(Me.btnSubmitReference)
        Me.pnlReferenceEntry.Controls.Add(Me.txtDeliveryRefEntry)
        Me.pnlReferenceEntry.Controls.Add(Me.lblDeliveryDateEntry)
        Me.pnlReferenceEntry.Location = New System.Drawing.Point(4, 100)
        Me.pnlReferenceEntry.Name = "pnlReferenceEntry"
        Me.pnlReferenceEntry.Size = New System.Drawing.Size(620, 35)
        Me.pnlReferenceEntry.TabIndex = 2
        '
        'btnSubmitReference
        '
        Me.btnSubmitReference.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnSubmitReference.BackColor = System.Drawing.Color.LimeGreen
        Me.btnSubmitReference.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSubmitReference.ForeColor = System.Drawing.Color.Black
        Me.btnSubmitReference.Location = New System.Drawing.Point(351, 3)
        Me.btnSubmitReference.Name = "btnSubmitReference"
        Me.btnSubmitReference.Size = New System.Drawing.Size(180, 28)
        Me.btnSubmitReference.TabIndex = 31
        Me.btnSubmitReference.Tag = "btn5"
        Me.btnSubmitReference.Text = "Submit Reference"
        Me.btnSubmitReference.UseVisualStyleBackColor = False
        '
        'txtDeliveryRefEntry
        '
        Me.txtDeliveryRefEntry.BackColor = System.Drawing.Color.LightGray
        Me.txtDeliveryRefEntry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDeliveryRefEntry.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryRefEntry.ForeColor = System.Drawing.Color.Black
        Me.txtDeliveryRefEntry.Location = New System.Drawing.Point(225, 6)
        Me.txtDeliveryRefEntry.Name = "txtDeliveryRefEntry"
        Me.txtDeliveryRefEntry.Size = New System.Drawing.Size(120, 23)
        Me.txtDeliveryRefEntry.TabIndex = 29
        Me.txtDeliveryRefEntry.Tag = "1"
        Me.txtDeliveryRefEntry.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblDeliveryDateEntry
        '
        Me.lblDeliveryDateEntry.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDeliveryDateEntry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDeliveryDateEntry.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeliveryDateEntry.ForeColor = System.Drawing.Color.Yellow
        Me.lblDeliveryDateEntry.Location = New System.Drawing.Point(3, 6)
        Me.lblDeliveryDateEntry.Name = "lblDeliveryDateEntry"
        Me.lblDeliveryDateEntry.Size = New System.Drawing.Size(222, 24)
        Me.lblDeliveryDateEntry.TabIndex = 30
        Me.lblDeliveryDateEntry.Text = "Enter Delivery Reference:"
        Me.lblDeliveryDateEntry.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlDeliveryInfo
        '
        Me.pnlDeliveryInfo.BackColor = System.Drawing.Color.AliceBlue
        Me.pnlDeliveryInfo.Controls.Add(Me.txtDueTime)
        Me.pnlDeliveryInfo.Controls.Add(Me.Label3)
        Me.pnlDeliveryInfo.Controls.Add(Me.txtDeliveryDate)
        Me.pnlDeliveryInfo.Controls.Add(Me.Label2)
        Me.pnlDeliveryInfo.Controls.Add(Me.txtSupplierName)
        Me.pnlDeliveryInfo.Controls.Add(Me.Label1)
        Me.pnlDeliveryInfo.Controls.Add(Me.txtSupplierCode)
        Me.pnlDeliveryInfo.Controls.Add(Me.lblSupplierCode)
        Me.pnlDeliveryInfo.Location = New System.Drawing.Point(4, 141)
        Me.pnlDeliveryInfo.Name = "pnlDeliveryInfo"
        Me.pnlDeliveryInfo.Size = New System.Drawing.Size(620, 70)
        Me.pnlDeliveryInfo.TabIndex = 32
        Me.pnlDeliveryInfo.Visible = False
        '
        'txtDueTime
        '
        Me.txtDueTime.BackColor = System.Drawing.Color.LightGray
        Me.txtDueTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDueTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueTime.ForeColor = System.Drawing.Color.Black
        Me.txtDueTime.Location = New System.Drawing.Point(487, 38)
        Me.txtDueTime.Name = "txtDueTime"
        Me.txtDueTime.Size = New System.Drawing.Size(120, 23)
        Me.txtDueTime.TabIndex = 36
        Me.txtDueTime.Tag = "1"
        Me.txtDueTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Yellow
        Me.Label3.Location = New System.Drawing.Point(351, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 24)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Due Time:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDeliveryDate
        '
        Me.txtDeliveryDate.BackColor = System.Drawing.Color.LightGray
        Me.txtDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryDate.ForeColor = System.Drawing.Color.Black
        Me.txtDeliveryDate.Location = New System.Drawing.Point(487, 3)
        Me.txtDeliveryDate.Name = "txtDeliveryDate"
        Me.txtDeliveryDate.Size = New System.Drawing.Size(120, 23)
        Me.txtDeliveryDate.TabIndex = 34
        Me.txtDeliveryDate.Tag = "1"
        Me.txtDeliveryDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Yellow
        Me.Label2.Location = New System.Drawing.Point(351, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 24)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Delivery Date:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSupplierName
        '
        Me.txtSupplierName.BackColor = System.Drawing.Color.LightGray
        Me.txtSupplierName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSupplierName.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplierName.ForeColor = System.Drawing.Color.Black
        Me.txtSupplierName.Location = New System.Drawing.Point(225, 38)
        Me.txtSupplierName.Name = "txtSupplierName"
        Me.txtSupplierName.Size = New System.Drawing.Size(120, 23)
        Me.txtSupplierName.TabIndex = 32
        Me.txtSupplierName.Tag = "1"
        Me.txtSupplierName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Yellow
        Me.Label1.Location = New System.Drawing.Point(3, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(222, 24)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Supplier Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.BackColor = System.Drawing.Color.LightGray
        Me.txtSupplierCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSupplierCode.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplierCode.ForeColor = System.Drawing.Color.Black
        Me.txtSupplierCode.Location = New System.Drawing.Point(225, 3)
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.Size = New System.Drawing.Size(120, 23)
        Me.txtSupplierCode.TabIndex = 29
        Me.txtSupplierCode.Tag = "1"
        Me.txtSupplierCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblSupplierCode
        '
        Me.lblSupplierCode.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSupplierCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSupplierCode.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSupplierCode.ForeColor = System.Drawing.Color.Yellow
        Me.lblSupplierCode.Location = New System.Drawing.Point(3, 3)
        Me.lblSupplierCode.Name = "lblSupplierCode"
        Me.lblSupplierCode.Size = New System.Drawing.Size(222, 24)
        Me.lblSupplierCode.TabIndex = 30
        Me.lblSupplierCode.Text = "Supplier Code:"
        Me.lblSupplierCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlArrivalTime
        '
        Me.pnlArrivalTime.BackColor = System.Drawing.Color.AliceBlue
        Me.pnlArrivalTime.Controls.Add(Me.msktxtArrivalTime)
        Me.pnlArrivalTime.Controls.Add(Me.msktxtArrivalDate)
        Me.pnlArrivalTime.Controls.Add(Me.btnSubmitTime)
        Me.pnlArrivalTime.Controls.Add(Me.lblArrivalTime)
        Me.pnlArrivalTime.Location = New System.Drawing.Point(4, 257)
        Me.pnlArrivalTime.Name = "pnlArrivalTime"
        Me.pnlArrivalTime.Size = New System.Drawing.Size(620, 59)
        Me.pnlArrivalTime.TabIndex = 38
        Me.pnlArrivalTime.Visible = False
        '
        'btnSubmitTime
        '
        Me.btnSubmitTime.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnSubmitTime.BackColor = System.Drawing.Color.LimeGreen
        Me.btnSubmitTime.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSubmitTime.ForeColor = System.Drawing.Color.Black
        Me.btnSubmitTime.Location = New System.Drawing.Point(225, 32)
        Me.btnSubmitTime.Name = "btnSubmitTime"
        Me.btnSubmitTime.Size = New System.Drawing.Size(180, 28)
        Me.btnSubmitTime.TabIndex = 31
        Me.btnSubmitTime.Tag = "btn5"
        Me.btnSubmitTime.Text = "Submit Time"
        Me.btnSubmitTime.UseVisualStyleBackColor = False
        '
        'lblArrivalTime
        '
        Me.lblArrivalTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblArrivalTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblArrivalTime.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArrivalTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblArrivalTime.Location = New System.Drawing.Point(3, 3)
        Me.lblArrivalTime.Name = "lblArrivalTime"
        Me.lblArrivalTime.Size = New System.Drawing.Size(222, 24)
        Me.lblArrivalTime.TabIndex = 30
        Me.lblArrivalTime.Text = "Arrival Time:"
        Me.lblArrivalTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.AliceBlue
        Me.Panel2.Controls.Add(Me.btnSubmitGRNNumber)
        Me.Panel2.Controls.Add(Me.txtGRNNumberEntry)
        Me.Panel2.Controls.Add(Me.lblGRNNumber)
        Me.Panel2.Location = New System.Drawing.Point(7, 345)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(620, 38)
        Me.Panel2.TabIndex = 39
        Me.Panel2.Visible = False
        '
        'btnSubmitGRNNumber
        '
        Me.btnSubmitGRNNumber.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnSubmitGRNNumber.BackColor = System.Drawing.Color.LimeGreen
        Me.btnSubmitGRNNumber.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSubmitGRNNumber.ForeColor = System.Drawing.Color.Black
        Me.btnSubmitGRNNumber.Location = New System.Drawing.Point(351, 6)
        Me.btnSubmitGRNNumber.Name = "btnSubmitGRNNumber"
        Me.btnSubmitGRNNumber.Size = New System.Drawing.Size(180, 28)
        Me.btnSubmitGRNNumber.TabIndex = 34
        Me.btnSubmitGRNNumber.Tag = "btn5"
        Me.btnSubmitGRNNumber.Text = "Submit GRN Number"
        Me.btnSubmitGRNNumber.UseVisualStyleBackColor = False
        '
        'txtGRNNumberEntry
        '
        Me.txtGRNNumberEntry.BackColor = System.Drawing.Color.LightGray
        Me.txtGRNNumberEntry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGRNNumberEntry.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGRNNumberEntry.ForeColor = System.Drawing.Color.Black
        Me.txtGRNNumberEntry.Location = New System.Drawing.Point(225, 8)
        Me.txtGRNNumberEntry.Name = "txtGRNNumberEntry"
        Me.txtGRNNumberEntry.Size = New System.Drawing.Size(120, 23)
        Me.txtGRNNumberEntry.TabIndex = 32
        Me.txtGRNNumberEntry.Tag = "1"
        Me.txtGRNNumberEntry.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblGRNNumber
        '
        Me.lblGRNNumber.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblGRNNumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblGRNNumber.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGRNNumber.ForeColor = System.Drawing.Color.Yellow
        Me.lblGRNNumber.Location = New System.Drawing.Point(3, 8)
        Me.lblGRNNumber.Name = "lblGRNNumber"
        Me.lblGRNNumber.Size = New System.Drawing.Size(222, 24)
        Me.lblGRNNumber.TabIndex = 33
        Me.lblGRNNumber.Text = "GRN Number:"
        Me.lblGRNNumber.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'msktxtArrivalDate
        '
        Me.msktxtArrivalDate.BackColor = System.Drawing.Color.AliceBlue
        Me.msktxtArrivalDate.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.msktxtArrivalDate.ForeColor = System.Drawing.Color.Black
        Me.msktxtArrivalDate.Location = New System.Drawing.Point(228, 3)
        Me.msktxtArrivalDate.Mask = "00/00/0000"
        Me.msktxtArrivalDate.Name = "msktxtArrivalDate"
        Me.msktxtArrivalDate.Size = New System.Drawing.Size(117, 30)
        Me.msktxtArrivalDate.TabIndex = 33
        Me.msktxtArrivalDate.ValidatingType = GetType(Date)
        '
        'msktxtArrivalTime
        '
        Me.msktxtArrivalTime.BackColor = System.Drawing.Color.AliceBlue
        Me.msktxtArrivalTime.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.msktxtArrivalTime.ForeColor = System.Drawing.Color.Black
        Me.msktxtArrivalTime.Location = New System.Drawing.Point(354, 3)
        Me.msktxtArrivalTime.Mask = "00:00:00"
        Me.msktxtArrivalTime.Name = "msktxtArrivalTime"
        Me.msktxtArrivalTime.Size = New System.Drawing.Size(117, 30)
        Me.msktxtArrivalTime.TabIndex = 34
        '
        'frmInboundOfficeEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(623, 500)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.pnlArrivalTime)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.pnlDeliveryInfo)
        Me.Controls.Add(Me.pnlReferenceEntry)
        Me.Controls.Add(Me.pnlTOP)
        Me.KeyPreview = True
        Me.Name = "frmInboundOfficeEntry"
        Me.Text = "ELECTRONIC TIMESHEET - INBOUND OFFICE ENTRY"
        Me.pnlTOP.ResumeLayout(False)
        Me.pnlTOP.PerformLayout()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlReferenceEntry.ResumeLayout(False)
        Me.pnlReferenceEntry.PerformLayout()
        Me.pnlDeliveryInfo.ResumeLayout(False)
        Me.pnlDeliveryInfo.PerformLayout()
        Me.pnlArrivalTime.ResumeLayout(False)
        Me.pnlArrivalTime.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTOP As Panel
    Friend WithEvents btnExit As Button
    Friend WithEvents imgLineDivider7 As PictureBox
    Friend WithEvents imgECP_ThumbUpLogo As PictureBox
    Friend WithEvents imgECPCarLogo As PictureBox
    Friend WithEvents imgECPLogo As PictureBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents pnlReferenceEntry As Panel
    Friend WithEvents txtDeliveryRefEntry As TextBox
    Friend WithEvents lblDeliveryDateEntry As Label
    Friend WithEvents btnSubmitReference As Button
    Friend WithEvents pnlDeliveryInfo As Panel
    Friend WithEvents txtDueTime As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtDeliveryDate As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtSupplierName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtSupplierCode As TextBox
    Friend WithEvents lblSupplierCode As Label
    Friend WithEvents pnlArrivalTime As Panel
    Friend WithEvents btnSubmitTime As Button
    Friend WithEvents lblArrivalTime As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnSubmitGRNNumber As Button
    Friend WithEvents txtGRNNumberEntry As TextBox
    Friend WithEvents lblGRNNumber As Label
    Friend WithEvents msktxtArrivalTime As MaskedTextBox
    Friend WithEvents msktxtArrivalDate As MaskedTextBox
End Class
