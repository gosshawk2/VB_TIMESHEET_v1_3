<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_GI_GRID
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_GI_GRID))
        Me.txtTotalExtras = New System.Windows.Forms.TextBox()
        Me.txtTotalShorts = New System.Windows.Forms.TextBox()
        Me.lblOpComment = New System.Windows.Forms.Label()
        Me.txtTotalHours = New System.Windows.Forms.TextBox()
        Me.lblTotalHours = New System.Windows.Forms.Label()
        Me.lblOpHash = New System.Windows.Forms.Label()
        Me.txtTotalOps = New System.Windows.Forms.TextBox()
        Me.lblTotalOps = New System.Windows.Forms.Label()
        Me.btnDeleteExtra = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnAddExtra = New System.Windows.Forms.Button()
        Me.btnDeleteShort = New System.Windows.Forms.Button()
        Me.Frame_Extra_Parts = New System.Windows.Forms.ScrollableControl()
        Me.dgvExtras = New System.Windows.Forms.DataGridView()
        Me.Frame_InboundSchedule = New System.Windows.Forms.ScrollableControl()
        Me.txtSupplier = New System.Windows.Forms.TextBox()
        Me.txtCalculatedHours = New System.Windows.Forms.TextBox()
        Me.txtEstimatedTotes = New System.Windows.Forms.TextBox()
        Me.txtEstimatedCages = New System.Windows.Forms.TextBox()
        Me.txtEstimatedPallets = New System.Windows.Forms.TextBox()
        Me.txtReadyLabel = New System.Windows.Forms.TextBox()
        Me.lblReadyLabel = New System.Windows.Forms.Label()
        Me.txtActualCases = New System.Windows.Forms.TextBox()
        Me.lblActualCases = New System.Windows.Forms.Label()
        Me.txtCartonsDue = New System.Windows.Forms.TextBox()
        Me.lblCartonsDue = New System.Windows.Forms.Label()
        Me.txtPalletsDue = New System.Windows.Forms.TextBox()
        Me.lblPalletsDue = New System.Windows.Forms.Label()
        Me.txtExpectedLines = New System.Windows.Forms.TextBox()
        Me.lblExpectedLines = New System.Windows.Forms.Label()
        Me.lblCalculatedHours = New System.Windows.Forms.Label()
        Me.lblEstimatedTotes = New System.Windows.Forms.Label()
        Me.lblEstimatedCages = New System.Windows.Forms.Label()
        Me.lblEstimatedPallets = New System.Windows.Forms.Label()
        Me.txtExpectedCases = New System.Windows.Forms.TextBox()
        Me.lblExpectedCases = New System.Windows.Forms.Label()
        Me.txtOrigin = New System.Windows.Forms.TextBox()
        Me.lblOrigin = New System.Windows.Forms.Label()
        Me.txtShift = New System.Windows.Forms.TextBox()
        Me.lblShift = New System.Windows.Forms.Label()
        Me.txtDueTime = New System.Windows.Forms.TextBox()
        Me.lblDueTime = New System.Windows.Forms.Label()
        Me.txtASNnum = New System.Windows.Forms.TextBox()
        Me.lblASNno = New System.Windows.Forms.Label()
        Me.lblSupplier = New System.Windows.Forms.Label()
        Me.txtDeliveryRef = New System.Windows.Forms.TextBox()
        Me.txtDeliveryDate = New System.Windows.Forms.TextBox()
        Me.lblDeliveryRef = New System.Windows.Forms.Label()
        Me.lblDeliveryDate = New System.Windows.Forms.Label()
        Me.imgLineDivider3 = New System.Windows.Forms.PictureBox()
        Me.Frame_OpsShortsAndExtras = New System.Windows.Forms.ScrollableControl()
        Me.Frame_Operatives = New System.Windows.Forms.ScrollableControl()
        Me.dgvOperatives = New System.Windows.Forms.DataGridView()
        Me.btnAddShort = New System.Windows.Forms.Button()
        Me.Frame_Short_Parts = New System.Windows.Forms.ScrollableControl()
        Me.dgvShorts = New System.Windows.Forms.DataGridView()
        Me.ZoomControl = New System.Windows.Forms.HScrollBar()
        Me.btnDeleteOperative = New System.Windows.Forms.Button()
        Me.btnAddOperative = New System.Windows.Forms.Button()
        Me.lblOpFinishTime = New System.Windows.Forms.Label()
        Me.lblOpStartTime = New System.Windows.Forms.Label()
        Me.lblOpActivity = New System.Windows.Forms.Label()
        Me.lblOpName = New System.Windows.Forms.Label()
        Me.imgECPCarLogo = New System.Windows.Forms.PictureBox()
        Me.lblTitle_OperationalInput = New System.Windows.Forms.Label()
        Me.lblTitle_InboundSchedule = New System.Windows.Forms.Label()
        Me.btnUpdateEmployees = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnCreateTimesheet = New System.Windows.Forms.Button()
        Me.btnLoadDeliveryRef = New System.Windows.Forms.Button()
        Me.btnSaveAndContinue = New System.Windows.Forms.Button()
        Me.imgLineDivider7 = New System.Windows.Forms.PictureBox()
        Me.imgECP_ThumbUpLogo = New System.Windows.Forms.PictureBox()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.imgECPLogo = New System.Windows.Forms.PictureBox()
        Me.imgCalendar_SelectImportDate = New System.Windows.Forms.PictureBox()
        Me.Frame_Buttons = New System.Windows.Forms.ScrollableControl()
        Me.btnImportData = New System.Windows.Forms.Button()
        Me.pnlTOP = New System.Windows.Forms.Panel()
        Me.lblGridNo = New System.Windows.Forms.Label()
        Me.rbDeliveryRef = New System.Windows.Forms.RadioButton()
        Me.txtCartonsArrived = New System.Windows.Forms.TextBox()
        Me.btnFLMFinishTime = New System.Windows.Forms.Button()
        Me.btnFLMStartTime = New System.Windows.Forms.Button()
        Me.comFLMName = New System.Windows.Forms.ComboBox()
        Me.txtFLMFinishTime = New System.Windows.Forms.TextBox()
        Me.txtFLMStartTime = New System.Windows.Forms.TextBox()
        Me.txtEmployeeNo = New System.Windows.Forms.TextBox()
        Me.lblFLMFinishTime = New System.Windows.Forms.Label()
        Me.rbASNNo = New System.Windows.Forms.RadioButton()
        Me.lblCartonsArrived = New System.Windows.Forms.Label()
        Me.lblSelectDeliveryRefASN = New System.Windows.Forms.Label()
        Me.imgCalendar_SelectDeliveryDate = New System.Windows.Forms.PictureBox()
        Me.txtPalletsArrived = New System.Windows.Forms.TextBox()
        Me.lblPalletsArrived = New System.Windows.Forms.Label()
        Me.txtSelectDeliveryDate = New System.Windows.Forms.TextBox()
        Me.lblSelectDeliveryDate = New System.Windows.Forms.Label()
        Me.imgLineDivider2 = New System.Windows.Forms.PictureBox()
        Me.Frame_SupplierCompliance = New System.Windows.Forms.ScrollableControl()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.btnBaggingReq = New System.Windows.Forms.Button()
        Me.txtBaggingReq = New System.Windows.Forms.TextBox()
        Me.btnCollar = New System.Windows.Forms.Button()
        Me.txtCollar = New System.Windows.Forms.TextBox()
        Me.btnWrapStrap = New System.Windows.Forms.Button()
        Me.txtWrapStrap = New System.Windows.Forms.TextBox()
        Me.btnPalletise = New System.Windows.Forms.Button()
        Me.txtPalletise = New System.Windows.Forms.TextBox()
        Me.txtFurtherComments = New System.Windows.Forms.TextBox()
        Me.lblTitle_FurtherComments = New System.Windows.Forms.Label()
        Me.txtCompletedComment = New System.Windows.Forms.TextBox()
        Me.btnCompleted = New System.Windows.Forms.Button()
        Me.btnIsItSafe = New System.Windows.Forms.Button()
        Me.btnArrivedOnTime = New System.Windows.Forms.Button()
        Me.txtIsItSafeComment = New System.Windows.Forms.TextBox()
        Me.txtIsItSafe = New System.Windows.Forms.TextBox()
        Me.txtIsItCompleted = New System.Windows.Forms.TextBox()
        Me.txtArrivedOnTimeComment = New System.Windows.Forms.TextBox()
        Me.txtArrivedOnTime = New System.Windows.Forms.TextBox()
        Me.comASNNo = New System.Windows.Forms.ComboBox()
        Me.lblTitle_SupplierCompliance = New System.Windows.Forms.Label()
        Me.lblFLMStartTime = New System.Windows.Forms.Label()
        Me.lblFLMName = New System.Windows.Forms.Label()
        Me.lblEmployeeNo = New System.Windows.Forms.Label()
        Me.txtControlValue = New System.Windows.Forms.TextBox()
        Me.txtLastSaved = New System.Windows.Forms.TextBox()
        Me.lblLastSaved = New System.Windows.Forms.Label()
        Me.txtControlName = New System.Windows.Forms.Label()
        Me.txtGridNo = New System.Windows.Forms.TextBox()
        Me.Frame_FLMDetails = New System.Windows.Forms.ScrollableControl()
        Me.imgLineDivider4 = New System.Windows.Forms.PictureBox()
        Me.imgLineDivider1 = New System.Windows.Forms.PictureBox()
        Me.imgLineDivider6 = New System.Windows.Forms.PictureBox()
        Me.imgLineDivider5 = New System.Windows.Forms.PictureBox()
        Me.Frame_OperationalInput = New System.Windows.Forms.ScrollableControl()
        Me.comDeliveryRef = New System.Windows.Forms.ComboBox()
        Me.pnlMain = New System.Windows.Forms.Panel()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_Extra_Parts.SuspendLayout()
        CType(Me.dgvExtras, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_InboundSchedule.SuspendLayout()
        CType(Me.imgLineDivider3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_OpsShortsAndExtras.SuspendLayout()
        Me.Frame_Operatives.SuspendLayout()
        CType(Me.dgvOperatives, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_Short_Parts.SuspendLayout()
        CType(Me.dgvShorts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgCalendar_SelectImportDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_Buttons.SuspendLayout()
        Me.pnlTOP.SuspendLayout()
        CType(Me.imgCalendar_SelectDeliveryDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLineDivider2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_SupplierCompliance.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_FLMDetails.SuspendLayout()
        CType(Me.imgLineDivider4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLineDivider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLineDivider6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgLineDivider5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame_OperationalInput.SuspendLayout()
        Me.pnlMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtTotalExtras
        '
        Me.txtTotalExtras.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTotalExtras.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalExtras.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalExtras.ForeColor = System.Drawing.Color.Yellow
        Me.txtTotalExtras.Location = New System.Drawing.Point(639, 183)
        Me.txtTotalExtras.Name = "txtTotalExtras"
        Me.txtTotalExtras.ReadOnly = True
        Me.txtTotalExtras.Size = New System.Drawing.Size(35, 23)
        Me.txtTotalExtras.TabIndex = 154
        Me.txtTotalExtras.Tag = "999"
        Me.txtTotalExtras.Text = "0"
        Me.txtTotalExtras.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtTotalShorts
        '
        Me.txtTotalShorts.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTotalShorts.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalShorts.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalShorts.ForeColor = System.Drawing.Color.Yellow
        Me.txtTotalShorts.Location = New System.Drawing.Point(248, 183)
        Me.txtTotalShorts.Name = "txtTotalShorts"
        Me.txtTotalShorts.ReadOnly = True
        Me.txtTotalShorts.Size = New System.Drawing.Size(35, 23)
        Me.txtTotalShorts.TabIndex = 153
        Me.txtTotalShorts.Tag = "998"
        Me.txtTotalShorts.Text = "0"
        Me.txtTotalShorts.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblOpComment
        '
        Me.lblOpComment.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpComment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpComment.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpComment.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpComment.Location = New System.Drawing.Point(583, 3)
        Me.lblOpComment.Name = "lblOpComment"
        Me.lblOpComment.Size = New System.Drawing.Size(99, 22)
        Me.lblOpComment.TabIndex = 152
        Me.lblOpComment.Text = "Comment"
        Me.lblOpComment.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTotalHours
        '
        Me.txtTotalHours.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTotalHours.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalHours.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalHours.ForeColor = System.Drawing.Color.Yellow
        Me.txtTotalHours.Location = New System.Drawing.Point(472, 144)
        Me.txtTotalHours.Name = "txtTotalHours"
        Me.txtTotalHours.ReadOnly = True
        Me.txtTotalHours.Size = New System.Drawing.Size(35, 23)
        Me.txtTotalHours.TabIndex = 151
        Me.txtTotalHours.Tag = "24"
        Me.txtTotalHours.Text = "0"
        Me.txtTotalHours.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblTotalHours
        '
        Me.lblTotalHours.BackColor = System.Drawing.Color.Gold
        Me.lblTotalHours.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalHours.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalHours.ForeColor = System.Drawing.Color.Black
        Me.lblTotalHours.Location = New System.Drawing.Point(387, 144)
        Me.lblTotalHours.Name = "lblTotalHours"
        Me.lblTotalHours.Size = New System.Drawing.Size(80, 22)
        Me.lblTotalHours.TabIndex = 150
        Me.lblTotalHours.Text = "Total HRS:"
        Me.lblTotalHours.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOpHash
        '
        Me.lblOpHash.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpHash.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpHash.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpHash.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpHash.Location = New System.Drawing.Point(7, 3)
        Me.lblOpHash.Name = "lblOpHash"
        Me.lblOpHash.Size = New System.Drawing.Size(28, 22)
        Me.lblOpHash.TabIndex = 149
        Me.lblOpHash.Text = "#"
        Me.lblOpHash.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTotalOps
        '
        Me.txtTotalOps.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTotalOps.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotalOps.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalOps.ForeColor = System.Drawing.Color.Yellow
        Me.txtTotalOps.Location = New System.Drawing.Point(287, 144)
        Me.txtTotalOps.Name = "txtTotalOps"
        Me.txtTotalOps.ReadOnly = True
        Me.txtTotalOps.Size = New System.Drawing.Size(35, 23)
        Me.txtTotalOps.TabIndex = 148
        Me.txtTotalOps.Tag = "23"
        Me.txtTotalOps.Text = "999"
        Me.txtTotalOps.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblTotalOps
        '
        Me.lblTotalOps.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTotalOps.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalOps.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalOps.ForeColor = System.Drawing.Color.Yellow
        Me.lblTotalOps.Location = New System.Drawing.Point(202, 144)
        Me.lblTotalOps.Name = "lblTotalOps"
        Me.lblTotalOps.Size = New System.Drawing.Size(80, 22)
        Me.lblTotalOps.TabIndex = 147
        Me.lblTotalOps.Text = "Total Ops:"
        Me.lblTotalOps.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnDeleteExtra
        '
        Me.btnDeleteExtra.BackColor = System.Drawing.Color.Red
        Me.btnDeleteExtra.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteExtra.Location = New System.Drawing.Point(532, 289)
        Me.btnDeleteExtra.Name = "btnDeleteExtra"
        Me.btnDeleteExtra.Size = New System.Drawing.Size(85, 25)
        Me.btnDeleteExtra.TabIndex = 146
        Me.btnDeleteExtra.Tag = "btn155"
        Me.btnDeleteExtra.Text = "Del Extra"
        Me.btnDeleteExtra.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-10, 169)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(700, 9)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 116
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Tag = "img6"
        '
        'btnAddExtra
        '
        Me.btnAddExtra.BackColor = System.Drawing.Color.Gold
        Me.btnAddExtra.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddExtra.Location = New System.Drawing.Point(408, 289)
        Me.btnAddExtra.Name = "btnAddExtra"
        Me.btnAddExtra.Size = New System.Drawing.Size(85, 25)
        Me.btnAddExtra.TabIndex = 145
        Me.btnAddExtra.Tag = "btn154"
        Me.btnAddExtra.Text = "Add Extra"
        Me.btnAddExtra.UseVisualStyleBackColor = False
        '
        'btnDeleteShort
        '
        Me.btnDeleteShort.BackColor = System.Drawing.Color.Red
        Me.btnDeleteShort.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteShort.Location = New System.Drawing.Point(141, 289)
        Me.btnDeleteShort.Name = "btnDeleteShort"
        Me.btnDeleteShort.Size = New System.Drawing.Size(85, 25)
        Me.btnDeleteShort.TabIndex = 120
        Me.btnDeleteShort.Tag = "btn153"
        Me.btnDeleteShort.Text = "Del Short"
        Me.btnDeleteShort.UseVisualStyleBackColor = False
        '
        'Frame_Extra_Parts
        '
        Me.Frame_Extra_Parts.AutoScroll = True
        Me.Frame_Extra_Parts.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_Extra_Parts.Controls.Add(Me.dgvExtras)
        Me.Frame_Extra_Parts.ForeColor = System.Drawing.Color.Black
        Me.Frame_Extra_Parts.Location = New System.Drawing.Point(392, 180)
        Me.Frame_Extra_Parts.Name = "Frame_Extra_Parts"
        Me.Frame_Extra_Parts.Size = New System.Drawing.Size(241, 103)
        Me.Frame_Extra_Parts.TabIndex = 144
        Me.Frame_Extra_Parts.Tag = "frmExtra"
        Me.Frame_Extra_Parts.Text = "ScrollableControl6"
        '
        'dgvExtras
        '
        Me.dgvExtras.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvExtras.Location = New System.Drawing.Point(4, 3)
        Me.dgvExtras.Name = "dgvExtras"
        Me.dgvExtras.Size = New System.Drawing.Size(232, 96)
        Me.dgvExtras.TabIndex = 1
        '
        'Frame_InboundSchedule
        '
        Me.Frame_InboundSchedule.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_InboundSchedule.Controls.Add(Me.txtSupplier)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtCalculatedHours)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtEstimatedTotes)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtEstimatedCages)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtEstimatedPallets)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtReadyLabel)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblReadyLabel)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtActualCases)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblActualCases)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtCartonsDue)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblCartonsDue)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtPalletsDue)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblPalletsDue)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtExpectedLines)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblExpectedLines)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblCalculatedHours)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblEstimatedTotes)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblEstimatedCages)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblEstimatedPallets)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtExpectedCases)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblExpectedCases)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtOrigin)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblOrigin)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtShift)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblShift)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtDueTime)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblDueTime)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtASNnum)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblASNno)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblSupplier)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtDeliveryRef)
        Me.Frame_InboundSchedule.Controls.Add(Me.txtDeliveryDate)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblDeliveryRef)
        Me.Frame_InboundSchedule.Controls.Add(Me.lblDeliveryDate)
        Me.Frame_InboundSchedule.Location = New System.Drawing.Point(0, 117)
        Me.Frame_InboundSchedule.Name = "Frame_InboundSchedule"
        Me.Frame_InboundSchedule.Size = New System.Drawing.Size(240, 500)
        Me.Frame_InboundSchedule.TabIndex = 84
        Me.Frame_InboundSchedule.Tag = "frm1"
        Me.Frame_InboundSchedule.Text = "ScrollableControl1"
        '
        'txtSupplier
        '
        Me.txtSupplier.BackColor = System.Drawing.Color.LightGray
        Me.txtSupplier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSupplier.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplier.ForeColor = System.Drawing.Color.Black
        Me.txtSupplier.Location = New System.Drawing.Point(115, 60)
        Me.txtSupplier.Multiline = True
        Me.txtSupplier.Name = "txtSupplier"
        Me.txtSupplier.Size = New System.Drawing.Size(120, 20)
        Me.txtSupplier.TabIndex = 10
        Me.txtSupplier.Tag = "3"
        Me.txtSupplier.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCalculatedHours
        '
        Me.txtCalculatedHours.BackColor = System.Drawing.Color.LightGray
        Me.txtCalculatedHours.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCalculatedHours.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCalculatedHours.ForeColor = System.Drawing.Color.Black
        Me.txtCalculatedHours.Location = New System.Drawing.Point(114, 451)
        Me.txtCalculatedHours.Name = "txtCalculatedHours"
        Me.txtCalculatedHours.Size = New System.Drawing.Size(120, 23)
        Me.txtCalculatedHours.TabIndex = 24
        Me.txtCalculatedHours.Tag = "18"
        Me.txtCalculatedHours.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEstimatedTotes
        '
        Me.txtEstimatedTotes.BackColor = System.Drawing.Color.LightGray
        Me.txtEstimatedTotes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEstimatedTotes.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEstimatedTotes.ForeColor = System.Drawing.Color.Black
        Me.txtEstimatedTotes.Location = New System.Drawing.Point(114, 424)
        Me.txtEstimatedTotes.Name = "txtEstimatedTotes"
        Me.txtEstimatedTotes.Size = New System.Drawing.Size(120, 23)
        Me.txtEstimatedTotes.TabIndex = 23
        Me.txtEstimatedTotes.Tag = "9"
        Me.txtEstimatedTotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEstimatedCages
        '
        Me.txtEstimatedCages.BackColor = System.Drawing.Color.LightGray
        Me.txtEstimatedCages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEstimatedCages.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEstimatedCages.ForeColor = System.Drawing.Color.Black
        Me.txtEstimatedCages.Location = New System.Drawing.Point(114, 397)
        Me.txtEstimatedCages.Name = "txtEstimatedCages"
        Me.txtEstimatedCages.Size = New System.Drawing.Size(120, 23)
        Me.txtEstimatedCages.TabIndex = 22
        Me.txtEstimatedCages.Tag = "8"
        Me.txtEstimatedCages.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEstimatedPallets
        '
        Me.txtEstimatedPallets.BackColor = System.Drawing.Color.LightGray
        Me.txtEstimatedPallets.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEstimatedPallets.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEstimatedPallets.ForeColor = System.Drawing.Color.Black
        Me.txtEstimatedPallets.Location = New System.Drawing.Point(115, 371)
        Me.txtEstimatedPallets.Name = "txtEstimatedPallets"
        Me.txtEstimatedPallets.Size = New System.Drawing.Size(120, 23)
        Me.txtEstimatedPallets.TabIndex = 21
        Me.txtEstimatedPallets.Tag = "7"
        Me.txtEstimatedPallets.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtReadyLabel
        '
        Me.txtReadyLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtReadyLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtReadyLabel.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReadyLabel.ForeColor = System.Drawing.Color.Black
        Me.txtReadyLabel.Location = New System.Drawing.Point(115, 343)
        Me.txtReadyLabel.Name = "txtReadyLabel"
        Me.txtReadyLabel.Size = New System.Drawing.Size(120, 25)
        Me.txtReadyLabel.TabIndex = 20
        Me.txtReadyLabel.Tag = "12"
        Me.txtReadyLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblReadyLabel
        '
        Me.lblReadyLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblReadyLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblReadyLabel.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReadyLabel.ForeColor = System.Drawing.Color.Yellow
        Me.lblReadyLabel.Location = New System.Drawing.Point(5, 343)
        Me.lblReadyLabel.Name = "lblReadyLabel"
        Me.lblReadyLabel.Size = New System.Drawing.Size(105, 22)
        Me.lblReadyLabel.TabIndex = 92
        Me.lblReadyLabel.Text = "Ready Label:"
        Me.lblReadyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtActualCases
        '
        Me.txtActualCases.BackColor = System.Drawing.Color.LightGray
        Me.txtActualCases.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtActualCases.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtActualCases.ForeColor = System.Drawing.Color.Black
        Me.txtActualCases.Location = New System.Drawing.Point(115, 316)
        Me.txtActualCases.Name = "txtActualCases"
        Me.txtActualCases.Size = New System.Drawing.Size(120, 23)
        Me.txtActualCases.TabIndex = 19
        Me.txtActualCases.Tag = "17"
        Me.txtActualCases.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblActualCases
        '
        Me.lblActualCases.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblActualCases.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblActualCases.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActualCases.ForeColor = System.Drawing.Color.Yellow
        Me.lblActualCases.Location = New System.Drawing.Point(5, 316)
        Me.lblActualCases.Name = "lblActualCases"
        Me.lblActualCases.Size = New System.Drawing.Size(105, 22)
        Me.lblActualCases.TabIndex = 90
        Me.lblActualCases.Text = "Actual Cases:"
        Me.lblActualCases.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCartonsDue
        '
        Me.txtCartonsDue.BackColor = System.Drawing.Color.LightGray
        Me.txtCartonsDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCartonsDue.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCartonsDue.ForeColor = System.Drawing.Color.Black
        Me.txtCartonsDue.Location = New System.Drawing.Point(115, 289)
        Me.txtCartonsDue.Name = "txtCartonsDue"
        Me.txtCartonsDue.Size = New System.Drawing.Size(120, 23)
        Me.txtCartonsDue.TabIndex = 18
        Me.txtCartonsDue.Tag = "10"
        Me.txtCartonsDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCartonsDue
        '
        Me.lblCartonsDue.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblCartonsDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCartonsDue.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCartonsDue.ForeColor = System.Drawing.Color.Yellow
        Me.lblCartonsDue.Location = New System.Drawing.Point(4, 289)
        Me.lblCartonsDue.Name = "lblCartonsDue"
        Me.lblCartonsDue.Size = New System.Drawing.Size(105, 22)
        Me.lblCartonsDue.TabIndex = 88
        Me.lblCartonsDue.Text = "Cartons Due:"
        Me.lblCartonsDue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPalletsDue
        '
        Me.txtPalletsDue.BackColor = System.Drawing.Color.LightGray
        Me.txtPalletsDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPalletsDue.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPalletsDue.ForeColor = System.Drawing.Color.Black
        Me.txtPalletsDue.Location = New System.Drawing.Point(115, 262)
        Me.txtPalletsDue.Name = "txtPalletsDue"
        Me.txtPalletsDue.Size = New System.Drawing.Size(120, 23)
        Me.txtPalletsDue.TabIndex = 17
        Me.txtPalletsDue.Tag = "11"
        Me.txtPalletsDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblPalletsDue
        '
        Me.lblPalletsDue.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblPalletsDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPalletsDue.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPalletsDue.ForeColor = System.Drawing.Color.Yellow
        Me.lblPalletsDue.Location = New System.Drawing.Point(5, 262)
        Me.lblPalletsDue.Name = "lblPalletsDue"
        Me.lblPalletsDue.Size = New System.Drawing.Size(105, 22)
        Me.lblPalletsDue.TabIndex = 86
        Me.lblPalletsDue.Text = "Pallets Due:"
        Me.lblPalletsDue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExpectedLines
        '
        Me.txtExpectedLines.BackColor = System.Drawing.Color.LightGray
        Me.txtExpectedLines.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExpectedLines.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExpectedLines.ForeColor = System.Drawing.Color.Black
        Me.txtExpectedLines.Location = New System.Drawing.Point(115, 235)
        Me.txtExpectedLines.Name = "txtExpectedLines"
        Me.txtExpectedLines.Size = New System.Drawing.Size(120, 23)
        Me.txtExpectedLines.TabIndex = 16
        Me.txtExpectedLines.Tag = "6"
        Me.txtExpectedLines.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblExpectedLines
        '
        Me.lblExpectedLines.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblExpectedLines.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblExpectedLines.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpectedLines.ForeColor = System.Drawing.Color.Yellow
        Me.lblExpectedLines.Location = New System.Drawing.Point(5, 235)
        Me.lblExpectedLines.Name = "lblExpectedLines"
        Me.lblExpectedLines.Size = New System.Drawing.Size(105, 22)
        Me.lblExpectedLines.TabIndex = 84
        Me.lblExpectedLines.Text = "Expected Lines:"
        Me.lblExpectedLines.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCalculatedHours
        '
        Me.lblCalculatedHours.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblCalculatedHours.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCalculatedHours.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCalculatedHours.ForeColor = System.Drawing.Color.Yellow
        Me.lblCalculatedHours.Location = New System.Drawing.Point(5, 451)
        Me.lblCalculatedHours.Name = "lblCalculatedHours"
        Me.lblCalculatedHours.Size = New System.Drawing.Size(105, 22)
        Me.lblCalculatedHours.TabIndex = 83
        Me.lblCalculatedHours.Text = "Calculated Hours:"
        Me.lblCalculatedHours.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEstimatedTotes
        '
        Me.lblEstimatedTotes.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstimatedTotes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstimatedTotes.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEstimatedTotes.ForeColor = System.Drawing.Color.Yellow
        Me.lblEstimatedTotes.Location = New System.Drawing.Point(5, 424)
        Me.lblEstimatedTotes.Name = "lblEstimatedTotes"
        Me.lblEstimatedTotes.Size = New System.Drawing.Size(105, 22)
        Me.lblEstimatedTotes.TabIndex = 82
        Me.lblEstimatedTotes.Text = "Estimated Totes:"
        Me.lblEstimatedTotes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEstimatedCages
        '
        Me.lblEstimatedCages.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstimatedCages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstimatedCages.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEstimatedCages.ForeColor = System.Drawing.Color.Yellow
        Me.lblEstimatedCages.Location = New System.Drawing.Point(5, 397)
        Me.lblEstimatedCages.Name = "lblEstimatedCages"
        Me.lblEstimatedCages.Size = New System.Drawing.Size(105, 22)
        Me.lblEstimatedCages.TabIndex = 81
        Me.lblEstimatedCages.Text = "Estimated Cages:"
        Me.lblEstimatedCages.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEstimatedPallets
        '
        Me.lblEstimatedPallets.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstimatedPallets.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstimatedPallets.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEstimatedPallets.ForeColor = System.Drawing.Color.Yellow
        Me.lblEstimatedPallets.Location = New System.Drawing.Point(5, 371)
        Me.lblEstimatedPallets.Name = "lblEstimatedPallets"
        Me.lblEstimatedPallets.Size = New System.Drawing.Size(105, 22)
        Me.lblEstimatedPallets.TabIndex = 80
        Me.lblEstimatedPallets.Text = "Estimated Pallets:"
        Me.lblEstimatedPallets.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExpectedCases
        '
        Me.txtExpectedCases.BackColor = System.Drawing.Color.LightGray
        Me.txtExpectedCases.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExpectedCases.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExpectedCases.ForeColor = System.Drawing.Color.Black
        Me.txtExpectedCases.Location = New System.Drawing.Point(115, 208)
        Me.txtExpectedCases.Name = "txtExpectedCases"
        Me.txtExpectedCases.Size = New System.Drawing.Size(120, 23)
        Me.txtExpectedCases.TabIndex = 15
        Me.txtExpectedCases.Tag = "5"
        Me.txtExpectedCases.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblExpectedCases
        '
        Me.lblExpectedCases.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblExpectedCases.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblExpectedCases.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpectedCases.ForeColor = System.Drawing.Color.Yellow
        Me.lblExpectedCases.Location = New System.Drawing.Point(5, 208)
        Me.lblExpectedCases.Name = "lblExpectedCases"
        Me.lblExpectedCases.Size = New System.Drawing.Size(105, 22)
        Me.lblExpectedCases.TabIndex = 78
        Me.lblExpectedCases.Text = "Expected Cases:"
        Me.lblExpectedCases.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOrigin
        '
        Me.txtOrigin.BackColor = System.Drawing.Color.LightGray
        Me.txtOrigin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOrigin.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOrigin.ForeColor = System.Drawing.Color.Black
        Me.txtOrigin.Location = New System.Drawing.Point(115, 181)
        Me.txtOrigin.Name = "txtOrigin"
        Me.txtOrigin.Size = New System.Drawing.Size(120, 23)
        Me.txtOrigin.TabIndex = 14
        Me.txtOrigin.Tag = "25"
        Me.txtOrigin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblOrigin
        '
        Me.lblOrigin.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOrigin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOrigin.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrigin.ForeColor = System.Drawing.Color.Yellow
        Me.lblOrigin.Location = New System.Drawing.Point(5, 181)
        Me.lblOrigin.Name = "lblOrigin"
        Me.lblOrigin.Size = New System.Drawing.Size(105, 22)
        Me.lblOrigin.TabIndex = 76
        Me.lblOrigin.Text = "Origin:"
        Me.lblOrigin.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtShift
        '
        Me.txtShift.BackColor = System.Drawing.Color.LightGray
        Me.txtShift.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtShift.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShift.ForeColor = System.Drawing.Color.Black
        Me.txtShift.Location = New System.Drawing.Point(115, 141)
        Me.txtShift.Name = "txtShift"
        Me.txtShift.Size = New System.Drawing.Size(120, 23)
        Me.txtShift.TabIndex = 13
        Me.txtShift.Tag = "27"
        Me.txtShift.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblShift
        '
        Me.lblShift.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblShift.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblShift.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShift.ForeColor = System.Drawing.Color.Yellow
        Me.lblShift.Location = New System.Drawing.Point(5, 141)
        Me.lblShift.Name = "lblShift"
        Me.lblShift.Size = New System.Drawing.Size(105, 22)
        Me.lblShift.TabIndex = 74
        Me.lblShift.Text = "Shift:"
        Me.lblShift.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDueTime
        '
        Me.txtDueTime.BackColor = System.Drawing.Color.LightGray
        Me.txtDueTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDueTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueTime.ForeColor = System.Drawing.Color.Black
        Me.txtDueTime.Location = New System.Drawing.Point(115, 114)
        Me.txtDueTime.Name = "txtDueTime"
        Me.txtDueTime.Size = New System.Drawing.Size(120, 23)
        Me.txtDueTime.TabIndex = 12
        Me.txtDueTime.Tag = "26"
        Me.txtDueTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblDueTime
        '
        Me.lblDueTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDueTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDueTime.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDueTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblDueTime.Location = New System.Drawing.Point(5, 114)
        Me.lblDueTime.Name = "lblDueTime"
        Me.lblDueTime.Size = New System.Drawing.Size(105, 22)
        Me.lblDueTime.TabIndex = 72
        Me.lblDueTime.Text = "Due Time:"
        Me.lblDueTime.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtASNnum
        '
        Me.txtASNnum.BackColor = System.Drawing.Color.White
        Me.txtASNnum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtASNnum.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtASNnum.ForeColor = System.Drawing.Color.Black
        Me.txtASNnum.Location = New System.Drawing.Point(115, 86)
        Me.txtASNnum.Name = "txtASNnum"
        Me.txtASNnum.Size = New System.Drawing.Size(120, 23)
        Me.txtASNnum.TabIndex = 11
        Me.txtASNnum.Tag = "4"
        Me.txtASNnum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblASNno
        '
        Me.lblASNno.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblASNno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblASNno.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblASNno.ForeColor = System.Drawing.Color.Yellow
        Me.lblASNno.Location = New System.Drawing.Point(5, 86)
        Me.lblASNno.Name = "lblASNno"
        Me.lblASNno.Size = New System.Drawing.Size(105, 22)
        Me.lblASNno.TabIndex = 70
        Me.lblASNno.Text = "ASN no:"
        Me.lblASNno.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSupplier
        '
        Me.lblSupplier.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSupplier.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSupplier.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSupplier.ForeColor = System.Drawing.Color.Yellow
        Me.lblSupplier.Location = New System.Drawing.Point(5, 59)
        Me.lblSupplier.Name = "lblSupplier"
        Me.lblSupplier.Size = New System.Drawing.Size(105, 22)
        Me.lblSupplier.TabIndex = 68
        Me.lblSupplier.Text = "Supplier"
        Me.lblSupplier.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDeliveryRef
        '
        Me.txtDeliveryRef.BackColor = System.Drawing.Color.LightGray
        Me.txtDeliveryRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDeliveryRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryRef.ForeColor = System.Drawing.Color.Black
        Me.txtDeliveryRef.Location = New System.Drawing.Point(115, 33)
        Me.txtDeliveryRef.Name = "txtDeliveryRef"
        Me.txtDeliveryRef.Size = New System.Drawing.Size(120, 23)
        Me.txtDeliveryRef.TabIndex = 9
        Me.txtDeliveryRef.Tag = "2"
        Me.txtDeliveryRef.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDeliveryDate
        '
        Me.txtDeliveryDate.BackColor = System.Drawing.Color.LightGray
        Me.txtDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryDate.ForeColor = System.Drawing.Color.Black
        Me.txtDeliveryDate.Location = New System.Drawing.Point(115, 3)
        Me.txtDeliveryDate.Name = "txtDeliveryDate"
        Me.txtDeliveryDate.Size = New System.Drawing.Size(120, 23)
        Me.txtDeliveryDate.TabIndex = 8
        Me.txtDeliveryDate.Tag = "1"
        Me.txtDeliveryDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblDeliveryRef
        '
        Me.lblDeliveryRef.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDeliveryRef.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDeliveryRef.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeliveryRef.ForeColor = System.Drawing.Color.Yellow
        Me.lblDeliveryRef.Location = New System.Drawing.Point(5, 33)
        Me.lblDeliveryRef.Name = "lblDeliveryRef"
        Me.lblDeliveryRef.Size = New System.Drawing.Size(105, 22)
        Me.lblDeliveryRef.TabIndex = 29
        Me.lblDeliveryRef.Text = "Delivery Ref:"
        Me.lblDeliveryRef.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDeliveryDate
        '
        Me.lblDeliveryDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeliveryDate.ForeColor = System.Drawing.Color.Yellow
        Me.lblDeliveryDate.Location = New System.Drawing.Point(5, 3)
        Me.lblDeliveryDate.Name = "lblDeliveryDate"
        Me.lblDeliveryDate.Size = New System.Drawing.Size(105, 22)
        Me.lblDeliveryDate.TabIndex = 28
        Me.lblDeliveryDate.Text = "Delivery Date:"
        Me.lblDeliveryDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'imgLineDivider3
        '
        Me.imgLineDivider3.AccessibleName = "0, 112, 192"
        Me.imgLineDivider3.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider3.Image = CType(resources.GetObject("imgLineDivider3.Image"), System.Drawing.Image)
        Me.imgLineDivider3.Location = New System.Drawing.Point(240, 80)
        Me.imgLineDivider3.Name = "imgLineDivider3"
        Me.imgLineDivider3.Size = New System.Drawing.Size(10, 540)
        Me.imgLineDivider3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider3.TabIndex = 82
        Me.imgLineDivider3.TabStop = False
        Me.imgLineDivider3.Tag = "img4"
        '
        'Frame_OpsShortsAndExtras
        '
        Me.Frame_OpsShortsAndExtras.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Frame_OpsShortsAndExtras.AutoScroll = True
        Me.Frame_OpsShortsAndExtras.BackColor = System.Drawing.Color.Azure
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.Frame_Operatives)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.txtTotalExtras)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.txtTotalShorts)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpComment)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.txtTotalHours)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblTotalHours)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpHash)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.txtTotalOps)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblTotalOps)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnDeleteExtra)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.PictureBox1)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnAddExtra)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnDeleteShort)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.Frame_Extra_Parts)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnAddShort)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.Frame_Short_Parts)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.ZoomControl)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnDeleteOperative)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.btnAddOperative)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpFinishTime)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpStartTime)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpActivity)
        Me.Frame_OpsShortsAndExtras.Controls.Add(Me.lblOpName)
        Me.Frame_OpsShortsAndExtras.ForeColor = System.Drawing.Color.Black
        Me.Frame_OpsShortsAndExtras.Location = New System.Drawing.Point(250, 295)
        Me.Frame_OpsShortsAndExtras.Name = "Frame_OpsShortsAndExtras"
        Me.Frame_OpsShortsAndExtras.Size = New System.Drawing.Size(690, 325)
        Me.Frame_OpsShortsAndExtras.TabIndex = 113
        Me.Frame_OpsShortsAndExtras.Tag = "frm1"
        Me.Frame_OpsShortsAndExtras.Text = "ScrollableControl3"
        '
        'Frame_Operatives
        '
        Me.Frame_Operatives.Controls.Add(Me.dgvOperatives)
        Me.Frame_Operatives.Location = New System.Drawing.Point(1, 28)
        Me.Frame_Operatives.Name = "Frame_Operatives"
        Me.Frame_Operatives.Size = New System.Drawing.Size(688, 108)
        Me.Frame_Operatives.TabIndex = 155
        Me.Frame_Operatives.Text = "ScrollableControl1"
        '
        'dgvOperatives
        '
        Me.dgvOperatives.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOperatives.Location = New System.Drawing.Point(5, 3)
        Me.dgvOperatives.Name = "dgvOperatives"
        Me.dgvOperatives.Size = New System.Drawing.Size(676, 102)
        Me.dgvOperatives.TabIndex = 0
        '
        'btnAddShort
        '
        Me.btnAddShort.BackColor = System.Drawing.Color.Gold
        Me.btnAddShort.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddShort.Location = New System.Drawing.Point(9, 289)
        Me.btnAddShort.Name = "btnAddShort"
        Me.btnAddShort.Size = New System.Drawing.Size(85, 25)
        Me.btnAddShort.TabIndex = 119
        Me.btnAddShort.Tag = "btn152"
        Me.btnAddShort.Text = "Add Short"
        Me.btnAddShort.UseVisualStyleBackColor = False
        '
        'Frame_Short_Parts
        '
        Me.Frame_Short_Parts.AutoScroll = True
        Me.Frame_Short_Parts.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_Short_Parts.Controls.Add(Me.dgvShorts)
        Me.Frame_Short_Parts.ForeColor = System.Drawing.Color.Black
        Me.Frame_Short_Parts.Location = New System.Drawing.Point(1, 180)
        Me.Frame_Short_Parts.Name = "Frame_Short_Parts"
        Me.Frame_Short_Parts.Size = New System.Drawing.Size(241, 103)
        Me.Frame_Short_Parts.TabIndex = 118
        Me.Frame_Short_Parts.Tag = "frmShorts"
        Me.Frame_Short_Parts.Text = "ScrollableControl5"
        '
        'dgvShorts
        '
        Me.dgvShorts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShorts.Location = New System.Drawing.Point(6, 4)
        Me.dgvShorts.Name = "dgvShorts"
        Me.dgvShorts.Size = New System.Drawing.Size(232, 96)
        Me.dgvShorts.TabIndex = 0
        '
        'ZoomControl
        '
        Me.ZoomControl.Location = New System.Drawing.Point(272, 211)
        Me.ZoomControl.Name = "ZoomControl"
        Me.ZoomControl.Size = New System.Drawing.Size(94, 25)
        Me.ZoomControl.TabIndex = 117
        Me.ZoomControl.Value = 50
        Me.ZoomControl.Visible = False
        '
        'btnDeleteOperative
        '
        Me.btnDeleteOperative.BackColor = System.Drawing.Color.Red
        Me.btnDeleteOperative.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteOperative.ForeColor = System.Drawing.Color.Black
        Me.btnDeleteOperative.Location = New System.Drawing.Point(528, 142)
        Me.btnDeleteOperative.Name = "btnDeleteOperative"
        Me.btnDeleteOperative.Size = New System.Drawing.Size(135, 25)
        Me.btnDeleteOperative.TabIndex = 116
        Me.btnDeleteOperative.Tag = "btn151"
        Me.btnDeleteOperative.Text = "Delete Operative"
        Me.btnDeleteOperative.UseVisualStyleBackColor = False
        '
        'btnAddOperative
        '
        Me.btnAddOperative.BackColor = System.Drawing.Color.Gold
        Me.btnAddOperative.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddOperative.Location = New System.Drawing.Point(30, 142)
        Me.btnAddOperative.Name = "btnAddOperative"
        Me.btnAddOperative.Size = New System.Drawing.Size(145, 25)
        Me.btnAddOperative.TabIndex = 115
        Me.btnAddOperative.Tag = "btn150"
        Me.btnAddOperative.Text = "Add Operative"
        Me.btnAddOperative.UseVisualStyleBackColor = False
        '
        'lblOpFinishTime
        '
        Me.lblOpFinishTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpFinishTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpFinishTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpFinishTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpFinishTime.Location = New System.Drawing.Point(463, 3)
        Me.lblOpFinishTime.Name = "lblOpFinishTime"
        Me.lblOpFinishTime.Size = New System.Drawing.Size(110, 22)
        Me.lblOpFinishTime.TabIndex = 109
        Me.lblOpFinishTime.Text = "Finish Time"
        Me.lblOpFinishTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOpStartTime
        '
        Me.lblOpStartTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpStartTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpStartTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpStartTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpStartTime.Location = New System.Drawing.Point(343, 3)
        Me.lblOpStartTime.Name = "lblOpStartTime"
        Me.lblOpStartTime.Size = New System.Drawing.Size(110, 22)
        Me.lblOpStartTime.TabIndex = 108
        Me.lblOpStartTime.Text = "Start Time"
        Me.lblOpStartTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOpActivity
        '
        Me.lblOpActivity.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpActivity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpActivity.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpActivity.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpActivity.Location = New System.Drawing.Point(206, 3)
        Me.lblOpActivity.Name = "lblOpActivity"
        Me.lblOpActivity.Size = New System.Drawing.Size(129, 22)
        Me.lblOpActivity.TabIndex = 107
        Me.lblOpActivity.Text = "Operative Activity"
        Me.lblOpActivity.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblOpName
        '
        Me.lblOpName.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblOpName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOpName.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpName.ForeColor = System.Drawing.Color.Yellow
        Me.lblOpName.Location = New System.Drawing.Point(41, 3)
        Me.lblOpName.Name = "lblOpName"
        Me.lblOpName.Size = New System.Drawing.Size(159, 22)
        Me.lblOpName.TabIndex = 106
        Me.lblOpName.Text = "Operative Name"
        Me.lblOpName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'imgECPCarLogo
        '
        Me.imgECPCarLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPCarLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPCarLogo.Image = CType(resources.GetObject("imgECPCarLogo.Image"), System.Drawing.Image)
        Me.imgECPCarLogo.Location = New System.Drawing.Point(497, 1)
        Me.imgECPCarLogo.Name = "imgECPCarLogo"
        Me.imgECPCarLogo.Size = New System.Drawing.Size(250, 40)
        Me.imgECPCarLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPCarLogo.TabIndex = 12
        Me.imgECPCarLogo.TabStop = False
        Me.imgECPCarLogo.Tag = "img2"
        '
        'lblTitle_OperationalInput
        '
        Me.lblTitle_OperationalInput.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTitle_OperationalInput.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle_OperationalInput.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitle_OperationalInput.Location = New System.Drawing.Point(250, 86)
        Me.lblTitle_OperationalInput.Name = "lblTitle_OperationalInput"
        Me.lblTitle_OperationalInput.Size = New System.Drawing.Size(690, 23)
        Me.lblTitle_OperationalInput.TabIndex = 81
        Me.lblTitle_OperationalInput.Text = "Operational Input"
        Me.lblTitle_OperationalInput.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTitle_InboundSchedule
        '
        Me.lblTitle_InboundSchedule.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTitle_InboundSchedule.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle_InboundSchedule.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitle_InboundSchedule.Location = New System.Drawing.Point(0, 86)
        Me.lblTitle_InboundSchedule.Name = "lblTitle_InboundSchedule"
        Me.lblTitle_InboundSchedule.Size = New System.Drawing.Size(240, 23)
        Me.lblTitle_InboundSchedule.TabIndex = 17
        Me.lblTitle_InboundSchedule.Text = "Inbound Schedule"
        Me.lblTitle_InboundSchedule.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnUpdateEmployees
        '
        Me.btnUpdateEmployees.BackColor = System.Drawing.Color.Gold
        Me.btnUpdateEmployees.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateEmployees.Location = New System.Drawing.Point(924, 3)
        Me.btnUpdateEmployees.Name = "btnUpdateEmployees"
        Me.btnUpdateEmployees.Size = New System.Drawing.Size(150, 28)
        Me.btnUpdateEmployees.TabIndex = 6
        Me.btnUpdateEmployees.Tag = "btn6"
        Me.btnUpdateEmployees.Text = "Update Employees"
        Me.btnUpdateEmployees.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.Red
        Me.btnExit.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.Black
        Me.btnExit.Location = New System.Drawing.Point(1125, 2)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(105, 28)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Tag = "btn5"
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnCreateTimesheet
        '
        Me.btnCreateTimesheet.BackColor = System.Drawing.Color.Gold
        Me.btnCreateTimesheet.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateTimesheet.Location = New System.Drawing.Point(572, 3)
        Me.btnCreateTimesheet.Name = "btnCreateTimesheet"
        Me.btnCreateTimesheet.Size = New System.Drawing.Size(150, 28)
        Me.btnCreateTimesheet.TabIndex = 4
        Me.btnCreateTimesheet.Tag = "btn4"
        Me.btnCreateTimesheet.Text = "Create Timesheet"
        Me.btnCreateTimesheet.UseVisualStyleBackColor = False
        '
        'btnLoadDeliveryRef
        '
        Me.btnLoadDeliveryRef.BackColor = System.Drawing.Color.Gold
        Me.btnLoadDeliveryRef.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLoadDeliveryRef.Location = New System.Drawing.Point(412, 3)
        Me.btnLoadDeliveryRef.Name = "btnLoadDeliveryRef"
        Me.btnLoadDeliveryRef.Size = New System.Drawing.Size(150, 28)
        Me.btnLoadDeliveryRef.TabIndex = 3
        Me.btnLoadDeliveryRef.Tag = "btn3"
        Me.btnLoadDeliveryRef.Text = "&Load Delivery Ref"
        Me.btnLoadDeliveryRef.UseVisualStyleBackColor = False
        '
        'btnSaveAndContinue
        '
        Me.btnSaveAndContinue.BackColor = System.Drawing.Color.Gold
        Me.btnSaveAndContinue.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveAndContinue.Location = New System.Drawing.Point(62, 3)
        Me.btnSaveAndContinue.Name = "btnSaveAndContinue"
        Me.btnSaveAndContinue.Size = New System.Drawing.Size(179, 27)
        Me.btnSaveAndContinue.TabIndex = 1
        Me.btnSaveAndContinue.Tag = "btn1"
        Me.btnSaveAndContinue.Text = "&Save And Continue"
        Me.btnSaveAndContinue.UseVisualStyleBackColor = False
        '
        'imgLineDivider7
        '
        Me.imgLineDivider7.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider7.Image = CType(resources.GetObject("imgLineDivider7.Image"), System.Drawing.Image)
        Me.imgLineDivider7.Location = New System.Drawing.Point(0, 40)
        Me.imgLineDivider7.Name = "imgLineDivider7"
        Me.imgLineDivider7.Size = New System.Drawing.Size(1259, 9)
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
        Me.imgECP_ThumbUpLogo.Location = New System.Drawing.Point(922, 1)
        Me.imgECP_ThumbUpLogo.Name = "imgECP_ThumbUpLogo"
        Me.imgECP_ThumbUpLogo.Size = New System.Drawing.Size(150, 40)
        Me.imgECP_ThumbUpLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECP_ThumbUpLogo.TabIndex = 13
        Me.imgECP_ThumbUpLogo.TabStop = False
        Me.imgECP_ThumbUpLogo.Tag = "img3"
        '
        'btnReset
        '
        Me.btnReset.BackColor = System.Drawing.Color.Gold
        Me.btnReset.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReset.Location = New System.Drawing.Point(252, 3)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(150, 28)
        Me.btnReset.TabIndex = 2
        Me.btnReset.Tag = "btn2"
        Me.btnReset.Text = "&Reset"
        Me.btnReset.UseVisualStyleBackColor = False
        '
        'imgECPLogo
        '
        Me.imgECPLogo.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgECPLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.imgECPLogo.Image = CType(resources.GetObject("imgECPLogo.Image"), System.Drawing.Image)
        Me.imgECPLogo.Location = New System.Drawing.Point(39, 1)
        Me.imgECPLogo.Name = "imgECPLogo"
        Me.imgECPLogo.Size = New System.Drawing.Size(250, 40)
        Me.imgECPLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgECPLogo.TabIndex = 11
        Me.imgECPLogo.TabStop = False
        Me.imgECPLogo.Tag = "img1"
        '
        'imgCalendar_SelectImportDate
        '
        Me.imgCalendar_SelectImportDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgCalendar_SelectImportDate.Image = CType(resources.GetObject("imgCalendar_SelectImportDate.Image"), System.Drawing.Image)
        Me.imgCalendar_SelectImportDate.Location = New System.Drawing.Point(890, 1)
        Me.imgCalendar_SelectImportDate.Name = "imgCalendar_SelectImportDate"
        Me.imgCalendar_SelectImportDate.Size = New System.Drawing.Size(24, 28)
        Me.imgCalendar_SelectImportDate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgCalendar_SelectImportDate.TabIndex = 136
        Me.imgCalendar_SelectImportDate.TabStop = False
        Me.imgCalendar_SelectImportDate.Tag = "btn7"
        '
        'Frame_Buttons
        '
        Me.Frame_Buttons.AutoScroll = True
        Me.Frame_Buttons.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_Buttons.Controls.Add(Me.imgCalendar_SelectImportDate)
        Me.Frame_Buttons.Controls.Add(Me.btnImportData)
        Me.Frame_Buttons.Controls.Add(Me.btnUpdateEmployees)
        Me.Frame_Buttons.Controls.Add(Me.btnCreateTimesheet)
        Me.Frame_Buttons.Controls.Add(Me.btnExit)
        Me.Frame_Buttons.Controls.Add(Me.btnLoadDeliveryRef)
        Me.Frame_Buttons.Controls.Add(Me.btnReset)
        Me.Frame_Buttons.Controls.Add(Me.btnSaveAndContinue)
        Me.Frame_Buttons.Location = New System.Drawing.Point(1, 47)
        Me.Frame_Buttons.Name = "Frame_Buttons"
        Me.Frame_Buttons.Size = New System.Drawing.Size(1246, 32)
        Me.Frame_Buttons.TabIndex = 118
        Me.Frame_Buttons.Tag = "frmButtons"
        Me.Frame_Buttons.Text = "ScrollableControl2"
        '
        'btnImportData
        '
        Me.btnImportData.BackColor = System.Drawing.Color.Gold
        Me.btnImportData.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImportData.Location = New System.Drawing.Point(732, 3)
        Me.btnImportData.Name = "btnImportData"
        Me.btnImportData.Size = New System.Drawing.Size(150, 28)
        Me.btnImportData.TabIndex = 8
        Me.btnImportData.Tag = "btn4"
        Me.btnImportData.Text = "Import Data"
        Me.btnImportData.UseVisualStyleBackColor = False
        '
        'pnlTOP
        '
        Me.pnlTOP.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlTOP.BackColor = System.Drawing.Color.Azure
        Me.pnlTOP.Controls.Add(Me.Frame_Buttons)
        Me.pnlTOP.Controls.Add(Me.imgLineDivider7)
        Me.pnlTOP.Controls.Add(Me.imgECP_ThumbUpLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPCarLogo)
        Me.pnlTOP.Controls.Add(Me.imgECPLogo)
        Me.pnlTOP.Location = New System.Drawing.Point(6, 4)
        Me.pnlTOP.Name = "pnlTOP"
        Me.pnlTOP.Size = New System.Drawing.Size(1260, 82)
        Me.pnlTOP.TabIndex = 1
        '
        'lblGridNo
        '
        Me.lblGridNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblGridNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblGridNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGridNo.ForeColor = System.Drawing.Color.Yellow
        Me.lblGridNo.Location = New System.Drawing.Point(397, 30)
        Me.lblGridNo.Name = "lblGridNo"
        Me.lblGridNo.Size = New System.Drawing.Size(135, 22)
        Me.lblGridNo.TabIndex = 130
        Me.lblGridNo.Text = "Enter Grid No."
        Me.lblGridNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'rbDeliveryRef
        '
        Me.rbDeliveryRef.AutoSize = True
        Me.rbDeliveryRef.Checked = True
        Me.rbDeliveryRef.Location = New System.Drawing.Point(402, 60)
        Me.rbDeliveryRef.Name = "rbDeliveryRef"
        Me.rbDeliveryRef.Size = New System.Drawing.Size(83, 17)
        Me.rbDeliveryRef.TabIndex = 29
        Me.rbDeliveryRef.TabStop = True
        Me.rbDeliveryRef.Text = "Delivery Ref"
        Me.rbDeliveryRef.UseVisualStyleBackColor = True
        '
        'txtCartonsArrived
        '
        Me.txtCartonsArrived.BackColor = System.Drawing.Color.LightGray
        Me.txtCartonsArrived.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCartonsArrived.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCartonsArrived.ForeColor = System.Drawing.Color.Black
        Me.txtCartonsArrived.Location = New System.Drawing.Point(543, 87)
        Me.txtCartonsArrived.Name = "txtCartonsArrived"
        Me.txtCartonsArrived.Size = New System.Drawing.Size(135, 23)
        Me.txtCartonsArrived.TabIndex = 32
        Me.txtCartonsArrived.Tag = "20"
        Me.txtCartonsArrived.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnFLMFinishTime
        '
        Me.btnFLMFinishTime.BackColor = System.Drawing.Color.Gold
        Me.btnFLMFinishTime.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFLMFinishTime.Location = New System.Drawing.Point(543, 27)
        Me.btnFLMFinishTime.Name = "btnFLMFinishTime"
        Me.btnFLMFinishTime.Size = New System.Drawing.Size(40, 25)
        Me.btnFLMFinishTime.TabIndex = 37
        Me.btnFLMFinishTime.Tag = "btn9"
        Me.btnFLMFinishTime.Text = "@F"
        Me.btnFLMFinishTime.UseVisualStyleBackColor = False
        '
        'btnFLMStartTime
        '
        Me.btnFLMStartTime.BackColor = System.Drawing.Color.Gold
        Me.btnFLMStartTime.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFLMStartTime.Location = New System.Drawing.Point(397, 27)
        Me.btnFLMStartTime.Name = "btnFLMStartTime"
        Me.btnFLMStartTime.Size = New System.Drawing.Size(40, 25)
        Me.btnFLMStartTime.TabIndex = 35
        Me.btnFLMStartTime.Tag = "btn8"
        Me.btnFLMStartTime.Text = "@S"
        Me.btnFLMStartTime.UseVisualStyleBackColor = False
        '
        'comFLMName
        '
        Me.comFLMName.BackColor = System.Drawing.Color.AliceBlue
        Me.comFLMName.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comFLMName.ForeColor = System.Drawing.Color.Black
        Me.comFLMName.FormattingEnabled = True
        Me.comFLMName.Items.AddRange(New Object() {"ASN NUMBERS"})
        Me.comFLMName.Location = New System.Drawing.Point(162, 28)
        Me.comFLMName.Name = "comFLMName"
        Me.comFLMName.Size = New System.Drawing.Size(220, 23)
        Me.comFLMName.Sorted = True
        Me.comFLMName.TabIndex = 34
        Me.comFLMName.Tag = "40"
        '
        'txtFLMFinishTime
        '
        Me.txtFLMFinishTime.BackColor = System.Drawing.Color.LightGray
        Me.txtFLMFinishTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFLMFinishTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFLMFinishTime.ForeColor = System.Drawing.Color.Black
        Me.txtFLMFinishTime.Location = New System.Drawing.Point(598, 28)
        Me.txtFLMFinishTime.Name = "txtFLMFinishTime"
        Me.txtFLMFinishTime.Size = New System.Drawing.Size(80, 23)
        Me.txtFLMFinishTime.TabIndex = 38
        Me.txtFLMFinishTime.Tag = "442"
        Me.txtFLMFinishTime.Text = "23:59:59"
        Me.txtFLMFinishTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtFLMStartTime
        '
        Me.txtFLMStartTime.BackColor = System.Drawing.Color.LightGray
        Me.txtFLMStartTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFLMStartTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFLMStartTime.ForeColor = System.Drawing.Color.Black
        Me.txtFLMStartTime.Location = New System.Drawing.Point(452, 28)
        Me.txtFLMStartTime.Name = "txtFLMStartTime"
        Me.txtFLMStartTime.Size = New System.Drawing.Size(80, 23)
        Me.txtFLMStartTime.TabIndex = 36
        Me.txtFLMStartTime.Tag = "441"
        Me.txtFLMStartTime.Text = "00:00:00"
        Me.txtFLMStartTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEmployeeNo
        '
        Me.txtEmployeeNo.BackColor = System.Drawing.Color.LightGray
        Me.txtEmployeeNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmployeeNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmployeeNo.ForeColor = System.Drawing.Color.Black
        Me.txtEmployeeNo.Location = New System.Drawing.Point(7, 28)
        Me.txtEmployeeNo.Name = "txtEmployeeNo"
        Me.txtEmployeeNo.Size = New System.Drawing.Size(150, 23)
        Me.txtEmployeeNo.TabIndex = 33
        Me.txtEmployeeNo.Tag = "30"
        Me.txtEmployeeNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblFLMFinishTime
        '
        Me.lblFLMFinishTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFLMFinishTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFLMFinishTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFLMFinishTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblFLMFinishTime.Location = New System.Drawing.Point(543, 1)
        Me.lblFLMFinishTime.Name = "lblFLMFinishTime"
        Me.lblFLMFinishTime.Size = New System.Drawing.Size(135, 22)
        Me.lblFLMFinishTime.TabIndex = 121
        Me.lblFLMFinishTime.Text = "Finish Time"
        Me.lblFLMFinishTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'rbASNNo
        '
        Me.rbASNNo.AutoSize = True
        Me.rbASNNo.Location = New System.Drawing.Point(542, 60)
        Me.rbASNNo.Name = "rbASNNo"
        Me.rbASNNo.Size = New System.Drawing.Size(67, 17)
        Me.rbASNNo.TabIndex = 30
        Me.rbASNNo.Text = "ASN No."
        Me.rbASNNo.UseVisualStyleBackColor = True
        '
        'lblCartonsArrived
        '
        Me.lblCartonsArrived.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblCartonsArrived.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCartonsArrived.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCartonsArrived.ForeColor = System.Drawing.Color.Yellow
        Me.lblCartonsArrived.Location = New System.Drawing.Point(397, 87)
        Me.lblCartonsArrived.Name = "lblCartonsArrived"
        Me.lblCartonsArrived.Size = New System.Drawing.Size(135, 22)
        Me.lblCartonsArrived.TabIndex = 126
        Me.lblCartonsArrived.Text = "Cartons Arrived:"
        Me.lblCartonsArrived.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSelectDeliveryRefASN
        '
        Me.lblSelectDeliveryRefASN.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSelectDeliveryRefASN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSelectDeliveryRefASN.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectDeliveryRefASN.ForeColor = System.Drawing.Color.Yellow
        Me.lblSelectDeliveryRefASN.Location = New System.Drawing.Point(7, 59)
        Me.lblSelectDeliveryRefASN.Name = "lblSelectDeliveryRefASN"
        Me.lblSelectDeliveryRefASN.Size = New System.Drawing.Size(150, 22)
        Me.lblSelectDeliveryRefASN.TabIndex = 124
        Me.lblSelectDeliveryRefASN.Text = "Select Del. Ref / ASN"
        Me.lblSelectDeliveryRefASN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'imgCalendar_SelectDeliveryDate
        '
        Me.imgCalendar_SelectDeliveryDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgCalendar_SelectDeliveryDate.Image = CType(resources.GetObject("imgCalendar_SelectDeliveryDate.Image"), System.Drawing.Image)
        Me.imgCalendar_SelectDeliveryDate.Location = New System.Drawing.Point(331, 29)
        Me.imgCalendar_SelectDeliveryDate.Name = "imgCalendar_SelectDeliveryDate"
        Me.imgCalendar_SelectDeliveryDate.Size = New System.Drawing.Size(24, 28)
        Me.imgCalendar_SelectDeliveryDate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgCalendar_SelectDeliveryDate.TabIndex = 123
        Me.imgCalendar_SelectDeliveryDate.TabStop = False
        Me.imgCalendar_SelectDeliveryDate.Tag = "btn7"
        '
        'txtPalletsArrived
        '
        Me.txtPalletsArrived.BackColor = System.Drawing.Color.LightGray
        Me.txtPalletsArrived.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPalletsArrived.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPalletsArrived.ForeColor = System.Drawing.Color.Black
        Me.txtPalletsArrived.Location = New System.Drawing.Point(162, 87)
        Me.txtPalletsArrived.Name = "txtPalletsArrived"
        Me.txtPalletsArrived.Size = New System.Drawing.Size(150, 23)
        Me.txtPalletsArrived.TabIndex = 31
        Me.txtPalletsArrived.Tag = "19"
        Me.txtPalletsArrived.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblPalletsArrived
        '
        Me.lblPalletsArrived.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblPalletsArrived.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPalletsArrived.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPalletsArrived.ForeColor = System.Drawing.Color.Yellow
        Me.lblPalletsArrived.Location = New System.Drawing.Point(7, 87)
        Me.lblPalletsArrived.Name = "lblPalletsArrived"
        Me.lblPalletsArrived.Size = New System.Drawing.Size(150, 22)
        Me.lblPalletsArrived.TabIndex = 121
        Me.lblPalletsArrived.Text = "Pallets Arrived:"
        Me.lblPalletsArrived.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSelectDeliveryDate
        '
        Me.txtSelectDeliveryDate.BackColor = System.Drawing.Color.LightGray
        Me.txtSelectDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSelectDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectDeliveryDate.ForeColor = System.Drawing.Color.Black
        Me.txtSelectDeliveryDate.Location = New System.Drawing.Point(162, 33)
        Me.txtSelectDeliveryDate.Name = "txtSelectDeliveryDate"
        Me.txtSelectDeliveryDate.Size = New System.Drawing.Size(150, 23)
        Me.txtSelectDeliveryDate.TabIndex = 25
        Me.txtSelectDeliveryDate.Tag = ""
        Me.txtSelectDeliveryDate.Text = "01/01/1970"
        Me.txtSelectDeliveryDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblSelectDeliveryDate
        '
        Me.lblSelectDeliveryDate.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSelectDeliveryDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSelectDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectDeliveryDate.ForeColor = System.Drawing.Color.Yellow
        Me.lblSelectDeliveryDate.Location = New System.Drawing.Point(7, 33)
        Me.lblSelectDeliveryDate.Name = "lblSelectDeliveryDate"
        Me.lblSelectDeliveryDate.Size = New System.Drawing.Size(150, 22)
        Me.lblSelectDeliveryDate.TabIndex = 120
        Me.lblSelectDeliveryDate.Text = "Select Del. Date:"
        Me.lblSelectDeliveryDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'imgLineDivider2
        '
        Me.imgLineDivider2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider2.Image = CType(resources.GetObject("imgLineDivider2.Image"), System.Drawing.Image)
        Me.imgLineDivider2.Location = New System.Drawing.Point(-1, 109)
        Me.imgLineDivider2.Name = "imgLineDivider2"
        Me.imgLineDivider2.Size = New System.Drawing.Size(1259, 9)
        Me.imgLineDivider2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider2.TabIndex = 83
        Me.imgLineDivider2.TabStop = False
        Me.imgLineDivider2.Tag = "img6"
        '
        'Frame_SupplierCompliance
        '
        Me.Frame_SupplierCompliance.AutoScroll = True
        Me.Frame_SupplierCompliance.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_SupplierCompliance.Controls.Add(Me.PictureBox2)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnBaggingReq)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtBaggingReq)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnCollar)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtCollar)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnWrapStrap)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtWrapStrap)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnPalletise)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtPalletise)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtFurtherComments)
        Me.Frame_SupplierCompliance.Controls.Add(Me.lblTitle_FurtherComments)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtCompletedComment)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnCompleted)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnIsItSafe)
        Me.Frame_SupplierCompliance.Controls.Add(Me.btnArrivedOnTime)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtIsItSafeComment)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtIsItSafe)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtIsItCompleted)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtArrivedOnTimeComment)
        Me.Frame_SupplierCompliance.Controls.Add(Me.txtArrivedOnTime)
        Me.Frame_SupplierCompliance.ForeColor = System.Drawing.Color.Black
        Me.Frame_SupplierCompliance.Location = New System.Drawing.Point(939, 117)
        Me.Frame_SupplierCompliance.Name = "Frame_SupplierCompliance"
        Me.Frame_SupplierCompliance.Size = New System.Drawing.Size(318, 500)
        Me.Frame_SupplierCompliance.TabIndex = 104
        Me.Frame_SupplierCompliance.Tag = "frmSupplierCompliance"
        Me.Frame_SupplierCompliance.Text = "ScrollableControl7"
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(-2, 415)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(319, 9)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 116
        Me.PictureBox2.TabStop = False
        Me.PictureBox2.Tag = "img6"
        '
        'btnBaggingReq
        '
        Me.btnBaggingReq.BackColor = System.Drawing.Color.Gold
        Me.btnBaggingReq.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBaggingReq.Location = New System.Drawing.Point(14, 384)
        Me.btnBaggingReq.Name = "btnBaggingReq"
        Me.btnBaggingReq.Size = New System.Drawing.Size(145, 27)
        Me.btnBaggingReq.TabIndex = 117
        Me.btnBaggingReq.Tag = "btn167"
        Me.btnBaggingReq.Text = "Bagging Req ?"
        Me.btnBaggingReq.UseVisualStyleBackColor = False
        '
        'txtBaggingReq
        '
        Me.txtBaggingReq.BackColor = System.Drawing.Color.LightGray
        Me.txtBaggingReq.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBaggingReq.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBaggingReq.ForeColor = System.Drawing.Color.Black
        Me.txtBaggingReq.Location = New System.Drawing.Point(164, 384)
        Me.txtBaggingReq.Name = "txtBaggingReq"
        Me.txtBaggingReq.Size = New System.Drawing.Size(145, 23)
        Me.txtBaggingReq.TabIndex = 11
        Me.txtBaggingReq.Tag = "16"
        Me.txtBaggingReq.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnCollar
        '
        Me.btnCollar.BackColor = System.Drawing.Color.Gold
        Me.btnCollar.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCollar.Location = New System.Drawing.Point(14, 355)
        Me.btnCollar.Name = "btnCollar"
        Me.btnCollar.Size = New System.Drawing.Size(145, 26)
        Me.btnCollar.TabIndex = 115
        Me.btnCollar.Tag = "btn166"
        Me.btnCollar.Text = "Collar ?"
        Me.btnCollar.UseVisualStyleBackColor = False
        '
        'txtCollar
        '
        Me.txtCollar.BackColor = System.Drawing.Color.LightGray
        Me.txtCollar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCollar.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCollar.ForeColor = System.Drawing.Color.Black
        Me.txtCollar.Location = New System.Drawing.Point(164, 356)
        Me.txtCollar.Name = "txtCollar"
        Me.txtCollar.Size = New System.Drawing.Size(145, 23)
        Me.txtCollar.TabIndex = 10
        Me.txtCollar.Tag = "15"
        Me.txtCollar.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnWrapStrap
        '
        Me.btnWrapStrap.BackColor = System.Drawing.Color.Gold
        Me.btnWrapStrap.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWrapStrap.Location = New System.Drawing.Point(14, 327)
        Me.btnWrapStrap.Name = "btnWrapStrap"
        Me.btnWrapStrap.Size = New System.Drawing.Size(145, 26)
        Me.btnWrapStrap.TabIndex = 113
        Me.btnWrapStrap.Tag = "btn165"
        Me.btnWrapStrap.Text = "Wrap/Strap ?"
        Me.btnWrapStrap.UseVisualStyleBackColor = False
        '
        'txtWrapStrap
        '
        Me.txtWrapStrap.BackColor = System.Drawing.Color.LightGray
        Me.txtWrapStrap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtWrapStrap.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWrapStrap.ForeColor = System.Drawing.Color.Black
        Me.txtWrapStrap.Location = New System.Drawing.Point(164, 328)
        Me.txtWrapStrap.Name = "txtWrapStrap"
        Me.txtWrapStrap.Size = New System.Drawing.Size(145, 23)
        Me.txtWrapStrap.TabIndex = 9
        Me.txtWrapStrap.Tag = "14"
        Me.txtWrapStrap.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnPalletise
        '
        Me.btnPalletise.BackColor = System.Drawing.Color.Gold
        Me.btnPalletise.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPalletise.Location = New System.Drawing.Point(14, 299)
        Me.btnPalletise.Name = "btnPalletise"
        Me.btnPalletise.Size = New System.Drawing.Size(145, 26)
        Me.btnPalletise.TabIndex = 111
        Me.btnPalletise.Tag = "btn164"
        Me.btnPalletise.Text = "Palletise ?"
        Me.btnPalletise.UseVisualStyleBackColor = False
        '
        'txtPalletise
        '
        Me.txtPalletise.BackColor = System.Drawing.Color.LightGray
        Me.txtPalletise.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPalletise.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPalletise.ForeColor = System.Drawing.Color.Black
        Me.txtPalletise.Location = New System.Drawing.Point(164, 300)
        Me.txtPalletise.Name = "txtPalletise"
        Me.txtPalletise.Size = New System.Drawing.Size(145, 23)
        Me.txtPalletise.TabIndex = 8
        Me.txtPalletise.Tag = "13"
        Me.txtPalletise.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtFurtherComments
        '
        Me.txtFurtherComments.BackColor = System.Drawing.Color.LightGray
        Me.txtFurtherComments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFurtherComments.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFurtherComments.ForeColor = System.Drawing.Color.Black
        Me.txtFurtherComments.Location = New System.Drawing.Point(15, 215)
        Me.txtFurtherComments.Multiline = True
        Me.txtFurtherComments.Name = "txtFurtherComments"
        Me.txtFurtherComments.Size = New System.Drawing.Size(295, 79)
        Me.txtFurtherComments.TabIndex = 7
        Me.txtFurtherComments.Tag = "807"
        Me.txtFurtherComments.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblTitle_FurtherComments
        '
        Me.lblTitle_FurtherComments.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTitle_FurtherComments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTitle_FurtherComments.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle_FurtherComments.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitle_FurtherComments.Location = New System.Drawing.Point(9, 182)
        Me.lblTitle_FurtherComments.Name = "lblTitle_FurtherComments"
        Me.lblTitle_FurtherComments.Size = New System.Drawing.Size(309, 22)
        Me.lblTitle_FurtherComments.TabIndex = 108
        Me.lblTitle_FurtherComments.Text = "Further Comments"
        Me.lblTitle_FurtherComments.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtCompletedComment
        '
        Me.txtCompletedComment.BackColor = System.Drawing.Color.LightGray
        Me.txtCompletedComment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCompletedComment.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompletedComment.ForeColor = System.Drawing.Color.Black
        Me.txtCompletedComment.Location = New System.Drawing.Point(14, 141)
        Me.txtCompletedComment.Name = "txtCompletedComment"
        Me.txtCompletedComment.Size = New System.Drawing.Size(295, 23)
        Me.txtCompletedComment.TabIndex = 6
        Me.txtCompletedComment.Tag = "806"
        Me.txtCompletedComment.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnCompleted
        '
        Me.btnCompleted.BackColor = System.Drawing.Color.Gold
        Me.btnCompleted.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCompleted.Location = New System.Drawing.Point(14, 113)
        Me.btnCompleted.Name = "btnCompleted"
        Me.btnCompleted.Size = New System.Drawing.Size(145, 25)
        Me.btnCompleted.TabIndex = 104
        Me.btnCompleted.Tag = "btn5"
        Me.btnCompleted.Text = "Completed ?"
        Me.btnCompleted.UseVisualStyleBackColor = False
        '
        'btnIsItSafe
        '
        Me.btnIsItSafe.BackColor = System.Drawing.Color.Gold
        Me.btnIsItSafe.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIsItSafe.Location = New System.Drawing.Point(14, 59)
        Me.btnIsItSafe.Name = "btnIsItSafe"
        Me.btnIsItSafe.Size = New System.Drawing.Size(145, 25)
        Me.btnIsItSafe.TabIndex = 103
        Me.btnIsItSafe.Tag = "btn5"
        Me.btnIsItSafe.Text = "Is It Safe ?"
        Me.btnIsItSafe.UseVisualStyleBackColor = False
        '
        'btnArrivedOnTime
        '
        Me.btnArrivedOnTime.BackColor = System.Drawing.Color.Gold
        Me.btnArrivedOnTime.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnArrivedOnTime.Location = New System.Drawing.Point(13, 2)
        Me.btnArrivedOnTime.Name = "btnArrivedOnTime"
        Me.btnArrivedOnTime.Size = New System.Drawing.Size(145, 25)
        Me.btnArrivedOnTime.TabIndex = 102
        Me.btnArrivedOnTime.Tag = "btn5"
        Me.btnArrivedOnTime.Text = "Arrived On Time ?"
        Me.btnArrivedOnTime.UseVisualStyleBackColor = False
        '
        'txtIsItSafeComment
        '
        Me.txtIsItSafeComment.BackColor = System.Drawing.Color.LightGray
        Me.txtIsItSafeComment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtIsItSafeComment.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIsItSafeComment.ForeColor = System.Drawing.Color.Black
        Me.txtIsItSafeComment.Location = New System.Drawing.Point(14, 87)
        Me.txtIsItSafeComment.Name = "txtIsItSafeComment"
        Me.txtIsItSafeComment.Size = New System.Drawing.Size(295, 23)
        Me.txtIsItSafeComment.TabIndex = 4
        Me.txtIsItSafeComment.Tag = "804"
        Me.txtIsItSafeComment.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtIsItSafe
        '
        Me.txtIsItSafe.BackColor = System.Drawing.Color.LightGray
        Me.txtIsItSafe.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtIsItSafe.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIsItSafe.ForeColor = System.Drawing.Color.Black
        Me.txtIsItSafe.Location = New System.Drawing.Point(164, 60)
        Me.txtIsItSafe.Name = "txtIsItSafe"
        Me.txtIsItSafe.Size = New System.Drawing.Size(145, 23)
        Me.txtIsItSafe.TabIndex = 3
        Me.txtIsItSafe.Tag = "803"
        Me.txtIsItSafe.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtIsItCompleted
        '
        Me.txtIsItCompleted.BackColor = System.Drawing.Color.LightGray
        Me.txtIsItCompleted.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtIsItCompleted.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIsItCompleted.ForeColor = System.Drawing.Color.Black
        Me.txtIsItCompleted.Location = New System.Drawing.Point(164, 114)
        Me.txtIsItCompleted.Name = "txtIsItCompleted"
        Me.txtIsItCompleted.Size = New System.Drawing.Size(145, 23)
        Me.txtIsItCompleted.TabIndex = 5
        Me.txtIsItCompleted.Tag = "805"
        Me.txtIsItCompleted.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtArrivedOnTimeComment
        '
        Me.txtArrivedOnTimeComment.BackColor = System.Drawing.Color.LightGray
        Me.txtArrivedOnTimeComment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtArrivedOnTimeComment.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArrivedOnTimeComment.ForeColor = System.Drawing.Color.Black
        Me.txtArrivedOnTimeComment.Location = New System.Drawing.Point(14, 33)
        Me.txtArrivedOnTimeComment.Name = "txtArrivedOnTimeComment"
        Me.txtArrivedOnTimeComment.Size = New System.Drawing.Size(295, 23)
        Me.txtArrivedOnTimeComment.TabIndex = 2
        Me.txtArrivedOnTimeComment.Tag = "802"
        Me.txtArrivedOnTimeComment.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtArrivedOnTime
        '
        Me.txtArrivedOnTime.BackColor = System.Drawing.Color.LightGray
        Me.txtArrivedOnTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtArrivedOnTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArrivedOnTime.ForeColor = System.Drawing.Color.Black
        Me.txtArrivedOnTime.Location = New System.Drawing.Point(164, 3)
        Me.txtArrivedOnTime.Name = "txtArrivedOnTime"
        Me.txtArrivedOnTime.Size = New System.Drawing.Size(145, 23)
        Me.txtArrivedOnTime.TabIndex = 1
        Me.txtArrivedOnTime.Tag = "801"
        Me.txtArrivedOnTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'comASNNo
        '
        Me.comASNNo.BackColor = System.Drawing.Color.AliceBlue
        Me.comASNNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comASNNo.ForeColor = System.Drawing.Color.Black
        Me.comASNNo.FormattingEnabled = True
        Me.comASNNo.Items.AddRange(New Object() {"ASN NUMBERS"})
        Me.comASNNo.Location = New System.Drawing.Point(163, 59)
        Me.comASNNo.Name = "comASNNo"
        Me.comASNNo.Size = New System.Drawing.Size(150, 23)
        Me.comASNNo.TabIndex = 28
        Me.comASNNo.Tag = "32"
        '
        'lblTitle_SupplierCompliance
        '
        Me.lblTitle_SupplierCompliance.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblTitle_SupplierCompliance.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle_SupplierCompliance.ForeColor = System.Drawing.Color.Yellow
        Me.lblTitle_SupplierCompliance.Location = New System.Drawing.Point(948, 86)
        Me.lblTitle_SupplierCompliance.Name = "lblTitle_SupplierCompliance"
        Me.lblTitle_SupplierCompliance.Size = New System.Drawing.Size(310, 23)
        Me.lblTitle_SupplierCompliance.TabIndex = 86
        Me.lblTitle_SupplierCompliance.Text = "Supplier Compliance"
        Me.lblTitle_SupplierCompliance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFLMStartTime
        '
        Me.lblFLMStartTime.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFLMStartTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFLMStartTime.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFLMStartTime.ForeColor = System.Drawing.Color.Yellow
        Me.lblFLMStartTime.Location = New System.Drawing.Point(397, 1)
        Me.lblFLMStartTime.Name = "lblFLMStartTime"
        Me.lblFLMStartTime.Size = New System.Drawing.Size(135, 22)
        Me.lblFLMStartTime.TabIndex = 120
        Me.lblFLMStartTime.Text = "Start Time"
        Me.lblFLMStartTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFLMName
        '
        Me.lblFLMName.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblFLMName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFLMName.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFLMName.ForeColor = System.Drawing.Color.Yellow
        Me.lblFLMName.Location = New System.Drawing.Point(162, 1)
        Me.lblFLMName.Name = "lblFLMName"
        Me.lblFLMName.Size = New System.Drawing.Size(220, 22)
        Me.lblFLMName.TabIndex = 116
        Me.lblFLMName.Text = "FLM Name"
        Me.lblFLMName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblEmployeeNo
        '
        Me.lblEmployeeNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEmployeeNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmployeeNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmployeeNo.ForeColor = System.Drawing.Color.Yellow
        Me.lblEmployeeNo.Location = New System.Drawing.Point(7, 1)
        Me.lblEmployeeNo.Name = "lblEmployeeNo"
        Me.lblEmployeeNo.Size = New System.Drawing.Size(150, 22)
        Me.lblEmployeeNo.TabIndex = 119
        Me.lblEmployeeNo.Text = "Employee No"
        Me.lblEmployeeNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtControlValue
        '
        Me.txtControlValue.BackColor = System.Drawing.Color.SteelBlue
        Me.txtControlValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtControlValue.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtControlValue.ForeColor = System.Drawing.Color.Yellow
        Me.txtControlValue.Location = New System.Drawing.Point(162, 3)
        Me.txtControlValue.Name = "txtControlValue"
        Me.txtControlValue.Size = New System.Drawing.Size(150, 23)
        Me.txtControlValue.TabIndex = 136
        Me.txtControlValue.Tag = ""
        Me.txtControlValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLastSaved
        '
        Me.txtLastSaved.BackColor = System.Drawing.Color.Lime
        Me.txtLastSaved.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLastSaved.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastSaved.ForeColor = System.Drawing.Color.Black
        Me.txtLastSaved.Location = New System.Drawing.Point(538, 1)
        Me.txtLastSaved.Name = "txtLastSaved"
        Me.txtLastSaved.ReadOnly = True
        Me.txtLastSaved.Size = New System.Drawing.Size(145, 23)
        Me.txtLastSaved.TabIndex = 135
        Me.txtLastSaved.Tag = "22"
        Me.txtLastSaved.Text = "01/01/1970 01:00"
        Me.txtLastSaved.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblLastSaved
        '
        Me.lblLastSaved.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblLastSaved.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastSaved.ForeColor = System.Drawing.Color.Yellow
        Me.lblLastSaved.Location = New System.Drawing.Point(397, 3)
        Me.lblLastSaved.Name = "lblLastSaved"
        Me.lblLastSaved.Size = New System.Drawing.Size(135, 20)
        Me.lblLastSaved.TabIndex = 134
        Me.lblLastSaved.Text = "Last Saved:"
        Me.lblLastSaved.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtControlName
        '
        Me.txtControlName.BackColor = System.Drawing.Color.SteelBlue
        Me.txtControlName.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtControlName.ForeColor = System.Drawing.Color.Yellow
        Me.txtControlName.Location = New System.Drawing.Point(7, 3)
        Me.txtControlName.Name = "txtControlName"
        Me.txtControlName.Size = New System.Drawing.Size(150, 20)
        Me.txtControlName.TabIndex = 132
        Me.txtControlName.Text = "Last Entry:"
        Me.txtControlName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtGridNo
        '
        Me.txtGridNo.BackColor = System.Drawing.Color.LightGray
        Me.txtGridNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGridNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGridNo.ForeColor = System.Drawing.Color.Black
        Me.txtGridNo.Location = New System.Drawing.Point(543, 30)
        Me.txtGridNo.Name = "txtGridNo"
        Me.txtGridNo.Size = New System.Drawing.Size(135, 23)
        Me.txtGridNo.TabIndex = 26
        Me.txtGridNo.Tag = "21"
        Me.txtGridNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Frame_FLMDetails
        '
        Me.Frame_FLMDetails.AutoScroll = True
        Me.Frame_FLMDetails.BackColor = System.Drawing.Color.AliceBlue
        Me.Frame_FLMDetails.Controls.Add(Me.btnFLMFinishTime)
        Me.Frame_FLMDetails.Controls.Add(Me.btnFLMStartTime)
        Me.Frame_FLMDetails.Controls.Add(Me.comFLMName)
        Me.Frame_FLMDetails.Controls.Add(Me.txtFLMFinishTime)
        Me.Frame_FLMDetails.Controls.Add(Me.txtFLMStartTime)
        Me.Frame_FLMDetails.Controls.Add(Me.txtEmployeeNo)
        Me.Frame_FLMDetails.Controls.Add(Me.lblFLMFinishTime)
        Me.Frame_FLMDetails.Controls.Add(Me.lblFLMStartTime)
        Me.Frame_FLMDetails.Controls.Add(Me.lblFLMName)
        Me.Frame_FLMDetails.Controls.Add(Me.lblEmployeeNo)
        Me.Frame_FLMDetails.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame_FLMDetails.ForeColor = System.Drawing.Color.Black
        Me.Frame_FLMDetails.Location = New System.Drawing.Point(0, 115)
        Me.Frame_FLMDetails.Name = "Frame_FLMDetails"
        Me.Frame_FLMDetails.Size = New System.Drawing.Size(690, 53)
        Me.Frame_FLMDetails.TabIndex = 121
        Me.Frame_FLMDetails.Tag = "frm1"
        Me.Frame_FLMDetails.Text = "ScrollableControl3"
        '
        'imgLineDivider4
        '
        Me.imgLineDivider4.AccessibleName = "0, 112, 192"
        Me.imgLineDivider4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.imgLineDivider4.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider4.Image = CType(resources.GetObject("imgLineDivider4.Image"), System.Drawing.Image)
        Me.imgLineDivider4.Location = New System.Drawing.Point(939, 88)
        Me.imgLineDivider4.Name = "imgLineDivider4"
        Me.imgLineDivider4.Size = New System.Drawing.Size(10, 541)
        Me.imgLineDivider4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider4.TabIndex = 85
        Me.imgLineDivider4.TabStop = False
        Me.imgLineDivider4.Tag = "img4"
        '
        'imgLineDivider1
        '
        Me.imgLineDivider1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider1.Image = CType(resources.GetObject("imgLineDivider1.Image"), System.Drawing.Image)
        Me.imgLineDivider1.Location = New System.Drawing.Point(-1, 78)
        Me.imgLineDivider1.Name = "imgLineDivider1"
        Me.imgLineDivider1.Size = New System.Drawing.Size(1259, 9)
        Me.imgLineDivider1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider1.TabIndex = 80
        Me.imgLineDivider1.TabStop = False
        Me.imgLineDivider1.Tag = "img6"
        '
        'imgLineDivider6
        '
        Me.imgLineDivider6.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider6.Image = CType(resources.GetObject("imgLineDivider6.Image"), System.Drawing.Image)
        Me.imgLineDivider6.Location = New System.Drawing.Point(-1, 618)
        Me.imgLineDivider6.Name = "imgLineDivider6"
        Me.imgLineDivider6.Size = New System.Drawing.Size(1258, 9)
        Me.imgLineDivider6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider6.TabIndex = 90
        Me.imgLineDivider6.TabStop = False
        Me.imgLineDivider6.Tag = "img6"
        '
        'imgLineDivider5
        '
        Me.imgLineDivider5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.imgLineDivider5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgLineDivider5.Image = CType(resources.GetObject("imgLineDivider5.Image"), System.Drawing.Image)
        Me.imgLineDivider5.Location = New System.Drawing.Point(-1, 286)
        Me.imgLineDivider5.Name = "imgLineDivider5"
        Me.imgLineDivider5.Size = New System.Drawing.Size(1259, 9)
        Me.imgLineDivider5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgLineDivider5.TabIndex = 87
        Me.imgLineDivider5.TabStop = False
        Me.imgLineDivider5.Tag = "img6"
        '
        'Frame_OperationalInput
        '
        Me.Frame_OperationalInput.AutoScroll = True
        Me.Frame_OperationalInput.BackColor = System.Drawing.Color.Azure
        Me.Frame_OperationalInput.Controls.Add(Me.txtControlValue)
        Me.Frame_OperationalInput.Controls.Add(Me.txtLastSaved)
        Me.Frame_OperationalInput.Controls.Add(Me.lblLastSaved)
        Me.Frame_OperationalInput.Controls.Add(Me.txtControlName)
        Me.Frame_OperationalInput.Controls.Add(Me.txtGridNo)
        Me.Frame_OperationalInput.Controls.Add(Me.Frame_FLMDetails)
        Me.Frame_OperationalInput.Controls.Add(Me.comDeliveryRef)
        Me.Frame_OperationalInput.Controls.Add(Me.lblGridNo)
        Me.Frame_OperationalInput.Controls.Add(Me.rbASNNo)
        Me.Frame_OperationalInput.Controls.Add(Me.rbDeliveryRef)
        Me.Frame_OperationalInput.Controls.Add(Me.txtCartonsArrived)
        Me.Frame_OperationalInput.Controls.Add(Me.lblCartonsArrived)
        Me.Frame_OperationalInput.Controls.Add(Me.comASNNo)
        Me.Frame_OperationalInput.Controls.Add(Me.lblSelectDeliveryRefASN)
        Me.Frame_OperationalInput.Controls.Add(Me.imgCalendar_SelectDeliveryDate)
        Me.Frame_OperationalInput.Controls.Add(Me.txtPalletsArrived)
        Me.Frame_OperationalInput.Controls.Add(Me.lblPalletsArrived)
        Me.Frame_OperationalInput.Controls.Add(Me.txtSelectDeliveryDate)
        Me.Frame_OperationalInput.Controls.Add(Me.lblSelectDeliveryDate)
        Me.Frame_OperationalInput.Location = New System.Drawing.Point(250, 117)
        Me.Frame_OperationalInput.Name = "Frame_OperationalInput"
        Me.Frame_OperationalInput.Size = New System.Drawing.Size(690, 168)
        Me.Frame_OperationalInput.TabIndex = 115
        Me.Frame_OperationalInput.Tag = "frmOpInput"
        Me.Frame_OperationalInput.Text = "ScrollableControl3"
        '
        'comDeliveryRef
        '
        Me.comDeliveryRef.BackColor = System.Drawing.Color.White
        Me.comDeliveryRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comDeliveryRef.ForeColor = System.Drawing.Color.Black
        Me.comDeliveryRef.FormattingEnabled = True
        Me.comDeliveryRef.Location = New System.Drawing.Point(162, 59)
        Me.comDeliveryRef.Name = "comDeliveryRef"
        Me.comDeliveryRef.Size = New System.Drawing.Size(150, 23)
        Me.comDeliveryRef.TabIndex = 27
        Me.comDeliveryRef.Tag = ""
        '
        'pnlMain
        '
        Me.pnlMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMain.Controls.Add(Me.imgLineDivider4)
        Me.pnlMain.Controls.Add(Me.imgLineDivider1)
        Me.pnlMain.Controls.Add(Me.imgLineDivider6)
        Me.pnlMain.Controls.Add(Me.imgLineDivider5)
        Me.pnlMain.Controls.Add(Me.Frame_OperationalInput)
        Me.pnlMain.Controls.Add(Me.imgLineDivider2)
        Me.pnlMain.Controls.Add(Me.Frame_SupplierCompliance)
        Me.pnlMain.Controls.Add(Me.lblTitle_SupplierCompliance)
        Me.pnlMain.Controls.Add(Me.Frame_InboundSchedule)
        Me.pnlMain.Controls.Add(Me.imgLineDivider3)
        Me.pnlMain.Controls.Add(Me.lblTitle_OperationalInput)
        Me.pnlMain.Controls.Add(Me.lblTitle_InboundSchedule)
        Me.pnlMain.Controls.Add(Me.Frame_OpsShortsAndExtras)
        Me.pnlMain.Location = New System.Drawing.Point(4, 2)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(1271, 654)
        Me.pnlMain.TabIndex = 2
        '
        'frm_GI_GRID
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1285, 662)
        Me.Controls.Add(Me.pnlTOP)
        Me.Controls.Add(Me.pnlMain)
        Me.Name = "frm_GI_GRID"
        Me.Text = "Form_GI_GRID_Userform"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_Extra_Parts.ResumeLayout(False)
        CType(Me.dgvExtras, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_InboundSchedule.ResumeLayout(False)
        Me.Frame_InboundSchedule.PerformLayout()
        CType(Me.imgLineDivider3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_OpsShortsAndExtras.ResumeLayout(False)
        Me.Frame_OpsShortsAndExtras.PerformLayout()
        Me.Frame_Operatives.ResumeLayout(False)
        CType(Me.dgvOperatives, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_Short_Parts.ResumeLayout(False)
        CType(Me.dgvShorts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPCarLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLineDivider7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECP_ThumbUpLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgECPLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgCalendar_SelectImportDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_Buttons.ResumeLayout(False)
        Me.pnlTOP.ResumeLayout(False)
        CType(Me.imgCalendar_SelectDeliveryDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLineDivider2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_SupplierCompliance.ResumeLayout(False)
        Me.Frame_SupplierCompliance.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_FLMDetails.ResumeLayout(False)
        Me.Frame_FLMDetails.PerformLayout()
        CType(Me.imgLineDivider4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLineDivider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLineDivider6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgLineDivider5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame_OperationalInput.ResumeLayout(False)
        Me.Frame_OperationalInput.PerformLayout()
        Me.pnlMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents txtTotalExtras As TextBox
    Friend WithEvents txtTotalShorts As TextBox
    Friend WithEvents lblOpComment As Label
    Friend WithEvents txtTotalHours As TextBox
    Friend WithEvents lblTotalHours As Label
    Friend WithEvents lblOpHash As Label
    Friend WithEvents txtTotalOps As TextBox
    Friend WithEvents lblTotalOps As Label
    Friend WithEvents btnDeleteExtra As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnAddExtra As Button
    Friend WithEvents btnDeleteShort As Button
    Friend WithEvents Frame_Extra_Parts As ScrollableControl
    Friend WithEvents Frame_InboundSchedule As ScrollableControl
    Friend WithEvents txtSupplier As TextBox
    Friend WithEvents txtCalculatedHours As TextBox
    Friend WithEvents txtEstimatedTotes As TextBox
    Friend WithEvents txtEstimatedCages As TextBox
    Friend WithEvents txtEstimatedPallets As TextBox
    Friend WithEvents txtReadyLabel As TextBox
    Friend WithEvents lblReadyLabel As Label
    Friend WithEvents txtActualCases As TextBox
    Friend WithEvents lblActualCases As Label
    Friend WithEvents txtCartonsDue As TextBox
    Friend WithEvents lblCartonsDue As Label
    Friend WithEvents txtPalletsDue As TextBox
    Friend WithEvents lblPalletsDue As Label
    Friend WithEvents txtExpectedLines As TextBox
    Friend WithEvents lblExpectedLines As Label
    Friend WithEvents lblCalculatedHours As Label
    Friend WithEvents lblEstimatedTotes As Label
    Friend WithEvents lblEstimatedCages As Label
    Friend WithEvents lblEstimatedPallets As Label
    Friend WithEvents txtExpectedCases As TextBox
    Friend WithEvents lblExpectedCases As Label
    Friend WithEvents txtOrigin As TextBox
    Friend WithEvents lblOrigin As Label
    Friend WithEvents txtShift As TextBox
    Friend WithEvents lblShift As Label
    Friend WithEvents txtDueTime As TextBox
    Friend WithEvents lblDueTime As Label
    Friend WithEvents txtASNnum As TextBox
    Friend WithEvents lblASNno As Label
    Friend WithEvents lblSupplier As Label
    Friend WithEvents txtDeliveryRef As TextBox
    Friend WithEvents txtDeliveryDate As TextBox
    Friend WithEvents lblDeliveryRef As Label
    Friend WithEvents lblDeliveryDate As Label
    Friend WithEvents imgLineDivider3 As PictureBox
    Friend WithEvents Frame_OpsShortsAndExtras As ScrollableControl
    Friend WithEvents btnAddShort As Button
    Friend WithEvents Frame_Short_Parts As ScrollableControl
    Friend WithEvents ZoomControl As HScrollBar
    Friend WithEvents btnDeleteOperative As Button
    Friend WithEvents btnAddOperative As Button
    Friend WithEvents lblOpFinishTime As Label
    Friend WithEvents lblOpStartTime As Label
    Friend WithEvents lblOpActivity As Label
    Friend WithEvents lblOpName As Label
    Friend WithEvents imgECPCarLogo As PictureBox
    Friend WithEvents lblTitle_OperationalInput As Label
    Friend WithEvents lblTitle_InboundSchedule As Label
    Friend WithEvents btnUpdateEmployees As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents btnCreateTimesheet As Button
    Friend WithEvents btnLoadDeliveryRef As Button
    Friend WithEvents btnSaveAndContinue As Button
    Friend WithEvents imgLineDivider7 As PictureBox
    Friend WithEvents imgECP_ThumbUpLogo As PictureBox
    Friend WithEvents btnReset As Button
    Friend WithEvents imgECPLogo As PictureBox
    Friend WithEvents imgCalendar_SelectImportDate As PictureBox
    Friend WithEvents Frame_Buttons As ScrollableControl
    Friend WithEvents btnImportData As Button
    Friend WithEvents pnlTOP As Panel
    Friend WithEvents lblGridNo As Label
    Friend WithEvents rbDeliveryRef As RadioButton
    Friend WithEvents txtCartonsArrived As TextBox
    Friend WithEvents btnFLMFinishTime As Button
    Friend WithEvents btnFLMStartTime As Button
    Friend WithEvents comFLMName As ComboBox
    Friend WithEvents txtFLMFinishTime As TextBox
    Friend WithEvents txtFLMStartTime As TextBox
    Friend WithEvents txtEmployeeNo As TextBox
    Friend WithEvents lblFLMFinishTime As Label
    Friend WithEvents rbASNNo As RadioButton
    Friend WithEvents lblCartonsArrived As Label
    Friend WithEvents lblSelectDeliveryRefASN As Label
    Friend WithEvents imgCalendar_SelectDeliveryDate As PictureBox
    Friend WithEvents txtPalletsArrived As TextBox
    Friend WithEvents lblPalletsArrived As Label
    Friend WithEvents txtSelectDeliveryDate As TextBox
    Friend WithEvents lblSelectDeliveryDate As Label
    Friend WithEvents imgLineDivider2 As PictureBox
    Friend WithEvents Frame_SupplierCompliance As ScrollableControl
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents btnBaggingReq As Button
    Friend WithEvents txtBaggingReq As TextBox
    Friend WithEvents btnCollar As Button
    Friend WithEvents txtCollar As TextBox
    Friend WithEvents btnWrapStrap As Button
    Friend WithEvents txtWrapStrap As TextBox
    Friend WithEvents btnPalletise As Button
    Friend WithEvents txtPalletise As TextBox
    Friend WithEvents txtFurtherComments As TextBox
    Friend WithEvents lblTitle_FurtherComments As Label
    Friend WithEvents txtCompletedComment As TextBox
    Friend WithEvents btnCompleted As Button
    Friend WithEvents btnIsItSafe As Button
    Friend WithEvents btnArrivedOnTime As Button
    Friend WithEvents txtIsItSafeComment As TextBox
    Friend WithEvents txtIsItSafe As TextBox
    Friend WithEvents txtIsItCompleted As TextBox
    Friend WithEvents txtArrivedOnTimeComment As TextBox
    Friend WithEvents txtArrivedOnTime As TextBox
    Friend WithEvents comASNNo As ComboBox
    Friend WithEvents lblTitle_SupplierCompliance As Label
    Friend WithEvents lblFLMStartTime As Label
    Friend WithEvents lblFLMName As Label
    Friend WithEvents lblEmployeeNo As Label
    Friend WithEvents txtControlValue As TextBox
    Friend WithEvents txtLastSaved As TextBox
    Friend WithEvents lblLastSaved As Label
    Friend WithEvents txtControlName As Label
    Friend WithEvents txtGridNo As TextBox
    Friend WithEvents Frame_FLMDetails As ScrollableControl
    Friend WithEvents imgLineDivider4 As PictureBox
    Friend WithEvents imgLineDivider1 As PictureBox
    Friend WithEvents imgLineDivider6 As PictureBox
    Friend WithEvents imgLineDivider5 As PictureBox
    Friend WithEvents Frame_OperationalInput As ScrollableControl
    Friend WithEvents comDeliveryRef As ComboBox
    Friend WithEvents pnlMain As Panel
    Friend WithEvents dgvExtras As DataGridView
    Friend WithEvents Frame_Operatives As ScrollableControl
    Friend WithEvents dgvOperatives As DataGridView
    Friend WithEvents dgvShorts As DataGridView
End Class
