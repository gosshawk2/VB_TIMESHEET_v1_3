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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReferenceProgress))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.pnlTitle = New System.Windows.Forms.Panel()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.pnlbuttons = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSearchRef = New System.Windows.Forms.TextBox()
        Me.btnSearchEmpNo = New System.Windows.Forms.Button()
        Me.lblSearchASN = New System.Windows.Forms.Label()
        Me.txtSearchASN = New System.Windows.Forms.TextBox()
        Me.btnSavePreferences = New System.Windows.Forms.Button()
        Me.btnExportToCSV = New System.Windows.Forms.Button()
        Me.dgvDeliveryProgress = New System.Windows.Forms.DataGridView()
        Me.STATUSMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.NOSHOWToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CANCELLEDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ROLLEDOVERToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer_UpdateProgressGrid = New System.Windows.Forms.Timer(Me.components)
        Me.txtNumRows = New System.Windows.Forms.TextBox()
        Me.pnlControls = New System.Windows.Forms.Panel()
        Me.lblSetNewStatus = New System.Windows.Forms.Label()
        Me.lblID = New System.Windows.Forms.Label()
        Me.btnSetNewStatus = New System.Windows.Forms.Button()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.comSTATUS = New System.Windows.Forms.ComboBox()
        Me.txtNewStatus = New System.Windows.Forms.TextBox()
        Me.lblCurrentStatus = New System.Windows.Forms.Label()
        Me.txtCurrentStatus = New System.Windows.Forms.TextBox()
        Me.lblLastSaved = New System.Windows.Forms.Label()
        Me.txtLastSaved = New System.Windows.Forms.TextBox()
        Me.lblGrid = New System.Windows.Forms.Label()
        Me.txtGrid = New System.Windows.Forms.TextBox()
        Me.lblSupplier = New System.Windows.Forms.Label()
        Me.txtSupplier = New System.Windows.Forms.TextBox()
        Me.lblReference = New System.Windows.Forms.Label()
        Me.txtDeliveryReference = New System.Windows.Forms.TextBox()
        Me.lblDeliveryDate = New System.Windows.Forms.Label()
        Me.txtDeliveryDate = New System.Windows.Forms.TextBox()
        Me.gbProgressFilter = New System.Windows.Forms.GroupBox()
        Me.rbCancelled = New System.Windows.Forms.RadioButton()
        Me.rbNoShow = New System.Windows.Forms.RadioButton()
        Me.rbRolledOver = New System.Windows.Forms.RadioButton()
        Me.rbCompleted = New System.Windows.Forms.RadioButton()
        Me.rbInProgress = New System.Windows.Forms.RadioButton()
        Me.rbRolled = New System.Windows.Forms.RadioButton()
        Me.rbALL = New System.Windows.Forms.RadioButton()
        Me.lblTotalRows = New System.Windows.Forms.Label()
        Me.btnTodayOnly = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlTopButtons = New System.Windows.Forms.Panel()
        Me.btnImportData = New System.Windows.Forms.Button()
        Me.imglistImportData = New System.Windows.Forms.ImageList(Me.components)
        Me.pnlTitle.SuspendLayout()
        Me.pnlbuttons.SuspendLayout()
        CType(Me.dgvDeliveryProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.STATUSMenuStrip.SuspendLayout()
        Me.pnlControls.SuspendLayout()
        Me.gbProgressFilter.SuspendLayout()
        Me.pnlTopButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTitle
        '
        Me.pnlTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTitle.BackColor = System.Drawing.Color.Blue
        Me.pnlTitle.Controls.Add(Me.txtTitle)
        Me.pnlTitle.ForeColor = System.Drawing.Color.Black
        Me.pnlTitle.Location = New System.Drawing.Point(3, -4)
        Me.pnlTitle.Name = "pnlTitle"
        Me.pnlTitle.Size = New System.Drawing.Size(1063, 48)
        Me.pnlTitle.TabIndex = 41
        '
        'txtTitle
        '
        Me.txtTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTitle.BackColor = System.Drawing.Color.DarkBlue
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.AliceBlue
        Me.txtTitle.Location = New System.Drawing.Point(3, 6)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(1055, 39)
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
        Me.pnlbuttons.Controls.Add(Me.btnSavePreferences)
        Me.pnlbuttons.Location = New System.Drawing.Point(3, 235)
        Me.pnlbuttons.Name = "pnlbuttons"
        Me.pnlbuttons.Size = New System.Drawing.Size(1058, 50)
        Me.pnlbuttons.TabIndex = 42
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.Label1.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(435, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 19)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Search Reference:"
        '
        'txtSearchRef
        '
        Me.txtSearchRef.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchRef.Location = New System.Drawing.Point(585, 15)
        Me.txtSearchRef.Name = "txtSearchRef"
        Me.txtSearchRef.Size = New System.Drawing.Size(125, 23)
        Me.txtSearchRef.TabIndex = 8
        '
        'btnSearchEmpNo
        '
        Me.btnSearchEmpNo.BackColor = System.Drawing.Color.SpringGreen
        Me.btnSearchEmpNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchEmpNo.Location = New System.Drawing.Point(719, 10)
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
        Me.lblSearchASN.Location = New System.Drawing.Point(168, 15)
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
        'btnSavePreferences
        '
        Me.btnSavePreferences.BackColor = System.Drawing.Color.SpringGreen
        Me.btnSavePreferences.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSavePreferences.Location = New System.Drawing.Point(18, 10)
        Me.btnSavePreferences.Name = "btnSavePreferences"
        Me.btnSavePreferences.Size = New System.Drawing.Size(140, 30)
        Me.btnSavePreferences.TabIndex = 1
        Me.btnSavePreferences.Text = "Save Preferences"
        Me.btnSavePreferences.UseVisualStyleBackColor = False
        Me.btnSavePreferences.Visible = False
        '
        'btnExportToCSV
        '
        Me.btnExportToCSV.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportToCSV.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnExportToCSV.BackColor = System.Drawing.Color.Transparent
        Me.btnExportToCSV.BackgroundImage = CType(resources.GetObject("btnExportToCSV.BackgroundImage"), System.Drawing.Image)
        Me.btnExportToCSV.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnExportToCSV.FlatAppearance.BorderSize = 0
        Me.btnExportToCSV.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExportToCSV.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportToCSV.Location = New System.Drawing.Point(779, 8)
        Me.btnExportToCSV.Name = "btnExportToCSV"
        Me.btnExportToCSV.Size = New System.Drawing.Size(166, 30)
        Me.btnExportToCSV.TabIndex = 3
        Me.btnExportToCSV.UseVisualStyleBackColor = False
        '
        'dgvDeliveryProgress
        '
        Me.dgvDeliveryProgress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDeliveryProgress.BackgroundColor = System.Drawing.Color.LightGray
        Me.dgvDeliveryProgress.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDeliveryProgress.ContextMenuStrip = Me.STATUSMenuStrip
        Me.dgvDeliveryProgress.GridColor = System.Drawing.Color.Silver
        Me.dgvDeliveryProgress.Location = New System.Drawing.Point(3, 288)
        Me.dgvDeliveryProgress.Name = "dgvDeliveryProgress"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.NullValue = Nothing
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Cyan
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDeliveryProgress.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.NullValue = Nothing
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Cyan
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.dgvDeliveryProgress.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDeliveryProgress.Size = New System.Drawing.Size(1058, 106)
        Me.dgvDeliveryProgress.TabIndex = 43
        '
        'STATUSMenuStrip
        '
        Me.STATUSMenuStrip.BackColor = System.Drawing.Color.Silver
        Me.STATUSMenuStrip.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.STATUSMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NOSHOWToolStripMenuItem, Me.CANCELLEDToolStripMenuItem, Me.ROLLEDOVERToolStripMenuItem1})
        Me.STATUSMenuStrip.Name = "STATUSMenuStrip"
        Me.STATUSMenuStrip.Size = New System.Drawing.Size(261, 76)
        '
        'NOSHOWToolStripMenuItem
        '
        Me.NOSHOWToolStripMenuItem.BackColor = System.Drawing.Color.MediumPurple
        Me.NOSHOWToolStripMenuItem.Name = "NOSHOWToolStripMenuItem"
        Me.NOSHOWToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.NOSHOWToolStripMenuItem.Size = New System.Drawing.Size(260, 24)
        Me.NOSHOWToolStripMenuItem.Text = "NO SHOW"
        '
        'CANCELLEDToolStripMenuItem
        '
        Me.CANCELLEDToolStripMenuItem.BackColor = System.Drawing.Color.Red
        Me.CANCELLEDToolStripMenuItem.Name = "CANCELLEDToolStripMenuItem"
        Me.CANCELLEDToolStripMenuItem.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CANCELLEDToolStripMenuItem.Size = New System.Drawing.Size(260, 24)
        Me.CANCELLEDToolStripMenuItem.Text = "CANCELLED"
        '
        'ROLLEDOVERToolStripMenuItem1
        '
        Me.ROLLEDOVERToolStripMenuItem1.BackColor = System.Drawing.Color.Firebrick
        Me.ROLLEDOVERToolStripMenuItem1.ForeColor = System.Drawing.Color.White
        Me.ROLLEDOVERToolStripMenuItem1.Name = "ROLLEDOVERToolStripMenuItem1"
        Me.ROLLEDOVERToolStripMenuItem1.Size = New System.Drawing.Size(260, 24)
        Me.ROLLEDOVERToolStripMenuItem1.Text = "ROLLED OVER"
        '
        'Timer_UpdateProgressGrid
        '
        Me.Timer_UpdateProgressGrid.Enabled = True
        Me.Timer_UpdateProgressGrid.Interval = 10000
        '
        'txtNumRows
        '
        Me.txtNumRows.BackColor = System.Drawing.Color.Black
        Me.txtNumRows.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumRows.ForeColor = System.Drawing.Color.Red
        Me.txtNumRows.Location = New System.Drawing.Point(818, 12)
        Me.txtNumRows.Name = "txtNumRows"
        Me.txtNumRows.Size = New System.Drawing.Size(86, 26)
        Me.txtNumRows.TabIndex = 8
        Me.txtNumRows.Text = "9999999"
        Me.txtNumRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlControls
        '
        Me.pnlControls.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlControls.BackColor = System.Drawing.Color.LightSteelBlue
        Me.pnlControls.Controls.Add(Me.lblSetNewStatus)
        Me.pnlControls.Controls.Add(Me.lblID)
        Me.pnlControls.Controls.Add(Me.btnSetNewStatus)
        Me.pnlControls.Controls.Add(Me.txtID)
        Me.pnlControls.Controls.Add(Me.comSTATUS)
        Me.pnlControls.Controls.Add(Me.txtNewStatus)
        Me.pnlControls.Controls.Add(Me.lblCurrentStatus)
        Me.pnlControls.Controls.Add(Me.txtCurrentStatus)
        Me.pnlControls.Controls.Add(Me.lblLastSaved)
        Me.pnlControls.Controls.Add(Me.txtLastSaved)
        Me.pnlControls.Controls.Add(Me.lblGrid)
        Me.pnlControls.Controls.Add(Me.txtGrid)
        Me.pnlControls.Controls.Add(Me.lblSupplier)
        Me.pnlControls.Controls.Add(Me.txtSupplier)
        Me.pnlControls.Controls.Add(Me.lblReference)
        Me.pnlControls.Controls.Add(Me.txtDeliveryReference)
        Me.pnlControls.Controls.Add(Me.lblDeliveryDate)
        Me.pnlControls.Controls.Add(Me.txtDeliveryDate)
        Me.pnlControls.Controls.Add(Me.gbProgressFilter)
        Me.pnlControls.Controls.Add(Me.lblTotalRows)
        Me.pnlControls.Controls.Add(Me.txtNumRows)
        Me.pnlControls.ForeColor = System.Drawing.Color.Black
        Me.pnlControls.Location = New System.Drawing.Point(6, 98)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(1055, 135)
        Me.pnlControls.TabIndex = 44
        '
        'lblSetNewStatus
        '
        Me.lblSetNewStatus.AutoSize = True
        Me.lblSetNewStatus.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblSetNewStatus.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSetNewStatus.Location = New System.Drawing.Point(435, 101)
        Me.lblSetNewStatus.Name = "lblSetNewStatus"
        Me.lblSetNewStatus.Size = New System.Drawing.Size(121, 19)
        Me.lblSetNewStatus.TabIndex = 59
        Me.lblSetNewStatus.Text = "Set New Status:"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblID.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblID.Location = New System.Drawing.Point(16, 99)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(30, 19)
        Me.lblID.TabIndex = 58
        Me.lblID.Text = "ID:"
        '
        'btnSetNewStatus
        '
        Me.btnSetNewStatus.BackColor = System.Drawing.Color.Transparent
        Me.btnSetNewStatus.BackgroundImage = CType(resources.GetObject("btnSetNewStatus.BackgroundImage"), System.Drawing.Image)
        Me.btnSetNewStatus.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSetNewStatus.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSetNewStatus.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSetNewStatus.Location = New System.Drawing.Point(817, 94)
        Me.btnSetNewStatus.Name = "btnSetNewStatus"
        Me.btnSetNewStatus.Size = New System.Drawing.Size(125, 34)
        Me.btnSetNewStatus.TabIndex = 55
        Me.btnSetNewStatus.UseVisualStyleBackColor = False
        '
        'txtID
        '
        Me.txtID.BackColor = System.Drawing.Color.AliceBlue
        Me.txtID.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(58, 97)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(61, 23)
        Me.txtID.TabIndex = 57
        Me.txtID.Text = "0"
        '
        'comSTATUS
        '
        Me.comSTATUS.BackColor = System.Drawing.Color.AliceBlue
        Me.comSTATUS.FormattingEnabled = True
        Me.comSTATUS.Location = New System.Drawing.Point(693, 100)
        Me.comSTATUS.MinimumSize = New System.Drawing.Size(60, 0)
        Me.comSTATUS.Name = "comSTATUS"
        Me.comSTATUS.Size = New System.Drawing.Size(108, 21)
        Me.comSTATUS.TabIndex = 56
        '
        'txtNewStatus
        '
        Me.txtNewStatus.BackColor = System.Drawing.Color.AliceBlue
        Me.txtNewStatus.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewStatus.Location = New System.Drawing.Point(562, 98)
        Me.txtNewStatus.Name = "txtNewStatus"
        Me.txtNewStatus.Size = New System.Drawing.Size(125, 26)
        Me.txtNewStatus.TabIndex = 54
        '
        'lblCurrentStatus
        '
        Me.lblCurrentStatus.AutoSize = True
        Me.lblCurrentStatus.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblCurrentStatus.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentStatus.Location = New System.Drawing.Point(132, 100)
        Me.lblCurrentStatus.Name = "lblCurrentStatus"
        Me.lblCurrentStatus.Size = New System.Drawing.Size(120, 19)
        Me.lblCurrentStatus.TabIndex = 53
        Me.lblCurrentStatus.Text = "Current Status:"
        '
        'txtCurrentStatus
        '
        Me.txtCurrentStatus.BackColor = System.Drawing.Color.WhiteSmoke
        Me.txtCurrentStatus.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurrentStatus.ForeColor = System.Drawing.Color.Black
        Me.txtCurrentStatus.Location = New System.Drawing.Point(260, 98)
        Me.txtCurrentStatus.Name = "txtCurrentStatus"
        Me.txtCurrentStatus.Size = New System.Drawing.Size(168, 26)
        Me.txtCurrentStatus.TabIndex = 52
        Me.txtCurrentStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblLastSaved
        '
        Me.lblLastSaved.AutoSize = True
        Me.lblLastSaved.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblLastSaved.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastSaved.Location = New System.Drawing.Point(818, 60)
        Me.lblLastSaved.Name = "lblLastSaved"
        Me.lblLastSaved.Size = New System.Drawing.Size(90, 19)
        Me.lblLastSaved.TabIndex = 23
        Me.lblLastSaved.Text = "Last Saved:"
        '
        'txtLastSaved
        '
        Me.txtLastSaved.BackColor = System.Drawing.Color.AliceBlue
        Me.txtLastSaved.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastSaved.Location = New System.Drawing.Point(910, 58)
        Me.txtLastSaved.Name = "txtLastSaved"
        Me.txtLastSaved.Size = New System.Drawing.Size(140, 23)
        Me.txtLastSaved.TabIndex = 22
        '
        'lblGrid
        '
        Me.lblGrid.AutoSize = True
        Me.lblGrid.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblGrid.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGrid.Location = New System.Drawing.Point(661, 60)
        Me.lblGrid.Name = "lblGrid"
        Me.lblGrid.Size = New System.Drawing.Size(45, 19)
        Me.lblGrid.TabIndex = 21
        Me.lblGrid.Text = "Grid:"
        '
        'txtGrid
        '
        Me.txtGrid.BackColor = System.Drawing.Color.AliceBlue
        Me.txtGrid.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGrid.Location = New System.Drawing.Point(715, 58)
        Me.txtGrid.Name = "txtGrid"
        Me.txtGrid.Size = New System.Drawing.Size(86, 23)
        Me.txtGrid.TabIndex = 20
        '
        'lblSupplier
        '
        Me.lblSupplier.AutoSize = True
        Me.lblSupplier.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblSupplier.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSupplier.Location = New System.Drawing.Point(435, 60)
        Me.lblSupplier.Name = "lblSupplier"
        Me.lblSupplier.Size = New System.Drawing.Size(76, 19)
        Me.lblSupplier.TabIndex = 19
        Me.lblSupplier.Text = "Supplier:"
        '
        'txtSupplier
        '
        Me.txtSupplier.BackColor = System.Drawing.Color.AliceBlue
        Me.txtSupplier.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplier.Location = New System.Drawing.Point(519, 58)
        Me.txtSupplier.Name = "txtSupplier"
        Me.txtSupplier.Size = New System.Drawing.Size(125, 20)
        Me.txtSupplier.TabIndex = 18
        '
        'lblReference
        '
        Me.lblReference.AutoSize = True
        Me.lblReference.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblReference.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReference.Location = New System.Drawing.Point(261, 60)
        Me.lblReference.Name = "lblReference"
        Me.lblReference.Size = New System.Drawing.Size(37, 19)
        Me.lblReference.TabIndex = 17
        Me.lblReference.Text = "Ref:"
        '
        'txtDeliveryReference
        '
        Me.txtDeliveryReference.BackColor = System.Drawing.Color.AliceBlue
        Me.txtDeliveryReference.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryReference.Location = New System.Drawing.Point(303, 58)
        Me.txtDeliveryReference.Name = "txtDeliveryReference"
        Me.txtDeliveryReference.Size = New System.Drawing.Size(125, 23)
        Me.txtDeliveryReference.TabIndex = 16
        '
        'lblDeliveryDate
        '
        Me.lblDeliveryDate.AutoSize = True
        Me.lblDeliveryDate.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblDeliveryDate.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeliveryDate.Location = New System.Drawing.Point(15, 60)
        Me.lblDeliveryDate.Name = "lblDeliveryDate"
        Me.lblDeliveryDate.Size = New System.Drawing.Size(111, 19)
        Me.lblDeliveryDate.TabIndex = 15
        Me.lblDeliveryDate.Text = "Delivery Date:"
        '
        'txtDeliveryDate
        '
        Me.txtDeliveryDate.BackColor = System.Drawing.Color.AliceBlue
        Me.txtDeliveryDate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeliveryDate.Location = New System.Drawing.Point(131, 58)
        Me.txtDeliveryDate.Name = "txtDeliveryDate"
        Me.txtDeliveryDate.Size = New System.Drawing.Size(125, 23)
        Me.txtDeliveryDate.TabIndex = 14
        '
        'gbProgressFilter
        '
        Me.gbProgressFilter.BackColor = System.Drawing.Color.AliceBlue
        Me.gbProgressFilter.Controls.Add(Me.rbCancelled)
        Me.gbProgressFilter.Controls.Add(Me.rbNoShow)
        Me.gbProgressFilter.Controls.Add(Me.rbRolledOver)
        Me.gbProgressFilter.Controls.Add(Me.rbCompleted)
        Me.gbProgressFilter.Controls.Add(Me.rbInProgress)
        Me.gbProgressFilter.Controls.Add(Me.rbRolled)
        Me.gbProgressFilter.Controls.Add(Me.rbALL)
        Me.gbProgressFilter.ForeColor = System.Drawing.Color.White
        Me.gbProgressFilter.Location = New System.Drawing.Point(1, -2)
        Me.gbProgressFilter.Name = "gbProgressFilter"
        Me.gbProgressFilter.Size = New System.Drawing.Size(800, 48)
        Me.gbProgressFilter.TabIndex = 10
        Me.gbProgressFilter.TabStop = False
        Me.gbProgressFilter.Text = "Status Filter"
        '
        'rbCancelled
        '
        Me.rbCancelled.AutoSize = True
        Me.rbCancelled.BackColor = System.Drawing.Color.Red
        Me.rbCancelled.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCancelled.ForeColor = System.Drawing.Color.Black
        Me.rbCancelled.Location = New System.Drawing.Point(440, 15)
        Me.rbCancelled.Name = "rbCancelled"
        Me.rbCancelled.Size = New System.Drawing.Size(105, 21)
        Me.rbCancelled.TabIndex = 6
        Me.rbCancelled.Text = "CANCELLED"
        Me.rbCancelled.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbCancelled.UseVisualStyleBackColor = False
        '
        'rbNoShow
        '
        Me.rbNoShow.AutoSize = True
        Me.rbNoShow.BackColor = System.Drawing.Color.Firebrick
        Me.rbNoShow.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNoShow.ForeColor = System.Drawing.Color.White
        Me.rbNoShow.Location = New System.Drawing.Point(325, 15)
        Me.rbNoShow.Name = "rbNoShow"
        Me.rbNoShow.Size = New System.Drawing.Size(105, 21)
        Me.rbNoShow.TabIndex = 5
        Me.rbNoShow.Text = "NO SHOW     "
        Me.rbNoShow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbNoShow.UseVisualStyleBackColor = False
        '
        'rbRolledOver
        '
        Me.rbRolledOver.AutoSize = True
        Me.rbRolledOver.BackColor = System.Drawing.Color.Maroon
        Me.rbRolledOver.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRolledOver.ForeColor = System.Drawing.Color.White
        Me.rbRolledOver.Location = New System.Drawing.Point(195, 15)
        Me.rbRolledOver.Name = "rbRolledOver"
        Me.rbRolledOver.Size = New System.Drawing.Size(120, 21)
        Me.rbRolledOver.TabIndex = 4
        Me.rbRolledOver.Text = "ROLLED OVER"
        Me.rbRolledOver.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbRolledOver.UseVisualStyleBackColor = False
        '
        'rbCompleted
        '
        Me.rbCompleted.AutoSize = True
        Me.rbCompleted.BackColor = System.Drawing.Color.LimeGreen
        Me.rbCompleted.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCompleted.ForeColor = System.Drawing.Color.Black
        Me.rbCompleted.Location = New System.Drawing.Point(675, 15)
        Me.rbCompleted.Name = "rbCompleted"
        Me.rbCompleted.Size = New System.Drawing.Size(105, 21)
        Me.rbCompleted.TabIndex = 3
        Me.rbCompleted.Text = "Completed    "
        Me.rbCompleted.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbCompleted.UseVisualStyleBackColor = False
        '
        'rbInProgress
        '
        Me.rbInProgress.AutoSize = True
        Me.rbInProgress.BackColor = System.Drawing.Color.Orange
        Me.rbInProgress.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbInProgress.ForeColor = System.Drawing.Color.Black
        Me.rbInProgress.Location = New System.Drawing.Point(555, 15)
        Me.rbInProgress.Name = "rbInProgress"
        Me.rbInProgress.Size = New System.Drawing.Size(106, 21)
        Me.rbInProgress.TabIndex = 2
        Me.rbInProgress.Text = "In Progress   "
        Me.rbInProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbInProgress.UseVisualStyleBackColor = False
        '
        'rbRolled
        '
        Me.rbRolled.AutoSize = True
        Me.rbRolled.BackColor = System.Drawing.Color.Crimson
        Me.rbRolled.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRolled.ForeColor = System.Drawing.Color.AliceBlue
        Me.rbRolled.Location = New System.Drawing.Point(80, 15)
        Me.rbRolled.Name = "rbRolled"
        Me.rbRolled.Size = New System.Drawing.Size(105, 21)
        Me.rbRolled.TabIndex = 1
        Me.rbRolled.Text = "Rolled           "
        Me.rbRolled.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.rbRolled.UseVisualStyleBackColor = False
        '
        'rbALL
        '
        Me.rbALL.AutoSize = True
        Me.rbALL.BackColor = System.Drawing.Color.AliceBlue
        Me.rbALL.Checked = True
        Me.rbALL.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbALL.ForeColor = System.Drawing.Color.Black
        Me.rbALL.Location = New System.Drawing.Point(10, 15)
        Me.rbALL.Name = "rbALL"
        Me.rbALL.Size = New System.Drawing.Size(58, 21)
        Me.rbALL.TabIndex = 0
        Me.rbALL.TabStop = True
        Me.rbALL.Text = "ALL  "
        Me.rbALL.UseVisualStyleBackColor = False
        '
        'lblTotalRows
        '
        Me.lblTotalRows.AutoSize = True
        Me.lblTotalRows.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblTotalRows.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalRows.Location = New System.Drawing.Point(912, 15)
        Me.lblTotalRows.Name = "lblTotalRows"
        Me.lblTotalRows.Size = New System.Drawing.Size(49, 19)
        Me.lblTotalRows.TabIndex = 9
        Me.lblTotalRows.Text = "Rows"
        '
        'btnTodayOnly
        '
        Me.btnTodayOnly.BackColor = System.Drawing.Color.Red
        Me.btnTodayOnly.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnTodayOnly.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnTodayOnly.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTodayOnly.Location = New System.Drawing.Point(88, 3)
        Me.btnTodayOnly.Name = "btnTodayOnly"
        Me.btnTodayOnly.Size = New System.Drawing.Size(101, 44)
        Me.btnTodayOnly.TabIndex = 13
        Me.btnTodayOnly.Text = "Today Only"
        Me.btnTodayOnly.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.Transparent
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(952, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(101, 44)
        Me.btnClose.TabIndex = 12
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'pnlTopButtons
        '
        Me.pnlTopButtons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTopButtons.BackColor = System.Drawing.Color.RoyalBlue
        Me.pnlTopButtons.Controls.Add(Me.btnImportData)
        Me.pnlTopButtons.Controls.Add(Me.btnClose)
        Me.pnlTopButtons.Controls.Add(Me.btnTodayOnly)
        Me.pnlTopButtons.Controls.Add(Me.btnExportToCSV)
        Me.pnlTopButtons.Location = New System.Drawing.Point(3, 46)
        Me.pnlTopButtons.Name = "pnlTopButtons"
        Me.pnlTopButtons.Size = New System.Drawing.Size(1060, 53)
        Me.pnlTopButtons.TabIndex = 45
        '
        'btnImportData
        '
        Me.btnImportData.BackColor = System.Drawing.Color.Transparent
        Me.btnImportData.BackgroundImage = CType(resources.GetObject("btnImportData.BackgroundImage"), System.Drawing.Image)
        Me.btnImportData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnImportData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnImportData.Font = New System.Drawing.Font("Cambria", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImportData.ImageKey = "(none)"
        Me.btnImportData.ImageList = Me.imglistImportData
        Me.btnImportData.Location = New System.Drawing.Point(437, 8)
        Me.btnImportData.Name = "btnImportData"
        Me.btnImportData.Size = New System.Drawing.Size(195, 30)
        Me.btnImportData.TabIndex = 14
        Me.btnImportData.Tag = "btn4"
        Me.btnImportData.UseVisualStyleBackColor = False
        '
        'imglistImportData
        '
        Me.imglistImportData.ImageStream = CType(resources.GetObject("imglistImportData.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imglistImportData.TransparentColor = System.Drawing.Color.Transparent
        Me.imglistImportData.Images.SetKeyName(0, "Rounded_Gradiant_Green_Button__Unselected_ImportData_30pt_TRANSPARENT_BlackText_C" &
        "ambria.png")
        Me.imglistImportData.Images.SetKeyName(1, "Rounded_Gradiant_Green_Button__Pushed_Button_ImportData_30pt_TRANSPARENT_Cambria." &
        "png")
        Me.imglistImportData.Images.SetKeyName(2, "Rounded_Gradiant_Green_Button__Released_Button_ImportData_30pt_TRANSPARENT_Cambri" &
        "a.png")
        Me.imglistImportData.Images.SetKeyName(3, "Rounded_Gradiant_Amber_Button__Unselected_ImportData_30pt_TRANSPARENT_BlackText_C" &
        "ambria.png")
        Me.imglistImportData.Images.SetKeyName(4, "Rounded_Gradiant_Amber_Button__Pushed_Button_ImportData_30pt_TRANSPARENT_Cambria." &
        "png")
        Me.imglistImportData.Images.SetKeyName(5, "Rounded_Gradiant_Amber_Button__Released_Button_ImportData_30pt_TRANSPARENT_Cambri" &
        "a.png")
        Me.imglistImportData.Images.SetKeyName(6, "Rounded_Gradiant_RED_Button__Unselected_ImportData_30pt_TRANSPARENT_BlackText_Cam" &
        "bria.png")
        Me.imglistImportData.Images.SetKeyName(7, "Rounded_Gradiant_Red_Button__Pushed_Button_ImportData_30pt_TRANSPARENT_Cambria.pn" &
        "g")
        Me.imglistImportData.Images.SetKeyName(8, "Rounded_Gradiant_Red_Button__Released_Button_ImportData_30pt_TRANSPARENT_Cambria." &
        "png")
        '
        'frmReferenceProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Blue
        Me.ClientSize = New System.Drawing.Size(1072, 401)
        Me.Controls.Add(Me.pnlTopButtons)
        Me.Controls.Add(Me.pnlTitle)
        Me.Controls.Add(Me.pnlControls)
        Me.Controls.Add(Me.dgvDeliveryProgress)
        Me.Controls.Add(Me.pnlbuttons)
        Me.ForeColor = System.Drawing.Color.Black
        Me.MinimumSize = New System.Drawing.Size(1088, 440)
        Me.Name = "frmReferenceProgress"
        Me.Text = "Reference Progress"
        Me.pnlTitle.ResumeLayout(False)
        Me.pnlTitle.PerformLayout()
        Me.pnlbuttons.ResumeLayout(False)
        Me.pnlbuttons.PerformLayout()
        CType(Me.dgvDeliveryProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.STATUSMenuStrip.ResumeLayout(False)
        Me.pnlControls.ResumeLayout(False)
        Me.pnlControls.PerformLayout()
        Me.gbProgressFilter.ResumeLayout(False)
        Me.gbProgressFilter.PerformLayout()
        Me.pnlTopButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTitle As Panel
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents pnlbuttons As Panel
    Friend WithEvents btnSearchEmpNo As Button
    Friend WithEvents lblSearchASN As Label
    Friend WithEvents txtSearchASN As TextBox
    Friend WithEvents btnSavePreferences As Button
    Friend WithEvents btnExportToCSV As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtSearchRef As TextBox
    Friend WithEvents dgvDeliveryProgress As DataGridView
    Friend WithEvents Timer_UpdateProgressGrid As Timer
    Friend WithEvents txtNumRows As TextBox
    Friend WithEvents pnlControls As Panel
    Friend WithEvents lblTotalRows As Label
    Friend WithEvents gbProgressFilter As GroupBox
    Friend WithEvents rbCompleted As RadioButton
    Friend WithEvents rbInProgress As RadioButton
    Friend WithEvents rbRolled As RadioButton
    Friend WithEvents rbALL As RadioButton
    Friend WithEvents btnClose As Button
    Friend WithEvents btnTodayOnly As Button
    Friend WithEvents STATUSMenuStrip As ContextMenuStrip
    Friend WithEvents NOSHOWToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CANCELLEDToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents lblLastSaved As Label
    Friend WithEvents txtLastSaved As TextBox
    Friend WithEvents lblGrid As Label
    Friend WithEvents txtGrid As TextBox
    Friend WithEvents lblSupplier As Label
    Friend WithEvents txtSupplier As TextBox
    Friend WithEvents lblReference As Label
    Friend WithEvents txtDeliveryReference As TextBox
    Friend WithEvents lblDeliveryDate As Label
    Friend WithEvents txtDeliveryDate As TextBox
    Friend WithEvents rbCancelled As RadioButton
    Friend WithEvents rbNoShow As RadioButton
    Friend WithEvents rbRolledOver As RadioButton
    Friend WithEvents lblSetNewStatus As Label
    Friend WithEvents lblID As Label
    Friend WithEvents txtID As TextBox
    Friend WithEvents comSTATUS As ComboBox
    Friend WithEvents btnSetNewStatus As Button
    Friend WithEvents txtNewStatus As TextBox
    Friend WithEvents lblCurrentStatus As Label
    Friend WithEvents txtCurrentStatus As TextBox
    Friend WithEvents pnlTopButtons As Panel
    Friend WithEvents ROLLEDOVERToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents btnImportData As Button
    Friend WithEvents imglistImportData As ImageList
End Class
