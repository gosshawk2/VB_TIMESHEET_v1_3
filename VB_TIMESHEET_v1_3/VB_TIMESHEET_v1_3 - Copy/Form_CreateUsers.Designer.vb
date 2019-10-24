<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmCreateUsers
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCreateUsers))
        Me.tabUserEntry = New System.Windows.Forms.TabControl()
        Me.pg_Accounts = New System.Windows.Forms.TabPage()
        Me.pnlAccounts = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtEmpNo = New System.Windows.Forms.TextBox()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.gbEmployee = New System.Windows.Forms.GroupBox()
        Me.rbAgency = New System.Windows.Forms.RadioButton()
        Me.rbECP = New System.Windows.Forms.RadioButton()
        Me.btnRemoveOperator = New System.Windows.Forms.Button()
        Me.btnAddOperator = New System.Windows.Forms.Button()
        Me.txtStartNum = New System.Windows.Forms.TextBox()
        Me.txtTotalQty = New System.Windows.Forms.TextBox()
        Me.btnPopulateUsers = New System.Windows.Forms.Button()
        Me.txtHostname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPasswordRepeat = New System.Windows.Forms.TextBox()
        Me.lblPasswordRepeat = New System.Windows.Forms.Label()
        Me.comAccessRights = New System.Windows.Forms.ComboBox()
        Me.txtAccessRights = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblAccessRights = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.lblLastname_copy = New System.Windows.Forms.Label()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.lblUserBarCode = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.lblFirstname_copy = New System.Windows.Forms.Label()
        Me.txtUserBarcode = New System.Windows.Forms.TextBox()
        Me.txtLastname = New System.Windows.Forms.TextBox()
        Me.txtFirstname = New System.Windows.Forms.TextBox()
        Me.pbBar_Accounts = New System.Windows.Forms.ProgressBar()
        Me.pg_Personal = New System.Windows.Forms.TabPage()
        Me.pnlPersonal = New System.Windows.Forms.Panel()
        Me.lblDateStartedFormat = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.lblJobtitle = New System.Windows.Forms.Label()
        Me.lblLastname = New System.Windows.Forms.Label()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.rtbComments = New System.Windows.Forms.RichTextBox()
        Me.txtJobTitle = New System.Windows.Forms.TextBox()
        Me.txtDepartment = New System.Windows.Forms.TextBox()
        Me.lblDateStarted = New System.Windows.Forms.Label()
        Me.txtWarehouse = New System.Windows.Forms.TextBox()
        Me.lblDepartment = New System.Windows.Forms.Label()
        Me.lblWarehouse = New System.Windows.Forms.Label()
        Me.lblFirstname = New System.Windows.Forms.Label()
        Me.lblEmpNo = New System.Windows.Forms.Label()
        Me.txtDateStarted = New System.Windows.Forms.TextBox()
        Me.txtLastname_view = New System.Windows.Forms.TextBox()
        Me.txtFirstname_view = New System.Windows.Forms.TextBox()
        Me.txtEmpNum = New System.Windows.Forms.TextBox()
        Me.pbBar_Personal = New System.Windows.Forms.ProgressBar()
        Me.pnlbuttons = New System.Windows.Forms.Panel()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSearchEmpNo = New System.Windows.Forms.Button()
        Me.lblSearchEMPNO = New System.Windows.Forms.Label()
        Me.txtSearchEmpNo = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlTitle = New System.Windows.Forms.Panel()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.btnAddExtraLogin = New System.Windows.Forms.Button()
        Me.btnRemoveExtraLogin = New System.Windows.Forms.Button()
        Me.tabUserEntry.SuspendLayout()
        Me.pg_Accounts.SuspendLayout()
        Me.pnlAccounts.SuspendLayout()
        Me.gbEmployee.SuspendLayout()
        Me.pg_Personal.SuspendLayout()
        Me.pnlPersonal.SuspendLayout()
        Me.pnlbuttons.SuspendLayout()
        Me.pnlTitle.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabUserEntry
        '
        Me.tabUserEntry.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.tabUserEntry.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.tabUserEntry.Controls.Add(Me.pg_Accounts)
        Me.tabUserEntry.Controls.Add(Me.pg_Personal)
        Me.tabUserEntry.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabUserEntry.HotTrack = True
        Me.tabUserEntry.ItemSize = New System.Drawing.Size(86, 30)
        Me.tabUserEntry.Location = New System.Drawing.Point(2, 124)
        Me.tabUserEntry.Multiline = True
        Me.tabUserEntry.Name = "tabUserEntry"
        Me.tabUserEntry.SelectedIndex = 0
        Me.tabUserEntry.Size = New System.Drawing.Size(993, 427)
        Me.tabUserEntry.TabIndex = 5
        Me.tabUserEntry.Tag = "0"
        '
        'pg_Accounts
        '
        Me.pg_Accounts.BackColor = System.Drawing.Color.MidnightBlue
        Me.pg_Accounts.Controls.Add(Me.pnlAccounts)
        Me.pg_Accounts.Font = New System.Drawing.Font("Cambria", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pg_Accounts.Location = New System.Drawing.Point(4, 34)
        Me.pg_Accounts.Name = "pg_Accounts"
        Me.pg_Accounts.Padding = New System.Windows.Forms.Padding(3)
        Me.pg_Accounts.Size = New System.Drawing.Size(985, 389)
        Me.pg_Accounts.TabIndex = 1
        Me.pg_Accounts.Text = "Accounts"
        '
        'pnlAccounts
        '
        Me.pnlAccounts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlAccounts.BackColor = System.Drawing.Color.RoyalBlue
        Me.pnlAccounts.Controls.Add(Me.btnRemoveExtraLogin)
        Me.pnlAccounts.Controls.Add(Me.btnAddExtraLogin)
        Me.pnlAccounts.Controls.Add(Me.Label1)
        Me.pnlAccounts.Controls.Add(Me.txtEmpNo)
        Me.pnlAccounts.Controls.Add(Me.txtOutput)
        Me.pnlAccounts.Controls.Add(Me.gbEmployee)
        Me.pnlAccounts.Controls.Add(Me.btnRemoveOperator)
        Me.pnlAccounts.Controls.Add(Me.btnAddOperator)
        Me.pnlAccounts.Controls.Add(Me.txtStartNum)
        Me.pnlAccounts.Controls.Add(Me.txtTotalQty)
        Me.pnlAccounts.Controls.Add(Me.btnPopulateUsers)
        Me.pnlAccounts.Controls.Add(Me.txtHostname)
        Me.pnlAccounts.Controls.Add(Me.Label2)
        Me.pnlAccounts.Controls.Add(Me.txtPasswordRepeat)
        Me.pnlAccounts.Controls.Add(Me.lblPasswordRepeat)
        Me.pnlAccounts.Controls.Add(Me.comAccessRights)
        Me.pnlAccounts.Controls.Add(Me.txtAccessRights)
        Me.pnlAccounts.Controls.Add(Me.txtPassword)
        Me.pnlAccounts.Controls.Add(Me.lblAccessRights)
        Me.pnlAccounts.Controls.Add(Me.lblPassword)
        Me.pnlAccounts.Controls.Add(Me.lblLastname_copy)
        Me.pnlAccounts.Controls.Add(Me.lblUsername)
        Me.pnlAccounts.Controls.Add(Me.lblUserBarCode)
        Me.pnlAccounts.Controls.Add(Me.txtUsername)
        Me.pnlAccounts.Controls.Add(Me.lblFirstname_copy)
        Me.pnlAccounts.Controls.Add(Me.txtUserBarcode)
        Me.pnlAccounts.Controls.Add(Me.txtLastname)
        Me.pnlAccounts.Controls.Add(Me.txtFirstname)
        Me.pnlAccounts.Controls.Add(Me.pbBar_Accounts)
        Me.pnlAccounts.Location = New System.Drawing.Point(5, 5)
        Me.pnlAccounts.Name = "pnlAccounts"
        Me.pnlAccounts.Size = New System.Drawing.Size(974, 376)
        Me.pnlAccounts.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.GreenYellow
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 106)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 21)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "EmpNo:"
        '
        'txtEmpNo
        '
        Me.txtEmpNo.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtEmpNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmpNo.ForeColor = System.Drawing.Color.Yellow
        Me.txtEmpNo.Location = New System.Drawing.Point(180, 106)
        Me.txtEmpNo.Name = "txtEmpNo"
        Me.txtEmpNo.Size = New System.Drawing.Size(96, 23)
        Me.txtEmpNo.TabIndex = 46
        '
        'txtOutput
        '
        Me.txtOutput.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtOutput.Font = New System.Drawing.Font("Cambria", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutput.ForeColor = System.Drawing.Color.Black
        Me.txtOutput.Location = New System.Drawing.Point(550, 221)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(385, 113)
        Me.txtOutput.TabIndex = 45
        Me.txtOutput.Visible = False
        '
        'gbEmployee
        '
        Me.gbEmployee.Controls.Add(Me.rbAgency)
        Me.gbEmployee.Controls.Add(Me.rbECP)
        Me.gbEmployee.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.gbEmployee.Location = New System.Drawing.Point(553, 64)
        Me.gbEmployee.Name = "gbEmployee"
        Me.gbEmployee.Size = New System.Drawing.Size(382, 53)
        Me.gbEmployee.TabIndex = 44
        Me.gbEmployee.TabStop = False
        Me.gbEmployee.Visible = False
        '
        'rbAgency
        '
        Me.rbAgency.AutoSize = True
        Me.rbAgency.Location = New System.Drawing.Point(205, 21)
        Me.rbAgency.Name = "rbAgency"
        Me.rbAgency.Size = New System.Drawing.Size(80, 21)
        Me.rbAgency.TabIndex = 1
        Me.rbAgency.Text = "AGENCY"
        Me.rbAgency.UseVisualStyleBackColor = True
        '
        'rbECP
        '
        Me.rbECP.AutoSize = True
        Me.rbECP.Checked = True
        Me.rbECP.Location = New System.Drawing.Point(40, 21)
        Me.rbECP.Name = "rbECP"
        Me.rbECP.Size = New System.Drawing.Size(52, 21)
        Me.rbECP.TabIndex = 0
        Me.rbECP.TabStop = True
        Me.rbECP.Text = "ECP"
        Me.rbECP.UseVisualStyleBackColor = True
        '
        'btnRemoveOperator
        '
        Me.btnRemoveOperator.BackColor = System.Drawing.Color.PaleGreen
        Me.btnRemoveOperator.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveOperator.Location = New System.Drawing.Point(550, 174)
        Me.btnRemoveOperator.Name = "btnRemoveOperator"
        Me.btnRemoveOperator.Size = New System.Drawing.Size(185, 30)
        Me.btnRemoveOperator.TabIndex = 43
        Me.btnRemoveOperator.Text = "Remove From Operators"
        Me.btnRemoveOperator.UseVisualStyleBackColor = False
        Me.btnRemoveOperator.Visible = False
        '
        'btnAddOperator
        '
        Me.btnAddOperator.BackColor = System.Drawing.Color.PaleGreen
        Me.btnAddOperator.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddOperator.Location = New System.Drawing.Point(550, 123)
        Me.btnAddOperator.Name = "btnAddOperator"
        Me.btnAddOperator.Size = New System.Drawing.Size(185, 30)
        Me.btnAddOperator.TabIndex = 42
        Me.btnAddOperator.Text = "Add to Operators"
        Me.btnAddOperator.UseVisualStyleBackColor = False
        Me.btnAddOperator.Visible = False
        '
        'txtStartNum
        '
        Me.txtStartNum.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtStartNum.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartNum.ForeColor = System.Drawing.Color.Yellow
        Me.txtStartNum.Location = New System.Drawing.Point(758, 39)
        Me.txtStartNum.Name = "txtStartNum"
        Me.txtStartNum.Size = New System.Drawing.Size(74, 23)
        Me.txtStartNum.TabIndex = 41
        Me.txtStartNum.Text = "Start Num"
        Me.txtStartNum.Visible = False
        '
        'txtTotalQty
        '
        Me.txtTotalQty.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtTotalQty.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalQty.ForeColor = System.Drawing.Color.Yellow
        Me.txtTotalQty.Location = New System.Drawing.Point(847, 39)
        Me.txtTotalQty.Name = "txtTotalQty"
        Me.txtTotalQty.Size = New System.Drawing.Size(74, 23)
        Me.txtTotalQty.TabIndex = 40
        Me.txtTotalQty.Text = "Qty"
        Me.txtTotalQty.Visible = False
        '
        'btnPopulateUsers
        '
        Me.btnPopulateUsers.BackColor = System.Drawing.Color.DarkGreen
        Me.btnPopulateUsers.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPopulateUsers.Location = New System.Drawing.Point(550, 34)
        Me.btnPopulateUsers.Name = "btnPopulateUsers"
        Me.btnPopulateUsers.Size = New System.Drawing.Size(185, 30)
        Me.btnPopulateUsers.TabIndex = 4
        Me.btnPopulateUsers.Text = "Populate"
        Me.btnPopulateUsers.UseVisualStyleBackColor = False
        Me.btnPopulateUsers.Visible = False
        '
        'txtHostname
        '
        Me.txtHostname.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtHostname.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHostname.ForeColor = System.Drawing.Color.Yellow
        Me.txtHostname.Location = New System.Drawing.Point(180, 302)
        Me.txtHostname.Name = "txtHostname"
        Me.txtHostname.Size = New System.Drawing.Size(330, 23)
        Me.txtHostname.TabIndex = 38
        Me.txtHostname.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.GreenYellow
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 302)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 21)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Hostname:"
        Me.Label2.Visible = False
        '
        'txtPasswordRepeat
        '
        Me.txtPasswordRepeat.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtPasswordRepeat.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPasswordRepeat.ForeColor = System.Drawing.Color.Yellow
        Me.txtPasswordRepeat.Location = New System.Drawing.Point(180, 260)
        Me.txtPasswordRepeat.Name = "txtPasswordRepeat"
        Me.txtPasswordRepeat.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPasswordRepeat.Size = New System.Drawing.Size(330, 23)
        Me.txtPasswordRepeat.TabIndex = 21
        '
        'lblPasswordRepeat
        '
        Me.lblPasswordRepeat.AutoSize = True
        Me.lblPasswordRepeat.BackColor = System.Drawing.Color.GreenYellow
        Me.lblPasswordRepeat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPasswordRepeat.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPasswordRepeat.Location = New System.Drawing.Point(12, 260)
        Me.lblPasswordRepeat.Name = "lblPasswordRepeat"
        Me.lblPasswordRepeat.Size = New System.Drawing.Size(141, 21)
        Me.lblPasswordRepeat.TabIndex = 37
        Me.lblPasswordRepeat.Text = "Renter Password:"
        '
        'comAccessRights
        '
        Me.comAccessRights.BackColor = System.Drawing.Color.AliceBlue
        Me.comAccessRights.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comAccessRights.ForeColor = System.Drawing.Color.Black
        Me.comAccessRights.FormattingEnabled = True
        Me.comAccessRights.ItemHeight = 19
        Me.comAccessRights.Location = New System.Drawing.Point(550, 343)
        Me.comAccessRights.Name = "comAccessRights"
        Me.comAccessRights.Size = New System.Drawing.Size(192, 27)
        Me.comAccessRights.TabIndex = 23
        '
        'txtAccessRights
        '
        Me.txtAccessRights.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtAccessRights.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccessRights.ForeColor = System.Drawing.Color.Yellow
        Me.txtAccessRights.Location = New System.Drawing.Point(180, 344)
        Me.txtAccessRights.Name = "txtAccessRights"
        Me.txtAccessRights.Size = New System.Drawing.Size(330, 23)
        Me.txtAccessRights.TabIndex = 22
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtPassword.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.ForeColor = System.Drawing.Color.Yellow
        Me.txtPassword.Location = New System.Drawing.Point(180, 220)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(330, 23)
        Me.txtPassword.TabIndex = 20
        '
        'lblAccessRights
        '
        Me.lblAccessRights.AutoSize = True
        Me.lblAccessRights.BackColor = System.Drawing.Color.GreenYellow
        Me.lblAccessRights.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAccessRights.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAccessRights.Location = New System.Drawing.Point(12, 344)
        Me.lblAccessRights.Name = "lblAccessRights"
        Me.lblAccessRights.Size = New System.Drawing.Size(114, 21)
        Me.lblAccessRights.TabIndex = 32
        Me.lblAccessRights.Text = "Access Rights:"
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.BackColor = System.Drawing.Color.GreenYellow
        Me.lblPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPassword.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.Location = New System.Drawing.Point(12, 220)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(87, 21)
        Me.lblPassword.TabIndex = 31
        Me.lblPassword.Text = "Password:"
        '
        'lblLastname_copy
        '
        Me.lblLastname_copy.AutoSize = True
        Me.lblLastname_copy.BackColor = System.Drawing.Color.GreenYellow
        Me.lblLastname_copy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLastname_copy.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastname_copy.Location = New System.Drawing.Point(12, 64)
        Me.lblLastname_copy.Name = "lblLastname_copy"
        Me.lblLastname_copy.Size = New System.Drawing.Size(87, 21)
        Me.lblLastname_copy.TabIndex = 29
        Me.lblLastname_copy.Text = "Lastname:"
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.BackColor = System.Drawing.Color.GreenYellow
        Me.lblUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblUsername.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsername.Location = New System.Drawing.Point(12, 180)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(89, 21)
        Me.lblUsername.TabIndex = 28
        Me.lblUsername.Text = "Username:"
        '
        'lblUserBarCode
        '
        Me.lblUserBarCode.AutoSize = True
        Me.lblUserBarCode.BackColor = System.Drawing.Color.GreenYellow
        Me.lblUserBarCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblUserBarCode.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserBarCode.Location = New System.Drawing.Point(12, 142)
        Me.lblUserBarCode.Name = "lblUserBarCode"
        Me.lblUserBarCode.Size = New System.Drawing.Size(118, 21)
        Me.lblUserBarCode.TabIndex = 22
        Me.lblUserBarCode.Text = "User Bar Code:"
        '
        'txtUsername
        '
        Me.txtUsername.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtUsername.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.ForeColor = System.Drawing.Color.Yellow
        Me.txtUsername.Location = New System.Drawing.Point(180, 180)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(330, 23)
        Me.txtUsername.TabIndex = 19
        '
        'lblFirstname_copy
        '
        Me.lblFirstname_copy.AutoSize = True
        Me.lblFirstname_copy.BackColor = System.Drawing.Color.GreenYellow
        Me.lblFirstname_copy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFirstname_copy.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirstname_copy.Location = New System.Drawing.Point(12, 34)
        Me.lblFirstname_copy.Name = "lblFirstname_copy"
        Me.lblFirstname_copy.Size = New System.Drawing.Size(90, 21)
        Me.lblFirstname_copy.TabIndex = 14
        Me.lblFirstname_copy.Text = "Firstname:"
        '
        'txtUserBarcode
        '
        Me.txtUserBarcode.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtUserBarcode.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserBarcode.ForeColor = System.Drawing.Color.Yellow
        Me.txtUserBarcode.Location = New System.Drawing.Point(180, 142)
        Me.txtUserBarcode.Name = "txtUserBarcode"
        Me.txtUserBarcode.Size = New System.Drawing.Size(190, 23)
        Me.txtUserBarcode.TabIndex = 18
        '
        'txtLastname
        '
        Me.txtLastname.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtLastname.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastname.ForeColor = System.Drawing.Color.Black
        Me.txtLastname.Location = New System.Drawing.Point(180, 64)
        Me.txtLastname.Name = "txtLastname"
        Me.txtLastname.Size = New System.Drawing.Size(330, 23)
        Me.txtLastname.TabIndex = 17
        '
        'txtFirstname
        '
        Me.txtFirstname.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtFirstname.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstname.ForeColor = System.Drawing.Color.Black
        Me.txtFirstname.Location = New System.Drawing.Point(180, 34)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(330, 23)
        Me.txtFirstname.TabIndex = 16
        '
        'pbBar_Accounts
        '
        Me.pbBar_Accounts.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbBar_Accounts.BackColor = System.Drawing.Color.Red
        Me.pbBar_Accounts.ForeColor = System.Drawing.Color.Chartreuse
        Me.pbBar_Accounts.Location = New System.Drawing.Point(1, 1)
        Me.pbBar_Accounts.Name = "pbBar_Accounts"
        Me.pbBar_Accounts.Size = New System.Drawing.Size(973, 20)
        Me.pbBar_Accounts.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pbBar_Accounts.TabIndex = 2
        Me.pbBar_Accounts.UseWaitCursor = True
        '
        'pg_Personal
        '
        Me.pg_Personal.BackColor = System.Drawing.Color.DeepPink
        Me.pg_Personal.Controls.Add(Me.pnlPersonal)
        Me.pg_Personal.Location = New System.Drawing.Point(4, 34)
        Me.pg_Personal.Name = "pg_Personal"
        Me.pg_Personal.Padding = New System.Windows.Forms.Padding(3)
        Me.pg_Personal.Size = New System.Drawing.Size(985, 389)
        Me.pg_Personal.TabIndex = 0
        Me.pg_Personal.Text = "Personal"
        '
        'pnlPersonal
        '
        Me.pnlPersonal.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlPersonal.BackColor = System.Drawing.Color.RoyalBlue
        Me.pnlPersonal.Controls.Add(Me.lblDateStartedFormat)
        Me.pnlPersonal.Controls.Add(Me.txtID)
        Me.pnlPersonal.Controls.Add(Me.lblJobtitle)
        Me.pnlPersonal.Controls.Add(Me.lblLastname)
        Me.pnlPersonal.Controls.Add(Me.lblComments)
        Me.pnlPersonal.Controls.Add(Me.rtbComments)
        Me.pnlPersonal.Controls.Add(Me.txtJobTitle)
        Me.pnlPersonal.Controls.Add(Me.txtDepartment)
        Me.pnlPersonal.Controls.Add(Me.lblDateStarted)
        Me.pnlPersonal.Controls.Add(Me.txtWarehouse)
        Me.pnlPersonal.Controls.Add(Me.lblDepartment)
        Me.pnlPersonal.Controls.Add(Me.lblWarehouse)
        Me.pnlPersonal.Controls.Add(Me.lblFirstname)
        Me.pnlPersonal.Controls.Add(Me.lblEmpNo)
        Me.pnlPersonal.Controls.Add(Me.txtDateStarted)
        Me.pnlPersonal.Controls.Add(Me.txtLastname_view)
        Me.pnlPersonal.Controls.Add(Me.txtFirstname_view)
        Me.pnlPersonal.Controls.Add(Me.txtEmpNum)
        Me.pnlPersonal.Controls.Add(Me.pbBar_Personal)
        Me.pnlPersonal.Location = New System.Drawing.Point(5, 7)
        Me.pnlPersonal.Name = "pnlPersonal"
        Me.pnlPersonal.Size = New System.Drawing.Size(974, 375)
        Me.pnlPersonal.TabIndex = 3
        '
        'lblDateStartedFormat
        '
        Me.lblDateStartedFormat.AutoSize = True
        Me.lblDateStartedFormat.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateStartedFormat.Location = New System.Drawing.Point(410, 290)
        Me.lblDateStartedFormat.Name = "lblDateStartedFormat"
        Me.lblDateStartedFormat.Size = New System.Drawing.Size(101, 19)
        Me.lblDateStartedFormat.TabIndex = 37
        Me.lblDateStartedFormat.Text = "yyyy/mm/dd"
        '
        'txtID
        '
        Me.txtID.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtID.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.ForeColor = System.Drawing.Color.Yellow
        Me.txtID.Location = New System.Drawing.Point(414, 50)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(96, 23)
        Me.txtID.TabIndex = 7
        Me.txtID.TabStop = False
        Me.txtID.Visible = False
        '
        'lblJobtitle
        '
        Me.lblJobtitle.AutoSize = True
        Me.lblJobtitle.BackColor = System.Drawing.Color.GreenYellow
        Me.lblJobtitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblJobtitle.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobtitle.Location = New System.Drawing.Point(12, 250)
        Me.lblJobtitle.Name = "lblJobtitle"
        Me.lblJobtitle.Size = New System.Drawing.Size(76, 21)
        Me.lblJobtitle.TabIndex = 30
        Me.lblJobtitle.Text = "Job Title:"
        '
        'lblLastname
        '
        Me.lblLastname.AutoSize = True
        Me.lblLastname.BackColor = System.Drawing.Color.GreenYellow
        Me.lblLastname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLastname.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastname.Location = New System.Drawing.Point(12, 130)
        Me.lblLastname.Name = "lblLastname"
        Me.lblLastname.Size = New System.Drawing.Size(87, 21)
        Me.lblLastname.TabIndex = 29
        Me.lblLastname.Text = "Lastname:"
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.BackColor = System.Drawing.Color.GreenYellow
        Me.lblComments.Font = New System.Drawing.Font("Cambria", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.Location = New System.Drawing.Point(550, 50)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(108, 22)
        Me.lblComments.TabIndex = 25
        Me.lblComments.Text = "Comments:"
        '
        'rtbComments
        '
        Me.rtbComments.BackColor = System.Drawing.Color.AliceBlue
        Me.rtbComments.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbComments.Location = New System.Drawing.Point(550, 90)
        Me.rtbComments.Name = "rtbComments"
        Me.rtbComments.Size = New System.Drawing.Size(371, 183)
        Me.rtbComments.TabIndex = 15
        Me.rtbComments.Text = ""
        '
        'txtJobTitle
        '
        Me.txtJobTitle.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtJobTitle.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobTitle.ForeColor = System.Drawing.Color.Yellow
        Me.txtJobTitle.Location = New System.Drawing.Point(180, 250)
        Me.txtJobTitle.Name = "txtJobTitle"
        Me.txtJobTitle.Size = New System.Drawing.Size(330, 23)
        Me.txtJobTitle.TabIndex = 12
        '
        'txtDepartment
        '
        Me.txtDepartment.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtDepartment.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDepartment.ForeColor = System.Drawing.Color.Yellow
        Me.txtDepartment.Location = New System.Drawing.Point(180, 210)
        Me.txtDepartment.Name = "txtDepartment"
        Me.txtDepartment.Size = New System.Drawing.Size(330, 23)
        Me.txtDepartment.TabIndex = 11
        '
        'lblDateStarted
        '
        Me.lblDateStarted.AutoSize = True
        Me.lblDateStarted.BackColor = System.Drawing.Color.GreenYellow
        Me.lblDateStarted.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDateStarted.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateStarted.Location = New System.Drawing.Point(12, 290)
        Me.lblDateStarted.Name = "lblDateStarted"
        Me.lblDateStarted.Size = New System.Drawing.Size(107, 21)
        Me.lblDateStarted.TabIndex = 20
        Me.lblDateStarted.Text = "Date Started:"
        '
        'txtWarehouse
        '
        Me.txtWarehouse.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtWarehouse.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWarehouse.ForeColor = System.Drawing.Color.Yellow
        Me.txtWarehouse.Location = New System.Drawing.Point(180, 170)
        Me.txtWarehouse.Name = "txtWarehouse"
        Me.txtWarehouse.Size = New System.Drawing.Size(330, 23)
        Me.txtWarehouse.TabIndex = 10
        '
        'lblDepartment
        '
        Me.lblDepartment.AutoSize = True
        Me.lblDepartment.BackColor = System.Drawing.Color.GreenYellow
        Me.lblDepartment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDepartment.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDepartment.Location = New System.Drawing.Point(12, 210)
        Me.lblDepartment.Name = "lblDepartment"
        Me.lblDepartment.Size = New System.Drawing.Size(104, 21)
        Me.lblDepartment.TabIndex = 18
        Me.lblDepartment.Text = "Department:"
        '
        'lblWarehouse
        '
        Me.lblWarehouse.AutoSize = True
        Me.lblWarehouse.BackColor = System.Drawing.Color.GreenYellow
        Me.lblWarehouse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblWarehouse.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarehouse.Location = New System.Drawing.Point(12, 170)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(97, 21)
        Me.lblWarehouse.TabIndex = 16
        Me.lblWarehouse.Text = "Warehouse:"
        '
        'lblFirstname
        '
        Me.lblFirstname.AutoSize = True
        Me.lblFirstname.BackColor = System.Drawing.Color.GreenYellow
        Me.lblFirstname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFirstname.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirstname.Location = New System.Drawing.Point(12, 90)
        Me.lblFirstname.Name = "lblFirstname"
        Me.lblFirstname.Size = New System.Drawing.Size(90, 21)
        Me.lblFirstname.TabIndex = 14
        Me.lblFirstname.Text = "Firstname:"
        '
        'lblEmpNo
        '
        Me.lblEmpNo.AutoSize = True
        Me.lblEmpNo.BackColor = System.Drawing.Color.GreenYellow
        Me.lblEmpNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmpNo.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpNo.Location = New System.Drawing.Point(12, 50)
        Me.lblEmpNo.Name = "lblEmpNo"
        Me.lblEmpNo.Size = New System.Drawing.Size(68, 21)
        Me.lblEmpNo.TabIndex = 12
        Me.lblEmpNo.Text = "EmpNo:"
        '
        'txtDateStarted
        '
        Me.txtDateStarted.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtDateStarted.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateStarted.ForeColor = System.Drawing.Color.Yellow
        Me.txtDateStarted.Location = New System.Drawing.Point(180, 290)
        Me.txtDateStarted.Name = "txtDateStarted"
        Me.txtDateStarted.Size = New System.Drawing.Size(190, 23)
        Me.txtDateStarted.TabIndex = 13
        '
        'txtLastname_view
        '
        Me.txtLastname_view.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtLastname_view.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastname_view.ForeColor = System.Drawing.Color.Yellow
        Me.txtLastname_view.Location = New System.Drawing.Point(180, 130)
        Me.txtLastname_view.Name = "txtLastname_view"
        Me.txtLastname_view.ReadOnly = True
        Me.txtLastname_view.Size = New System.Drawing.Size(330, 23)
        Me.txtLastname_view.TabIndex = 9
        '
        'txtFirstname_view
        '
        Me.txtFirstname_view.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtFirstname_view.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstname_view.ForeColor = System.Drawing.Color.Yellow
        Me.txtFirstname_view.Location = New System.Drawing.Point(180, 90)
        Me.txtFirstname_view.Name = "txtFirstname_view"
        Me.txtFirstname_view.ReadOnly = True
        Me.txtFirstname_view.Size = New System.Drawing.Size(330, 23)
        Me.txtFirstname_view.TabIndex = 8
        '
        'txtEmpNum
        '
        Me.txtEmpNum.BackColor = System.Drawing.Color.DodgerBlue
        Me.txtEmpNum.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmpNum.ForeColor = System.Drawing.Color.Yellow
        Me.txtEmpNum.Location = New System.Drawing.Point(180, 50)
        Me.txtEmpNum.Name = "txtEmpNum"
        Me.txtEmpNum.Size = New System.Drawing.Size(96, 23)
        Me.txtEmpNum.TabIndex = 6
        '
        'pbBar_Personal
        '
        Me.pbBar_Personal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbBar_Personal.BackColor = System.Drawing.Color.Red
        Me.pbBar_Personal.ForeColor = System.Drawing.Color.Chartreuse
        Me.pbBar_Personal.Location = New System.Drawing.Point(0, 0)
        Me.pbBar_Personal.Name = "pbBar_Personal"
        Me.pbBar_Personal.Size = New System.Drawing.Size(974, 20)
        Me.pbBar_Personal.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pbBar_Personal.TabIndex = 2
        Me.pbBar_Personal.UseWaitCursor = True
        '
        'pnlbuttons
        '
        Me.pnlbuttons.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlbuttons.BackColor = System.Drawing.Color.MidnightBlue
        Me.pnlbuttons.Controls.Add(Me.btnUpdate)
        Me.pnlbuttons.Controls.Add(Me.btnSearchEmpNo)
        Me.pnlbuttons.Controls.Add(Me.lblSearchEMPNO)
        Me.pnlbuttons.Controls.Add(Me.txtSearchEmpNo)
        Me.pnlbuttons.Controls.Add(Me.btnClear)
        Me.pnlbuttons.Controls.Add(Me.btnSave)
        Me.pnlbuttons.Controls.Add(Me.btnClose)
        Me.pnlbuttons.Location = New System.Drawing.Point(5, 61)
        Me.pnlbuttons.Name = "pnlbuttons"
        Me.pnlbuttons.Size = New System.Drawing.Size(983, 50)
        Me.pnlbuttons.TabIndex = 39
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.MediumPurple
        Me.btnUpdate.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Location = New System.Drawing.Point(358, 11)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(140, 30)
        Me.btnUpdate.TabIndex = 8
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = False
        Me.btnUpdate.Visible = False
        '
        'btnSearchEmpNo
        '
        Me.btnSearchEmpNo.BackColor = System.Drawing.Color.SpringGreen
        Me.btnSearchEmpNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchEmpNo.Location = New System.Drawing.Point(790, 11)
        Me.btnSearchEmpNo.Name = "btnSearchEmpNo"
        Me.btnSearchEmpNo.Size = New System.Drawing.Size(49, 30)
        Me.btnSearchEmpNo.TabIndex = 7
        Me.btnSearchEmpNo.Text = "Go"
        Me.btnSearchEmpNo.UseVisualStyleBackColor = False
        '
        'lblSearchEMPNO
        '
        Me.lblSearchEMPNO.AutoSize = True
        Me.lblSearchEMPNO.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.lblSearchEMPNO.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchEMPNO.Location = New System.Drawing.Point(533, 15)
        Me.lblSearchEMPNO.Name = "lblSearchEMPNO"
        Me.lblSearchEMPNO.Size = New System.Drawing.Size(124, 19)
        Me.lblSearchEMPNO.TabIndex = 6
        Me.lblSearchEMPNO.Text = "Search Emp No:"
        '
        'txtSearchEmpNo
        '
        Me.txtSearchEmpNo.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchEmpNo.Location = New System.Drawing.Point(665, 15)
        Me.txtSearchEmpNo.Name = "txtSearchEmpNo"
        Me.txtSearchEmpNo.Size = New System.Drawing.Size(119, 23)
        Me.txtSearchEmpNo.TabIndex = 5
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.SpringGreen
        Me.btnClear.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(188, 11)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(140, 30)
        Me.btnClear.TabIndex = 2
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = False
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
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.Red
        Me.btnClose.BackgroundImage = CType(resources.GetObject("btnClose.BackgroundImage"), System.Drawing.Image)
        Me.btnClose.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnClose.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(868, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 44)
        Me.btnClose.TabIndex = 3
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'pnlTitle
        '
        Me.pnlTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTitle.BackColor = System.Drawing.Color.Blue
        Me.pnlTitle.Controls.Add(Me.txtTitle)
        Me.pnlTitle.ForeColor = System.Drawing.Color.Black
        Me.pnlTitle.Location = New System.Drawing.Point(5, 12)
        Me.pnlTitle.Name = "pnlTitle"
        Me.pnlTitle.Size = New System.Drawing.Size(983, 48)
        Me.pnlTitle.TabIndex = 40
        '
        'txtTitle
        '
        Me.txtTitle.BackColor = System.Drawing.Color.DarkBlue
        Me.txtTitle.Font = New System.Drawing.Font("Cambria", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.AliceBlue
        Me.txtTitle.Location = New System.Drawing.Point(0, 5)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(980, 36)
        Me.txtTitle.TabIndex = 0
        Me.txtTitle.Text = "Please Enter Account Details"
        Me.txtTitle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtMessage
        '
        Me.txtMessage.BackColor = System.Drawing.Color.AliceBlue
        Me.txtMessage.Font = New System.Drawing.Font("Cambria", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMessage.ForeColor = System.Drawing.Color.Black
        Me.txtMessage.Location = New System.Drawing.Point(6, 551)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(989, 57)
        Me.txtMessage.TabIndex = 38
        '
        'btnAddExtraLogin
        '
        Me.btnAddExtraLogin.BackColor = System.Drawing.Color.PaleGreen
        Me.btnAddExtraLogin.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddExtraLogin.Location = New System.Drawing.Point(758, 123)
        Me.btnAddExtraLogin.Name = "btnAddExtraLogin"
        Me.btnAddExtraLogin.Size = New System.Drawing.Size(177, 30)
        Me.btnAddExtraLogin.TabIndex = 48
        Me.btnAddExtraLogin.Text = "Add Extra Login:"
        Me.btnAddExtraLogin.UseVisualStyleBackColor = False
        Me.btnAddExtraLogin.Visible = False
        '
        'btnRemoveExtraLogin
        '
        Me.btnRemoveExtraLogin.BackColor = System.Drawing.Color.PaleGreen
        Me.btnRemoveExtraLogin.Font = New System.Drawing.Font("Cambria", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveExtraLogin.Location = New System.Drawing.Point(758, 173)
        Me.btnRemoveExtraLogin.Name = "btnRemoveExtraLogin"
        Me.btnRemoveExtraLogin.Size = New System.Drawing.Size(177, 30)
        Me.btnRemoveExtraLogin.TabIndex = 49
        Me.btnRemoveExtraLogin.Text = "Remove Extra Login:"
        Me.btnRemoveExtraLogin.UseVisualStyleBackColor = False
        Me.btnRemoveExtraLogin.Visible = False
        '
        'frmCreateUsers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1007, 613)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.pnlTitle)
        Me.Controls.Add(Me.pnlbuttons)
        Me.Controls.Add(Me.tabUserEntry)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCreateUsers"
        Me.Text = "Asset Register - Create Users"
        Me.tabUserEntry.ResumeLayout(False)
        Me.pg_Accounts.ResumeLayout(False)
        Me.pnlAccounts.ResumeLayout(False)
        Me.pnlAccounts.PerformLayout()
        Me.gbEmployee.ResumeLayout(False)
        Me.gbEmployee.PerformLayout()
        Me.pg_Personal.ResumeLayout(False)
        Me.pnlPersonal.ResumeLayout(False)
        Me.pnlPersonal.PerformLayout()
        Me.pnlbuttons.ResumeLayout(False)
        Me.pnlbuttons.PerformLayout()
        Me.pnlTitle.ResumeLayout(False)
        Me.pnlTitle.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tabUserEntry As TabControl
    Friend WithEvents pg_Accounts As TabPage
    Friend WithEvents pg_Personal As TabPage
    Friend WithEvents pnlAccounts As Panel
    Friend WithEvents comAccessRights As ComboBox
    Friend WithEvents txtAccessRights As TextBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents lblAccessRights As Label
    Friend WithEvents lblPassword As Label
    Friend WithEvents lblLastname_copy As Label
    Friend WithEvents lblUsername As Label
    Friend WithEvents lblUserBarCode As Label
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents lblFirstname_copy As Label
    Friend WithEvents txtUserBarcode As TextBox
    Friend WithEvents txtLastname As TextBox
    Friend WithEvents txtFirstname As TextBox
    Friend WithEvents pbBar_Accounts As ProgressBar
    Friend WithEvents pnlPersonal As Panel
    Friend WithEvents lblDateStartedFormat As Label
    Friend WithEvents txtID As TextBox
    Friend WithEvents lblJobtitle As Label
    Friend WithEvents lblLastname As Label
    Friend WithEvents lblComments As Label
    Friend WithEvents rtbComments As RichTextBox
    Friend WithEvents txtJobTitle As TextBox
    Friend WithEvents txtDepartment As TextBox
    Friend WithEvents lblDateStarted As Label
    Friend WithEvents txtWarehouse As TextBox
    Friend WithEvents lblDepartment As Label
    Friend WithEvents lblWarehouse As Label
    Friend WithEvents lblFirstname As Label
    Friend WithEvents lblEmpNo As Label
    Friend WithEvents txtDateStarted As TextBox
    Friend WithEvents txtLastname_view As TextBox
    Friend WithEvents txtFirstname_view As TextBox
    Friend WithEvents txtEmpNum As TextBox
    Friend WithEvents pbBar_Personal As ProgressBar
    Friend WithEvents pnlbuttons As Panel
    Friend WithEvents btnClear As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents pnlTitle As Panel
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents txtPasswordRepeat As TextBox
    Friend WithEvents lblPasswordRepeat As Label
    Friend WithEvents txtMessage As TextBox
    Friend WithEvents txtHostname As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnPopulateUsers As Button
    Friend WithEvents txtStartNum As TextBox
    Friend WithEvents txtTotalQty As TextBox
    Friend WithEvents btnSearchEmpNo As Button
    Friend WithEvents lblSearchEMPNO As Label
    Friend WithEvents txtSearchEmpNo As TextBox
    Friend WithEvents btnRemoveOperator As Button
    Friend WithEvents btnAddOperator As Button
    Friend WithEvents gbEmployee As GroupBox
    Friend WithEvents rbAgency As RadioButton
    Friend WithEvents rbECP As RadioButton
    Friend WithEvents btnUpdate As Button
    Friend WithEvents txtOutput As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtEmpNo As TextBox
    Friend WithEvents btnRemoveExtraLogin As Button
    Friend WithEvents btnAddExtraLogin As Button
End Class
