<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmLoadExcel
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLoadExcel))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblVendorDesc = New System.Windows.Forms.Label()
        Me.txtVendorNo = New System.Windows.Forms.TextBox()
        Me.btnValidVendor = New System.Windows.Forms.Button()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmdExcel = New System.Windows.Forms.Button()
        Me.lblExcel = New System.Windows.Forms.Label()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblUsrLog = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.clPRHCOD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDPTN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVMVNUM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDCTP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDMFR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPQPRC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPQMIN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDSTS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.dtProjectDate = New System.Windows.Forms.DateTimePicker()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.EditReference = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.AddReference = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.clPRDPTN2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVMVNUM2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDCTP2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDMFR2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPQPRC2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPQMIN2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clError = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BindingNavigator2 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorCountItem1 = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorMoveFirstItem1 = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem1 = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem1 = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem1 = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem1 = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.cmbPerCharge = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnSuccess = New System.Windows.Forms.Button()
        Me.btnCheck = New System.Windows.Forms.Button()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.txtProjectNo = New System.Windows.Forms.TextBox()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblPerCharge = New System.Windows.Forms.Label()
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.lblProjectNo = New System.Windows.Forms.Label()
        Me.lblVendorNo = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel10 = New System.Windows.Forms.TableLayoutPanel()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel8 = New System.Windows.Forms.TableLayoutPanel()
        Me.ac2 = New CTP_IS_VBNET.Autocomplete_Textbox()
        Me.cmbStatusMore = New System.Windows.Forms.ComboBox()
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingNavigator2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel10.SuspendLayout()
        Me.TableLayoutPanel8.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "excel.png")
        Me.ImageList1.Images.SetKeyName(1, "493-4936787_free-png-search-icon-magnifying-glass-icon-png.png")
        Me.ImageList1.Images.SetKeyName(2, "bulk-icon-24.jpg")
        Me.ImageList1.Images.SetKeyName(3, "proddev.png")
        Me.ImageList1.Images.SetKeyName(4, "proddev1.png")
        Me.ImageList1.Images.SetKeyName(5, "proddev2.png")
        Me.ImageList1.Images.SetKeyName(6, "proddev3.png")
        Me.ImageList1.Images.SetKeyName(7, "downloadex1.png")
        Me.ImageList1.Images.SetKeyName(8, "downloadex2.png")
        Me.ImageList1.Images.SetKeyName(9, "downloadex3.png")
        Me.ImageList1.Images.SetKeyName(10, "downloadex4.png")
        Me.ImageList1.Images.SetKeyName(11, "downloadex5.png")
        Me.ImageList1.Images.SetKeyName(12, "80785927-icono-de-información-aislado-en-la-ilustración-abstracta-de-botón-redond" &
        "o-amarillo-vidrioso.jpg")
        Me.ImageList1.Images.SetKeyName(13, "89850586-info-icon-isolated-on-special-yellow-round-button-abstract-illustration." &
        "jpg")
        Me.ImageList1.Images.SetKeyName(14, "images (1).png")
        Me.ImageList1.Images.SetKeyName(15, "images.png")
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 82.00837!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.99163!))
        Me.TableLayoutPanel4.Controls.Add(Me.lblVendorDesc, 0, 1)
        Me.TableLayoutPanel4.Controls.Add(Me.txtVendorNo, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.btnValidVendor, 1, 0)
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(248, 127)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 2
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(239, 62)
        Me.TableLayoutPanel4.TabIndex = 31
        '
        'lblVendorDesc
        '
        Me.TableLayoutPanel4.SetColumnSpan(Me.lblVendorDesc, 2)
        Me.lblVendorDesc.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblVendorDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorDesc.Location = New System.Drawing.Point(3, 33)
        Me.lblVendorDesc.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblVendorDesc.Name = "lblVendorDesc"
        Me.lblVendorDesc.Size = New System.Drawing.Size(233, 27)
        Me.lblVendorDesc.TabIndex = 33
        Me.lblVendorDesc.Text = "  "
        '
        'txtVendorNo
        '
        Me.txtVendorNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorNo.Location = New System.Drawing.Point(3, 3)
        Me.txtVendorNo.Name = "txtVendorNo"
        Me.txtVendorNo.Size = New System.Drawing.Size(190, 25)
        Me.txtVendorNo.TabIndex = 30
        '
        'btnValidVendor
        '
        Me.btnValidVendor.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnValidVendor.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnValidVendor.ImageIndex = 1
        Me.btnValidVendor.ImageList = Me.ImageList1
        Me.btnValidVendor.Location = New System.Drawing.Point(199, 3)
        Me.btnValidVendor.Name = "btnValidVendor"
        Me.btnValidVendor.Size = New System.Drawing.Size(37, 24)
        Me.btnValidVendor.TabIndex = 31
        Me.btnValidVendor.Text = " "
        Me.btnValidVendor.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22.17573!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 77.82426!))
        Me.TableLayoutPanel3.Controls.Add(Me.cmdExcel, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.lblExcel, 1, 0)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 587)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(239, 28)
        Me.TableLayoutPanel3.TabIndex = 28
        '
        'cmdExcel
        '
        Me.cmdExcel.ImageIndex = 0
        Me.cmdExcel.ImageList = Me.ImageList1
        Me.cmdExcel.Location = New System.Drawing.Point(3, 3)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(38, 22)
        Me.cmdExcel.TabIndex = 29
        Me.cmdExcel.UseVisualStyleBackColor = True
        Me.cmdExcel.Visible = False
        '
        'lblExcel
        '
        Me.lblExcel.AutoSize = True
        Me.lblExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblExcel.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExcel.Location = New System.Drawing.Point(61, 12)
        Me.lblExcel.Margin = New System.Windows.Forms.Padding(8, 12, 3, 0)
        Me.lblExcel.Name = "lblExcel"
        Me.lblExcel.Size = New System.Drawing.Size(165, 12)
        Me.lblExcel.TabIndex = 30
        Me.lblExcel.Text = "Print Errors to Excel Document."
        Me.lblExcel.Visible = False
        '
        'SplitContainer1
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.SplitContainer1, 3)
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 311)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TableLayoutPanel6)
        Me.SplitContainer1.Panel1MinSize = 60
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TableLayoutPanel5)
        Me.SplitContainer1.Panel2Collapsed = True
        Me.SplitContainer1.Panel2MinSize = 60
        Me.SplitContainer1.Size = New System.Drawing.Size(769, 254)
        Me.SplitContainer1.SplitterDistance = 60
        Me.SplitContainer1.TabIndex = 27
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.ColumnCount = 3
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 274.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 251.0!))
        Me.TableLayoutPanel6.Controls.Add(Me.lblUsrLog, 0, 1)
        Me.TableLayoutPanel6.Controls.Add(Me.DataGridView1, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.BindingNavigator1, 1, 1)
        Me.TableLayoutPanel6.Controls.Add(Me.ProgressBar1, 0, 2)
        Me.TableLayoutPanel6.Controls.Add(Me.TableLayoutPanel9, 2, 1)
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 3
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 79.85612!))
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.14388!))
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(763, 278)
        Me.TableLayoutPanel6.TabIndex = 0
        '
        'lblUsrLog
        '
        Me.lblUsrLog.AutoSize = True
        Me.lblUsrLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsrLog.Location = New System.Drawing.Point(3, 214)
        Me.lblUsrLog.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.lblUsrLog.Name = "lblUsrLog"
        Me.lblUsrLog.Size = New System.Drawing.Size(91, 13)
        Me.lblUsrLog.TabIndex = 27
        Me.lblUsrLog.Text = " User Logged: "
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clPRHCOD, Me.clPRDPTN, Me.clVMVNUM, Me.clPRDCTP, Me.clPRDMFR, Me.clPQPRC, Me.clPQMIN, Me.clPRDSTS})
        Me.TableLayoutPanel6.SetColumnSpan(Me.DataGridView1, 3)
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(757, 200)
        Me.DataGridView1.TabIndex = 11
        '
        'clPRHCOD
        '
        Me.clPRHCOD.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRHCOD.FillWeight = 85.27919!
        Me.clPRHCOD.HeaderText = "Project No."
        Me.clPRHCOD.MinimumWidth = 8
        Me.clPRHCOD.Name = "clPRHCOD"
        Me.clPRHCOD.ReadOnly = True
        Me.clPRHCOD.Visible = False
        '
        'clPRDPTN
        '
        Me.clPRDPTN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDPTN.FillWeight = 102.9442!
        Me.clPRDPTN.HeaderText = "Part No."
        Me.clPRDPTN.MinimumWidth = 8
        Me.clPRDPTN.Name = "clPRDPTN"
        Me.clPRDPTN.ReadOnly = True
        '
        'clVMVNUM
        '
        Me.clVMVNUM.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clVMVNUM.FillWeight = 102.9442!
        Me.clVMVNUM.HeaderText = "Vendor No."
        Me.clVMVNUM.MinimumWidth = 8
        Me.clVMVNUM.Name = "clVMVNUM"
        Me.clVMVNUM.ReadOnly = True
        '
        'clPRDCTP
        '
        Me.clPRDCTP.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDCTP.HeaderText = "CTP No."
        Me.clPRDCTP.Name = "clPRDCTP"
        Me.clPRDCTP.ReadOnly = True
        '
        'clPRDMFR
        '
        Me.clPRDMFR.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDMFR.FillWeight = 102.9442!
        Me.clPRDMFR.HeaderText = "Manufacturer No."
        Me.clPRDMFR.MinimumWidth = 8
        Me.clPRDMFR.Name = "clPRDMFR"
        Me.clPRDMFR.ReadOnly = True
        '
        'clPQPRC
        '
        Me.clPQPRC.HeaderText = "Unit Cost"
        Me.clPQPRC.Name = "clPQPRC"
        Me.clPQPRC.ReadOnly = True
        Me.clPQPRC.Width = 69
        '
        'clPQMIN
        '
        Me.clPQMIN.HeaderText = "Min Qty"
        Me.clPQMIN.Name = "clPQMIN"
        Me.clPQMIN.ReadOnly = True
        Me.clPQMIN.Width = 49
        '
        'clPRDSTS
        '
        Me.clPRDSTS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDSTS.FillWeight = 102.9442!
        Me.clPRDSTS.HeaderText = "Status"
        Me.clPRDSTS.MinimumWidth = 8
        Me.clPRDSTS.Name = "clPRDSTS"
        Me.clPRDSTS.ReadOnly = True
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.AutoSize = False
        Me.BindingNavigator1.CountItem = Me.BindingNavigatorCountItem
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.None
        Me.BindingNavigator1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(248, 206)
        Me.BindingNavigator1.Margin = New System.Windows.Forms.Padding(10, 0, 0, 0)
        Me.BindingNavigator1.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.BindingNavigator1.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.BindingNavigator1.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.BindingNavigator1.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Me.BindingNavigatorPositionItem
        Me.BindingNavigator1.Size = New System.Drawing.Size(259, 36)
        Me.BindingNavigator1.TabIndex = 1
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 33)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(28, 33)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(28, 33)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 36)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(75, 35)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 36)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(28, 33)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(28, 33)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 36)
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 260)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(100, 15)
        Me.ProgressBar1.TabIndex = 25
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.ColumnCount = 2
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel9.Controls.Add(Me.dtProjectDate, 1, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.btnInsert, 0, 0)
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(515, 209)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 2
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(200, 45)
        Me.TableLayoutPanel9.TabIndex = 26
        '
        'dtProjectDate
        '
        Me.dtProjectDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtProjectDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtProjectDate.Location = New System.Drawing.Point(103, 3)
        Me.dtProjectDate.Name = "dtProjectDate"
        Me.dtProjectDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dtProjectDate.Size = New System.Drawing.Size(94, 24)
        Me.dtProjectDate.TabIndex = 24
        Me.dtProjectDate.Visible = False
        '
        'btnInsert
        '
        Me.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnInsert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInsert.Location = New System.Drawing.Point(70, 8)
        Me.btnInsert.Margin = New System.Windows.Forms.Padding(70, 8, 3, 3)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(27, 10)
        Me.btnInsert.TabIndex = 10
        Me.btnInsert.Text = "Insert to DB"
        Me.btnInsert.UseVisualStyleBackColor = True
        Me.btnInsert.Visible = False
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 3
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 232.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 299.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.DataGridView2, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.BindingNavigator2, 1, 1)
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 2
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 87.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(763, 248)
        Me.TableLayoutPanel5.TabIndex = 2
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.EditReference, Me.AddReference, Me.clPRDPTN2, Me.clVMVNUM2, Me.clPRDCTP2, Me.clPRDMFR2, Me.clPQPRC2, Me.clPQMIN2, Me.clError})
        Me.TableLayoutPanel5.SetColumnSpan(Me.DataGridView2, 3)
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridView2.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView2.Name = "DataGridView2"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(757, 211)
        Me.DataGridView2.TabIndex = 2
        Me.DataGridView2.Visible = False
        '
        'EditReference
        '
        Me.EditReference.HeaderText = "Edit"
        Me.EditReference.Name = "EditReference"
        Me.EditReference.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.EditReference.Visible = False
        Me.EditReference.Width = 31
        '
        'AddReference
        '
        Me.AddReference.HeaderText = "Add"
        Me.AddReference.Name = "AddReference"
        Me.AddReference.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.AddReference.Visible = False
        Me.AddReference.Width = 32
        '
        'clPRDPTN2
        '
        Me.clPRDPTN2.HeaderText = "Part Number"
        Me.clPRDPTN2.Name = "clPRDPTN2"
        Me.clPRDPTN2.Width = 84
        '
        'clVMVNUM2
        '
        Me.clVMVNUM2.HeaderText = "Vendor Number"
        Me.clVMVNUM2.Name = "clVMVNUM2"
        Me.clVMVNUM2.Width = 97
        '
        'clPRDCTP2
        '
        Me.clPRDCTP2.HeaderText = "CTP No."
        Me.clPRDCTP2.Name = "clPRDCTP2"
        Me.clPRDCTP2.Width = 68
        '
        'clPRDMFR2
        '
        Me.clPRDMFR2.HeaderText = "Manufacturer No."
        Me.clPRDMFR2.Name = "clPRDMFR2"
        Me.clPRDMFR2.Width = 105
        '
        'clPQPRC2
        '
        Me.clPQPRC2.HeaderText = "Unit Cost"
        Me.clPQPRC2.Name = "clPQPRC2"
        Me.clPQPRC2.Width = 69
        '
        'clPQMIN2
        '
        Me.clPQMIN2.HeaderText = "Min Qty"
        Me.clPQMIN2.Name = "clPQMIN2"
        Me.clPQMIN2.Width = 49
        '
        'clError
        '
        Me.clError.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clError.HeaderText = "Error Description"
        Me.clError.Name = "clError"
        '
        'BindingNavigator2
        '
        Me.BindingNavigator2.AddNewItem = Nothing
        Me.BindingNavigator2.CountItem = Me.BindingNavigatorCountItem1
        Me.BindingNavigator2.DeleteItem = Nothing
        Me.BindingNavigator2.Dock = System.Windows.Forms.DockStyle.None
        Me.BindingNavigator2.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.BindingNavigator2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem1, Me.BindingNavigatorMovePreviousItem1, Me.BindingNavigatorSeparator3, Me.BindingNavigatorPositionItem1, Me.BindingNavigatorCountItem1, Me.BindingNavigatorSeparator4, Me.BindingNavigatorMoveNextItem1, Me.BindingNavigatorMoveLastItem1, Me.BindingNavigatorSeparator5})
        Me.BindingNavigator2.Location = New System.Drawing.Point(232, 217)
        Me.BindingNavigator2.MoveFirstItem = Me.BindingNavigatorMoveFirstItem1
        Me.BindingNavigator2.MoveLastItem = Me.BindingNavigatorMoveLastItem1
        Me.BindingNavigator2.MoveNextItem = Me.BindingNavigatorMoveNextItem1
        Me.BindingNavigator2.MovePreviousItem = Me.BindingNavigatorMovePreviousItem1
        Me.BindingNavigator2.Name = "BindingNavigator2"
        Me.BindingNavigator2.PositionItem = Me.BindingNavigatorPositionItem1
        Me.BindingNavigator2.Size = New System.Drawing.Size(229, 31)
        Me.BindingNavigator2.TabIndex = 1
        Me.BindingNavigator2.Text = "BindingNavigator2"
        '
        'BindingNavigatorCountItem1
        '
        Me.BindingNavigatorCountItem1.Name = "BindingNavigatorCountItem1"
        Me.BindingNavigatorCountItem1.Size = New System.Drawing.Size(35, 28)
        Me.BindingNavigatorCountItem1.Text = "of {0}"
        Me.BindingNavigatorCountItem1.ToolTipText = "Total number of items"
        '
        'BindingNavigatorMoveFirstItem1
        '
        Me.BindingNavigatorMoveFirstItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem1.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem1.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem1.Name = "BindingNavigatorMoveFirstItem1"
        Me.BindingNavigatorMoveFirstItem1.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem1.Size = New System.Drawing.Size(28, 28)
        Me.BindingNavigatorMoveFirstItem1.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem1
        '
        Me.BindingNavigatorMovePreviousItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem1.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem1.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem1.Name = "BindingNavigatorMovePreviousItem1"
        Me.BindingNavigatorMovePreviousItem1.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem1.Size = New System.Drawing.Size(28, 28)
        Me.BindingNavigatorMovePreviousItem1.Text = "Move previous"
        '
        'BindingNavigatorSeparator3
        '
        Me.BindingNavigatorSeparator3.Name = "BindingNavigatorSeparator3"
        Me.BindingNavigatorSeparator3.Size = New System.Drawing.Size(6, 31)
        '
        'BindingNavigatorPositionItem1
        '
        Me.BindingNavigatorPositionItem1.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem1.AutoSize = False
        Me.BindingNavigatorPositionItem1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.BindingNavigatorPositionItem1.Name = "BindingNavigatorPositionItem1"
        Me.BindingNavigatorPositionItem1.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem1.Text = "0"
        Me.BindingNavigatorPositionItem1.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator4
        '
        Me.BindingNavigatorSeparator4.Name = "BindingNavigatorSeparator4"
        Me.BindingNavigatorSeparator4.Size = New System.Drawing.Size(6, 31)
        '
        'BindingNavigatorMoveNextItem1
        '
        Me.BindingNavigatorMoveNextItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem1.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem1.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem1.Name = "BindingNavigatorMoveNextItem1"
        Me.BindingNavigatorMoveNextItem1.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem1.Size = New System.Drawing.Size(28, 28)
        Me.BindingNavigatorMoveNextItem1.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem1
        '
        Me.BindingNavigatorMoveLastItem1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem1.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem1.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem1.Name = "BindingNavigatorMoveLastItem1"
        Me.BindingNavigatorMoveLastItem1.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem1.Size = New System.Drawing.Size(28, 28)
        Me.BindingNavigatorMoveLastItem1.Text = "Move last"
        '
        'BindingNavigatorSeparator5
        '
        Me.BindingNavigatorSeparator5.Name = "BindingNavigatorSeparator5"
        Me.BindingNavigatorSeparator5.Size = New System.Drawing.Size(6, 31)
        '
        'cmbStatus
        '
        Me.cmbStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.IntegralHeight = False
        Me.cmbStatus.ItemHeight = 17
        Me.cmbStatus.Location = New System.Drawing.Point(493, 195)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(279, 25)
        Me.cmbStatus.TabIndex = 26
        Me.cmbStatus.Visible = False
        '
        'cmbPerCharge
        '
        Me.cmbPerCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPerCharge.FormattingEnabled = True
        Me.cmbPerCharge.Location = New System.Drawing.Point(3, 130)
        Me.cmbPerCharge.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.cmbPerCharge.Name = "cmbPerCharge"
        Me.cmbPerCharge.Size = New System.Drawing.Size(239, 25)
        Me.cmbPerCharge.TabIndex = 25
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TableLayoutPanel2.SetColumnSpan(Me.Panel2, 3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(769, 44)
        Me.Panel2.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(85, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(604, 25)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Please add the project header data and then add the excel with the references."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 267)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(239, 41)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please select the file that you want to process"
        '
        'btnSelect
        '
        Me.btnSelect.Enabled = False
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(315, 267)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(70, 10, 3, 3)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(147, 38)
        Me.btnSelect.TabIndex = 0
        Me.btnSelect.Text = "Load File"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TableLayoutPanel7)
        Me.Panel1.Location = New System.Drawing.Point(493, 260)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(279, 45)
        Me.Panel1.TabIndex = 2
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.ColumnCount = 3
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 71.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 58.0!))
        Me.TableLayoutPanel7.Controls.Add(Me.ComboBox1, 1, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.LinkLabel1, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.LinkLabel4, 2, 0)
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(0, 7)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 1
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(273, 36)
        Me.TableLayoutPanel7.TabIndex = 0
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(147, 6)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(48, 25)
        Me.ComboBox1.TabIndex = 2
        Me.ComboBox1.Visible = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LinkLabel1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LinkLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LinkLabel1.ImageIndex = 1
        Me.LinkLabel1.ImageList = Me.ImageList1
        Me.LinkLabel1.LinkColor = System.Drawing.Color.Black
        Me.LinkLabel1.Location = New System.Drawing.Point(30, 5)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(30, 5, 0, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Padding = New System.Windows.Forms.Padding(5, 5, 30, 5)
        Me.LinkLabel1.Size = New System.Drawing.Size(98, 25)
        Me.LinkLabel1.TabIndex = 37
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Clear Filters"
        '
        'LinkLabel4
        '
        Me.LinkLabel4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LinkLabel4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LinkLabel4.ImageIndex = 12
        Me.LinkLabel4.ImageList = Me.ImageList1
        Me.LinkLabel4.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.LinkLabel4.Location = New System.Drawing.Point(233, 9)
        Me.LinkLabel4.Margin = New System.Windows.Forms.Padding(18, 9, 0, 0)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(28, 20)
        Me.LinkLabel4.TabIndex = 38
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "        "
        '
        'ComboBox2
        '
        Me.ComboBox2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(3, 34)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(273, 25)
        Me.ComboBox2.TabIndex = 15
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnSuccess, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCheck, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(248, 587)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(239, 49)
        Me.TableLayoutPanel1.TabIndex = 18
        '
        'btnSuccess
        '
        Me.btnSuccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSuccess.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSuccess.Location = New System.Drawing.Point(3, 6)
        Me.btnSuccess.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.btnSuccess.Name = "btnSuccess"
        Me.btnSuccess.Size = New System.Drawing.Size(113, 39)
        Me.btnSuccess.TabIndex = 4
        Me.btnSuccess.Text = "Show Success"
        Me.btnSuccess.UseVisualStyleBackColor = True
        '
        'btnCheck
        '
        Me.btnCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheck.Location = New System.Drawing.Point(122, 6)
        Me.btnCheck.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.btnCheck.Name = "btnCheck"
        Me.btnCheck.Size = New System.Drawing.Size(114, 39)
        Me.btnCheck.TabIndex = 3
        Me.btnCheck.Text = "Check Errors"
        Me.btnCheck.UseVisualStyleBackColor = True
        '
        'txtProjectName
        '
        Me.txtProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectName.Location = New System.Drawing.Point(248, 75)
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(239, 25)
        Me.txtProjectName.TabIndex = 12
        '
        'txtProjectNo
        '
        Me.txtProjectNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectNo.Location = New System.Drawing.Point(3, 75)
        Me.txtProjectNo.Name = "txtProjectNo"
        Me.txtProjectNo.Size = New System.Drawing.Size(239, 25)
        Me.txtProjectNo.TabIndex = 11
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.TableLayoutPanel2.SetColumnSpan(Me.lblDesc, 2)
        Me.lblDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesc.Location = New System.Drawing.Point(3, 196)
        Me.lblDesc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 0)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(80, 15)
        Me.lblDesc.TabIndex = 8
        Me.lblDesc.Text = "Description"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(493, 53)
        Me.lblStatus.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(116, 15)
        Me.lblStatus.TabIndex = 7
        Me.lblStatus.Text = "Project Status (*)"
        '
        'lblPerCharge
        '
        Me.lblPerCharge.AutoSize = True
        Me.lblPerCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPerCharge.Location = New System.Drawing.Point(3, 104)
        Me.lblPerCharge.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblPerCharge.Name = "lblPerCharge"
        Me.lblPerCharge.Size = New System.Drawing.Size(118, 15)
        Me.lblPerCharge.TabIndex = 6
        Me.lblPerCharge.Text = "Person in Charge"
        '
        'lblProjectName
        '
        Me.lblProjectName.AutoSize = True
        Me.lblProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectName.Location = New System.Drawing.Point(248, 53)
        Me.lblProjectName.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(114, 15)
        Me.lblProjectName.TabIndex = 4
        Me.lblProjectName.Text = "Project Name (*)"
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectNo.Location = New System.Drawing.Point(3, 53)
        Me.lblProjectNo.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(78, 15)
        Me.lblProjectNo.TabIndex = 3
        Me.lblProjectNo.Text = "Project No."
        '
        'lblVendorNo
        '
        Me.lblVendorNo.AutoSize = True
        Me.lblVendorNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorNo.Location = New System.Drawing.Point(248, 104)
        Me.lblVendorNo.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblVendorNo.Name = "lblVendorNo"
        Me.lblVendorNo.Size = New System.Drawing.Size(127, 15)
        Me.lblVendorNo.TabIndex = 30
        Me.lblVendorNo.Text = "Vendor Number (*)"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.BackColor = System.Drawing.SystemColors.Control
        Me.TableLayoutPanel2.ColumnCount = 3
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 285.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.lblVendorNo, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectNo, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectName, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblPerCharge, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel10, 2, 10)
        Me.TableLayoutPanel2.Controls.Add(Me.lblDesc, 0, 5)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectNo, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectName, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtDesc, 0, 6)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel1, 2, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.btnSelect, 1, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbPerCharge, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.SplitContainer1, 0, 8)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel3, 0, 10)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel4, 1, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.lblStatus, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel8, 2, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbStatus, 2, 5)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbStatusMore, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel1, 1, 10)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(13, 21)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 11
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 68.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 51.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 267.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 9.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(775, 647)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'TableLayoutPanel10
        '
        Me.TableLayoutPanel10.ColumnCount = 2
        Me.TableLayoutPanel10.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel10.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 129.0!))
        Me.TableLayoutPanel10.Controls.Add(Me.LinkLabel2, 1, 0)
        Me.TableLayoutPanel10.Controls.Add(Me.LinkLabel3, 0, 0)
        Me.TableLayoutPanel10.Location = New System.Drawing.Point(493, 587)
        Me.TableLayoutPanel10.Name = "TableLayoutPanel10"
        Me.TableLayoutPanel10.RowCount = 2
        Me.TableLayoutPanel10.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel10.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8.0!))
        Me.TableLayoutPanel10.Size = New System.Drawing.Size(273, 58)
        Me.TableLayoutPanel10.TabIndex = 27
        '
        'LinkLabel2
        '
        Me.LinkLabel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LinkLabel2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LinkLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LinkLabel2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LinkLabel2.ImageIndex = 4
        Me.LinkLabel2.ImageList = Me.ImageList1
        Me.LinkLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.LinkLabel2.LinkColor = System.Drawing.Color.Black
        Me.LinkLabel2.Location = New System.Drawing.Point(144, 7)
        Me.LinkLabel2.Margin = New System.Windows.Forms.Padding(0, 7, 3, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Padding = New System.Windows.Forms.Padding(5, 5, 20, 5)
        Me.LinkLabel2.Size = New System.Drawing.Size(126, 43)
        Me.LinkLabel2.TabIndex = 38
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Go to Product Development   "
        '
        'LinkLabel3
        '
        Me.LinkLabel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LinkLabel3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LinkLabel3.Enabled = False
        Me.LinkLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.LinkLabel3.ImageIndex = 7
        Me.LinkLabel3.ImageList = Me.ImageList1
        Me.LinkLabel3.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.LinkLabel3.LinkColor = System.Drawing.Color.Black
        Me.LinkLabel3.Location = New System.Drawing.Point(0, 7)
        Me.LinkLabel3.Margin = New System.Windows.Forms.Padding(0, 7, 3, 0)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Padding = New System.Windows.Forms.Padding(5, 5, 20, 5)
        Me.LinkLabel3.Size = New System.Drawing.Size(138, 43)
        Me.LinkLabel3.TabIndex = 39
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Download Excel Template"
        '
        'txtDesc
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.txtDesc, 3)
        Me.txtDesc.Location = New System.Drawing.Point(3, 225)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesc.Size = New System.Drawing.Size(766, 29)
        Me.txtDesc.TabIndex = 16
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(493, 104)
        Me.Label3.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(114, 15)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Vendor Name (*)"
        '
        'TableLayoutPanel8
        '
        Me.TableLayoutPanel8.ColumnCount = 1
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.Controls.Add(Me.ComboBox2, 0, 1)
        Me.TableLayoutPanel8.Controls.Add(Me.ac2, 0, 0)
        Me.TableLayoutPanel8.Location = New System.Drawing.Point(493, 127)
        Me.TableLayoutPanel8.Name = "TableLayoutPanel8"
        Me.TableLayoutPanel8.RowCount = 2
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.Size = New System.Drawing.Size(279, 62)
        Me.TableLayoutPanel8.TabIndex = 35
        '
        'ac2
        '
        Me.ac2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ac2.Location = New System.Drawing.Point(3, 3)
        Me.ac2.lstSelectedValues = CType(resources.GetObject("ac2.lstSelectedValues"), System.Collections.Generic.List(Of String))
        Me.ac2.Name = "ac2"
        Me.ac2.Size = New System.Drawing.Size(273, 25)
        Me.ac2.TabIndex = 16
        Me.ac2.Values = Nothing
        '
        'cmbStatusMore
        '
        Me.cmbStatusMore.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbStatusMore.FormattingEnabled = True
        Me.cmbStatusMore.Location = New System.Drawing.Point(493, 75)
        Me.cmbStatusMore.Name = "cmbStatusMore"
        Me.cmbStatusMore.Size = New System.Drawing.Size(279, 25)
        Me.cmbStatusMore.TabIndex = 37
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        Me.BackgroundWorker2.WorkerSupportsCancellation = True
        '
        'BackgroundWorker4
        '
        Me.BackgroundWorker4.WorkerSupportsCancellation = True
        '
        'ToolTip1
        '
        Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.ToolTip1.ToolTipTitle = "Click Here!!"
        '
        'frmLoadExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(794, 701)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Name = "frmLoadExcel"
        Me.Text = "frmLoadExcel"
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.TableLayoutPanel6.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingNavigator2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator2.ResumeLayout(False)
        Me.BindingNavigator2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TableLayoutPanel10.ResumeLayout(False)
        Me.TableLayoutPanel8.ResumeLayout(False)
        Me.TableLayoutPanel8.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents BindingSource1 As BindingSource
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents txtVendorNo As TextBox
    Friend WithEvents btnValidVendor As Button
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents cmdExcel As Button
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents lblVendorNo As Label
    Friend WithEvents lblProjectNo As Label
    Friend WithEvents lblProjectName As Label
    Friend WithEvents lblPerCharge As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblDesc As Label
    Friend WithEvents btnInsert As Button
    Friend WithEvents txtProjectNo As TextBox
    Friend WithEvents txtProjectName As TextBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents btnSuccess As Button
    Friend WithEvents btnCheck As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnSelect As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbPerCharge As ComboBox
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents BindingNavigator2 As BindingNavigator
    Friend WithEvents BindingNavigatorCountItem1 As ToolStripLabel
    Friend WithEvents BindingNavigatorMoveFirstItem1 As ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem1 As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator3 As ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem1 As ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator4 As ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem1 As ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem1 As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator5 As ToolStripSeparator
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents lblExcel As Label
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Label3 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ac1 As Autocomplete_Textbox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents TableLayoutPanel8 As TableLayoutPanel
    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents cmbStatusMore As ComboBox
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents dtProjectDate As DateTimePicker
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents BindingNavigator1 As BindingNavigator
    Friend WithEvents BindingNavigatorCountItem As ToolStripLabel
    Friend WithEvents BindingNavigatorMoveFirstItem As ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As ToolStripSeparator
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents TableLayoutPanel9 As TableLayoutPanel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ac2 As Autocomplete_Textbox
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents EditReference As DataGridViewLinkColumn
    Friend WithEvents AddReference As DataGridViewLinkColumn
    Friend WithEvents clPRDPTN2 As DataGridViewTextBoxColumn
    Friend WithEvents clVMVNUM2 As DataGridViewTextBoxColumn
    Friend WithEvents clPRDCTP2 As DataGridViewTextBoxColumn
    Friend WithEvents clPRDMFR2 As DataGridViewTextBoxColumn
    Friend WithEvents clPQPRC2 As DataGridViewTextBoxColumn
    Friend WithEvents clPQMIN2 As DataGridViewTextBoxColumn
    Friend WithEvents clError As DataGridViewTextBoxColumn
    Friend WithEvents clPRHCOD As DataGridViewTextBoxColumn
    Friend WithEvents clPRDPTN As DataGridViewTextBoxColumn
    Friend WithEvents clVMVNUM As DataGridViewTextBoxColumn
    Friend WithEvents clPRDCTP As DataGridViewTextBoxColumn
    Friend WithEvents clPRDMFR As DataGridViewTextBoxColumn
    Friend WithEvents clPQPRC As DataGridViewTextBoxColumn
    Friend WithEvents clPQMIN As DataGridViewTextBoxColumn
    Friend WithEvents clPRDSTS As DataGridViewTextBoxColumn
    Friend WithEvents lblVendorDesc As Label
    Friend WithEvents LinkLabel2 As LinkLabel
    Friend WithEvents TableLayoutPanel10 As TableLayoutPanel
    Friend WithEvents lblUsrLog As Label
    Friend WithEvents LinkLabel3 As LinkLabel
    Friend WithEvents LinkLabel4 As LinkLabel
    Friend WithEvents ToolTip1 As ToolTip
End Class
