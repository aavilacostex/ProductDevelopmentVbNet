<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class test
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
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle22 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle23 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle24 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.clPRHCOD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDPTN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDCTP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDMFR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVMVNUM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDSTS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.lblProjectNo = New System.Windows.Forms.Label()
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.lblProjectDate = New System.Windows.Forms.Label()
        Me.lblPerCharge = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.txtProjectNo = New System.Windows.Forms.TextBox()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnSuccess = New System.Windows.Forms.Button()
        Me.btnCheck = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtProjectDate = New System.Windows.Forms.DateTimePicker()
        Me.cmbPerCharge = New System.Windows.Forms.ComboBox()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TableLayoutPanel2)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(912, 600)
        Me.Panel1.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 3
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 286.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.SplitContainer1, 0, 8)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectNo, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectName, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectDate, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblPerCharge, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.lblStatus, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.lblDesc, 0, 5)
        Me.TableLayoutPanel2.Controls.Add(Me.btnInsert, 2, 9)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectNo, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectName, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtDesc, 0, 6)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel1, 1, 9)
        Me.TableLayoutPanel2.Controls.Add(Me.lblMessage, 0, 9)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel2, 2, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.btnSelect, 1, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel3, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.dtProjectDate, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbPerCharge, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbStatus, 2, 4)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(7, 12)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 10
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 53.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 290.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 176.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(902, 582)
        Me.TableLayoutPanel2.TabIndex = 4
        '
        'SplitContainer1
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.SplitContainer1, 3)
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 247)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.AutoScroll = True
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView1)
        Me.SplitContainer1.Panel1MinSize = 60
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.AutoScroll = True
        Me.SplitContainer1.Panel2.Controls.Add(Me.DataGridView2)
        Me.SplitContainer1.Panel2Collapsed = True
        Me.SplitContainer1.Panel2MinSize = 60
        Me.SplitContainer1.Size = New System.Drawing.Size(896, 284)
        Me.SplitContainer1.SplitterDistance = 60
        Me.SplitContainer1.TabIndex = 27
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle19.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle19.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle19.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle19.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle19.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle19
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clPRHCOD, Me.clPRDPTN, Me.clPRDCTP, Me.clPRDMFR, Me.clVMVNUM, Me.clPRDSTS})
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle20
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle21
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(890, 278)
        Me.DataGridView1.TabIndex = 11
        '
        'clPRHCOD
        '
        Me.clPRHCOD.HeaderText = "Project No."
        Me.clPRHCOD.MinimumWidth = 8
        Me.clPRHCOD.Name = "clPRHCOD"
        Me.clPRHCOD.ReadOnly = True
        Me.clPRHCOD.Width = 150
        '
        'clPRDPTN
        '
        Me.clPRDPTN.HeaderText = "Part No."
        Me.clPRDPTN.MinimumWidth = 8
        Me.clPRDPTN.Name = "clPRDPTN"
        Me.clPRDPTN.ReadOnly = True
        Me.clPRDPTN.Width = 150
        '
        'clPRDCTP
        '
        Me.clPRDCTP.HeaderText = "CTP No."
        Me.clPRDCTP.MinimumWidth = 8
        Me.clPRDCTP.Name = "clPRDCTP"
        Me.clPRDCTP.ReadOnly = True
        Me.clPRDCTP.Width = 150
        '
        'clPRDMFR
        '
        Me.clPRDMFR.HeaderText = "Manufacturer No."
        Me.clPRDMFR.MinimumWidth = 8
        Me.clPRDMFR.Name = "clPRDMFR"
        Me.clPRDMFR.ReadOnly = True
        Me.clPRDMFR.Width = 150
        '
        'clVMVNUM
        '
        Me.clVMVNUM.HeaderText = "Vendor No."
        Me.clVMVNUM.MinimumWidth = 8
        Me.clVMVNUM.Name = "clVMVNUM"
        Me.clVMVNUM.ReadOnly = True
        Me.clVMVNUM.Width = 150
        '
        'clPRDSTS
        '
        Me.clPRDSTS.HeaderText = "Status"
        Me.clPRDSTS.MinimumWidth = 8
        Me.clPRDSTS.Name = "clPRDSTS"
        Me.clPRDSTS.ReadOnly = True
        Me.clPRDSTS.Width = 150
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToOrderColumns = True
        DataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle22.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle22.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle22.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle22.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle22.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle22.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle22
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2})
        DataGridViewCellStyle23.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle23.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle23.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle23.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle23.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle23.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle23
        Me.DataGridView2.Location = New System.Drawing.Point(-3, 3)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        DataGridViewCellStyle24.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle24.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle24.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle24.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle24.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle24.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.RowHeadersDefaultCellStyle = DataGridViewCellStyle24
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowHeadersWidth = 62
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(899, 279)
        Me.DataGridView2.TabIndex = 12
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectNo.Location = New System.Drawing.Point(3, 37)
        Me.lblProjectNo.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(78, 15)
        Me.lblProjectNo.TabIndex = 3
        Me.lblProjectNo.Text = "Project No."
        '
        'lblProjectName
        '
        Me.lblProjectName.AutoSize = True
        Me.lblProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectName.Location = New System.Drawing.Point(311, 37)
        Me.lblProjectName.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(94, 15)
        Me.lblProjectName.TabIndex = 4
        Me.lblProjectName.Text = "Project Name"
        '
        'lblProjectDate
        '
        Me.lblProjectDate.AutoSize = True
        Me.lblProjectDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectDate.Location = New System.Drawing.Point(619, 37)
        Me.lblProjectDate.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblProjectDate.Name = "lblProjectDate"
        Me.lblProjectDate.Size = New System.Drawing.Size(86, 15)
        Me.lblProjectDate.TabIndex = 5
        Me.lblProjectDate.Text = "Project Date"
        '
        'lblPerCharge
        '
        Me.lblPerCharge.AutoSize = True
        Me.TableLayoutPanel2.SetColumnSpan(Me.lblPerCharge, 2)
        Me.lblPerCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPerCharge.Location = New System.Drawing.Point(3, 90)
        Me.lblPerCharge.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblPerCharge.Name = "lblPerCharge"
        Me.lblPerCharge.Size = New System.Drawing.Size(118, 15)
        Me.lblPerCharge.TabIndex = 6
        Me.lblPerCharge.Text = "Person in Charge"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(619, 90)
        Me.lblStatus.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(96, 15)
        Me.lblStatus.TabIndex = 7
        Me.lblStatus.Text = "Project Status"
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.TableLayoutPanel2.SetColumnSpan(Me.lblDesc, 3)
        Me.lblDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesc.Location = New System.Drawing.Point(3, 142)
        Me.lblDesc.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(80, 15)
        Me.lblDesc.TabIndex = 8
        Me.lblDesc.Text = "Description"
        '
        'btnInsert
        '
        Me.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnInsert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInsert.Location = New System.Drawing.Point(686, 542)
        Me.btnInsert.Margin = New System.Windows.Forms.Padding(70, 8, 3, 3)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(107, 30)
        Me.btnInsert.TabIndex = 10
        Me.btnInsert.Text = "Insert to DB"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'txtProjectNo
        '
        Me.txtProjectNo.Enabled = False
        Me.txtProjectNo.Location = New System.Drawing.Point(3, 61)
        Me.txtProjectNo.Multiline = True
        Me.txtProjectNo.Name = "txtProjectNo"
        Me.txtProjectNo.Size = New System.Drawing.Size(231, 23)
        Me.txtProjectNo.TabIndex = 11
        '
        'txtProjectName
        '
        Me.txtProjectName.Location = New System.Drawing.Point(311, 61)
        Me.txtProjectName.Multiline = True
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(231, 23)
        Me.txtProjectName.TabIndex = 12
        '
        'txtDesc
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.txtDesc, 3)
        Me.txtDesc.Location = New System.Drawing.Point(3, 163)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesc.Size = New System.Drawing.Size(721, 24)
        Me.txtDesc.TabIndex = 16
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnSuccess, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCheck, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(311, 537)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(239, 40)
        Me.TableLayoutPanel1.TabIndex = 18
        '
        'btnSuccess
        '
        Me.btnSuccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSuccess.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSuccess.Location = New System.Drawing.Point(3, 6)
        Me.btnSuccess.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.btnSuccess.Name = "btnSuccess"
        Me.btnSuccess.Size = New System.Drawing.Size(113, 30)
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
        Me.btnCheck.Size = New System.Drawing.Size(114, 30)
        Me.btnCheck.TabIndex = 3
        Me.btnCheck.Text = "Check Errors"
        Me.btnCheck.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.Location = New System.Drawing.Point(3, 539)
        Me.lblMessage.Margin = New System.Windows.Forms.Padding(3, 5, 3, 0)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(239, 30)
        Me.lblMessage.TabIndex = 17
        Me.lblMessage.Text = "Label2"
        Me.lblMessage.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.RadioButton2)
        Me.Panel2.Controls.Add(Me.RadioButton1)
        Me.Panel2.Location = New System.Drawing.Point(619, 194)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(200, 44)
        Me.Panel2.TabIndex = 2
        Me.Panel2.Visible = False
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(3, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "RadioButton2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(95, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(378, 201)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(70, 10, 3, 3)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(98, 40)
        Me.btnSelect.TabIndex = 0
        Me.btnSelect.Text = "Load File"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 201)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(231, 43)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please select the file that you want to process"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TableLayoutPanel2.SetColumnSpan(Me.Panel3, 3)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(3, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(896, 26)
        Me.Panel3.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(107, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(604, 25)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Please add the project header data and then add the excel with the references."
        '
        'dtProjectDate
        '
        Me.dtProjectDate.Location = New System.Drawing.Point(619, 61)
        Me.dtProjectDate.Name = "dtProjectDate"
        Me.dtProjectDate.Size = New System.Drawing.Size(231, 20)
        Me.dtProjectDate.TabIndex = 24
        '
        'cmbPerCharge
        '
        Me.cmbPerCharge.FormattingEnabled = True
        Me.cmbPerCharge.Location = New System.Drawing.Point(3, 113)
        Me.cmbPerCharge.Name = "cmbPerCharge"
        Me.cmbPerCharge.Size = New System.Drawing.Size(231, 21)
        Me.cmbPerCharge.TabIndex = 25
        '
        'cmbStatus
        '
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Location = New System.Drawing.Point(619, 113)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(231, 21)
        Me.cmbStatus.TabIndex = 26
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Project No."
        Me.DataGridViewTextBoxColumn1.MinimumWidth = 8
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 150
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Part No."
        Me.DataGridViewTextBoxColumn2.MinimumWidth = 8
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 400
        '
        'test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(932, 623)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "test"
        Me.Text = "test"
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents lblProjectNo As Label
    Friend WithEvents lblProjectName As Label
    Friend WithEvents lblProjectDate As Label
    Friend WithEvents lblPerCharge As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblDesc As Label
    Friend WithEvents btnInsert As Button
    Friend WithEvents txtProjectNo As TextBox
    Friend WithEvents txtProjectName As TextBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents btnSuccess As Button
    Friend WithEvents btnCheck As Button
    Friend WithEvents lblMessage As Label
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents btnSelect As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents dtProjectDate As DateTimePicker
    Friend WithEvents cmbPerCharge As ComboBox
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents clPRHCOD As DataGridViewTextBoxColumn
    Friend WithEvents clPRDPTN As DataGridViewTextBoxColumn
    Friend WithEvents clPRDCTP As DataGridViewTextBoxColumn
    Friend WithEvents clPRDMFR As DataGridViewTextBoxColumn
    Friend WithEvents clVMVNUM As DataGridViewTextBoxColumn
    Friend WithEvents clPRDSTS As DataGridViewTextBoxColumn
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
End Class
