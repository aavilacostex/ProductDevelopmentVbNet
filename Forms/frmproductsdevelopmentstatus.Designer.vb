<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmproductsdevelopmentstatus
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmproductsdevelopmentstatus))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdSave1 = New System.Windows.Forms.Button()
        Me.cmdexit1 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblproject = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbstatus = New System.Windows.Forms.ComboBox()
        Me.clPartNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clCtpNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVendorNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clDescription = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clStatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clPartNo, Me.clCtpNo, Me.clVendorNo, Me.clDescription, Me.clStatus})
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.Menu
        Me.DataGridView1.Location = New System.Drawing.Point(3, 78)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(695, 263)
        Me.DataGridView1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.GrayText
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.cmdSave1)
        Me.Panel2.Controls.Add(Me.cmdexit1)
        Me.Panel2.Location = New System.Drawing.Point(18, 369)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(695, 34)
        Me.Panel2.TabIndex = 7
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(494, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 28)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Select All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmdSave1
        '
        Me.cmdSave1.Image = CType(resources.GetObject("cmdSave1.Image"), System.Drawing.Image)
        Me.cmdSave1.Location = New System.Drawing.Point(578, 3)
        Me.cmdSave1.Name = "cmdSave1"
        Me.cmdSave1.Size = New System.Drawing.Size(43, 28)
        Me.cmdSave1.TabIndex = 7
        Me.cmdSave1.UseVisualStyleBackColor = True
        '
        'cmdexit1
        '
        Me.cmdexit1.Image = CType(resources.GetObject("cmdexit1.Image"), System.Drawing.Image)
        Me.cmdexit1.Location = New System.Drawing.Point(630, 3)
        Me.cmdexit1.Name = "cmdexit1"
        Me.cmdexit1.Size = New System.Drawing.Size(43, 28)
        Me.cmdexit1.TabIndex = 8
        Me.cmdexit1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(707, 352)
        Me.Panel1.TabIndex = 6
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblproject, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 8)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.135718!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.23597!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 77.62831!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(701, 344)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'lblproject
        '
        Me.lblproject.AutoSize = True
        Me.lblproject.Location = New System.Drawing.Point(10, 10)
        Me.lblproject.Margin = New System.Windows.Forms.Padding(10, 10, 3, 0)
        Me.lblproject.Name = "lblproject"
        Me.lblproject.Size = New System.Drawing.Size(39, 13)
        Me.lblproject.TabIndex = 0
        Me.lblproject.Text = "Label1"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.48148!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 78.51852!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbstatus, 1, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 30)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(405, 34)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 10)
        Me.Label2.Margin = New System.Windows.Forms.Padding(8, 10, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "New Status"
        '
        'cmbstatus
        '
        Me.cmbstatus.FormattingEnabled = True
        Me.cmbstatus.Location = New System.Drawing.Point(89, 6)
        Me.cmbstatus.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.cmbstatus.Name = "cmbstatus"
        Me.cmbstatus.Size = New System.Drawing.Size(258, 21)
        Me.cmbstatus.TabIndex = 1
        '
        'clPartNo
        '
        Me.clPartNo.HeaderText = "Part No."
        Me.clPartNo.Name = "clPartNo"
        '
        'clCtpNo
        '
        Me.clCtpNo.HeaderText = "CTP No."
        Me.clCtpNo.Name = "clCtpNo"
        '
        'clVendorNo
        '
        Me.clVendorNo.HeaderText = "MFR No."
        Me.clVendorNo.Name = "clVendorNo"
        '
        'clDescription
        '
        Me.clDescription.HeaderText = "Description"
        Me.clDescription.Name = "clDescription"
        '
        'clStatus
        '
        Me.clStatus.HeaderText = "Status"
        Me.clStatus.Name = "clStatus"
        '
        'frmproductsdevelopmentstatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(728, 450)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmproductsdevelopmentstatus"
        Me.Text = "frmproductsdevelopmentstatus"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Panel2 As Panel
    Friend WithEvents cmdSave1 As Button
    Friend WithEvents cmdexit1 As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lblproject As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbstatus As ComboBox
    Friend WithEvents clPartNo As DataGridViewTextBoxColumn
    Friend WithEvents clCtpNo As DataGridViewTextBoxColumn
    Friend WithEvents clVendorNo As DataGridViewTextBoxColumn
    Friend WithEvents clDescription As DataGridViewTextBoxColumn
    Friend WithEvents clStatus As DataGridViewTextBoxColumn
End Class
