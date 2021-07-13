<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmproductsdevelopmentmanu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmproductsdevelopmentmanu))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdSave1 = New System.Windows.Forms.Button()
        Me.cmdexit1 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.clPartNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clDescription = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clMfrNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblproject = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.GrayText
        Me.Panel2.Controls.Add(Me.cmdSave1)
        Me.Panel2.Controls.Add(Me.cmdexit1)
        Me.Panel2.Location = New System.Drawing.Point(18, 368)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(582, 34)
        Me.Panel2.TabIndex = 3
        '
        'cmdSave1
        '
        Me.cmdSave1.Image = CType(resources.GetObject("cmdSave1.Image"), System.Drawing.Image)
        Me.cmdSave1.Location = New System.Drawing.Point(471, 0)
        Me.cmdSave1.Name = "cmdSave1"
        Me.cmdSave1.Size = New System.Drawing.Size(43, 28)
        Me.cmdSave1.TabIndex = 7
        Me.cmdSave1.UseVisualStyleBackColor = True
        '
        'cmdexit1
        '
        Me.cmdexit1.Image = CType(resources.GetObject("cmdexit1.Image"), System.Drawing.Image)
        Me.cmdexit1.Location = New System.Drawing.Point(523, 0)
        Me.cmdexit1.Name = "cmdexit1"
        Me.cmdexit1.Size = New System.Drawing.Size(43, 28)
        Me.cmdexit1.TabIndex = 8
        Me.cmdexit1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel1.Location = New System.Drawing.Point(12, 11)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(605, 352)
        Me.Panel1.TabIndex = 2
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblproject, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 2)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.486166!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90.51383!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(602, 344)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clPartNo, Me.clDescription, Me.clMfrNo})
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.Menu
        Me.DataGridView1.Location = New System.Drawing.Point(3, 35)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(576, 306)
        Me.DataGridView1.TabIndex = 1
        '
        'clPartNo
        '
        Me.clPartNo.HeaderText = "Part No."
        Me.clPartNo.Name = "clPartNo"
        '
        'clDescription
        '
        Me.clDescription.HeaderText = "Description"
        Me.clDescription.Name = "clDescription"
        '
        'clMfrNo
        '
        Me.clMfrNo.HeaderText = "Manufacture Part No."
        Me.clMfrNo.Name = "clMfrNo"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(592, 10)
        Me.Label1.Margin = New System.Windows.Forms.Padding(10, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 0
        '
        'lblproject
        '
        Me.lblproject.AutoSize = True
        Me.lblproject.Location = New System.Drawing.Point(3, 6)
        Me.lblproject.Margin = New System.Windows.Forms.Padding(3, 6, 3, 0)
        Me.lblproject.Name = "lblproject"
        Me.lblproject.Size = New System.Drawing.Size(0, 13)
        Me.lblproject.TabIndex = 2
        '
        'frmproductsdevelopmentmanu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(619, 418)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmproductsdevelopmentmanu"
        Me.Text = "frmproductsdevelopmentmanu"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel2 As Panel
    Friend WithEvents cmdSave1 As Button
    Friend WithEvents cmdexit1 As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents lblproject As Label
    Friend WithEvents clPartNo As DataGridViewTextBoxColumn
    Friend WithEvents clDescription As DataGridViewTextBoxColumn
    Friend WithEvents clMfrNo As DataGridViewTextBoxColumn
End Class
