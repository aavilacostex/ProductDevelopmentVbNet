<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmproductsdevelopmentcomments
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmproductsdevelopmentcomments))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCode = New System.Windows.Forms.TextBox()
        Me.DTPicker1 = New System.Windows.Forms.DateTimePicker()
        Me.txtpartno = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DTPicker2 = New System.Windows.Forms.DateTimePicker()
        Me.txtsubject = New System.Windows.Forms.TextBox()
        Me.dgvAddComments = New System.Windows.Forms.DataGridView()
        Me.clComments = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblNotVisible = New System.Windows.Forms.Label()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdnew = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.dgvAddComments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.0202!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 2.850877!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.17544!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 168.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCode, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DTPicker1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtpartno, 4, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.DTPicker2, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtsubject, 1, 2)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(27, 29)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(609, 100)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 10)
        Me.Label1.Margin = New System.Windows.Forms.Padding(20, 10, 3, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Project No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 43)
        Me.Label2.Margin = New System.Windows.Forms.Padding(20, 10, 3, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Transaction Date"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 72)
        Me.Label3.Margin = New System.Windows.Forms.Padding(20, 6, 3, 3)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Subject"
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(122, 7)
        Me.txtCode.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(214, 20)
        Me.txtCode.TabIndex = 3
        '
        'DTPicker1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.DTPicker1, 2)
        Me.DTPicker1.Location = New System.Drawing.Point(122, 40)
        Me.DTPicker1.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.DTPicker1.Name = "DTPicker1"
        Me.DTPicker1.Size = New System.Drawing.Size(219, 20)
        Me.DTPicker1.TabIndex = 7
        '
        'txtpartno
        '
        Me.txtpartno.Location = New System.Drawing.Point(442, 7)
        Me.txtpartno.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.txtpartno.Name = "txtpartno"
        Me.txtpartno.Size = New System.Drawing.Size(144, 20)
        Me.txtpartno.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(371, 10)
        Me.Label4.Margin = New System.Windows.Forms.Padding(20, 10, 3, 3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Part No."
        '
        'DTPicker2
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.DTPicker2, 2)
        Me.DTPicker2.Location = New System.Drawing.Point(371, 40)
        Me.DTPicker2.Margin = New System.Windows.Forms.Padding(20, 7, 3, 3)
        Me.DTPicker2.Name = "DTPicker2"
        Me.DTPicker2.Size = New System.Drawing.Size(219, 20)
        Me.DTPicker2.TabIndex = 8
        '
        'txtsubject
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtsubject, 4)
        Me.txtsubject.Location = New System.Drawing.Point(122, 73)
        Me.txtsubject.Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)
        Me.txtsubject.Name = "txtsubject"
        Me.txtsubject.Size = New System.Drawing.Size(477, 20)
        Me.txtsubject.TabIndex = 9
        '
        'dgvAddComments
        '
        Me.dgvAddComments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAddComments.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clComments})
        Me.dgvAddComments.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvAddComments.Location = New System.Drawing.Point(27, 145)
        Me.dgvAddComments.Name = "dgvAddComments"
        Me.dgvAddComments.RowHeadersVisible = False
        Me.dgvAddComments.Size = New System.Drawing.Size(609, 233)
        Me.dgvAddComments.TabIndex = 1
        '
        'clComments
        '
        Me.clComments.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clComments.HeaderText = "Comments"
        Me.clComments.MaxInputLength = 100
        Me.clComments.Name = "clComments"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.GrayText
        Me.Panel1.Controls.Add(Me.lblNotVisible)
        Me.Panel1.Controls.Add(Me.cmdExit)
        Me.Panel1.Controls.Add(Me.cmdSave)
        Me.Panel1.Controls.Add(Me.cmdnew)
        Me.Panel1.Location = New System.Drawing.Point(27, 397)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(609, 34)
        Me.Panel1.TabIndex = 2
        '
        'lblNotVisible
        '
        Me.lblNotVisible.AutoSize = True
        Me.lblNotVisible.ForeColor = System.Drawing.SystemColors.GrayText
        Me.lblNotVisible.Location = New System.Drawing.Point(337, 17)
        Me.lblNotVisible.Name = "lblNotVisible"
        Me.lblNotVisible.Size = New System.Drawing.Size(13, 13)
        Me.lblNotVisible.TabIndex = 3
        Me.lblNotVisible.Text = "0"
        Me.lblNotVisible.Visible = False
        '
        'cmdExit
        '
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.Location = New System.Drawing.Point(560, 3)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(43, 28)
        Me.cmdExit.TabIndex = 2
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Image)
        Me.cmdSave.Location = New System.Drawing.Point(511, 3)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(43, 28)
        Me.cmdSave.TabIndex = 1
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdnew
        '
        Me.cmdnew.Image = CType(resources.GetObject("cmdnew.Image"), System.Drawing.Image)
        Me.cmdnew.Location = New System.Drawing.Point(462, 3)
        Me.cmdnew.Name = "cmdnew"
        Me.cmdnew.Size = New System.Drawing.Size(43, 28)
        Me.cmdnew.TabIndex = 0
        Me.cmdnew.UseVisualStyleBackColor = True
        '
        'frmproductsdevelopmentcomments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(666, 450)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.dgvAddComments)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "frmproductsdevelopmentcomments"
        Me.Text = "frmproductsdevelopmentcomments"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.dgvAddComments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtCode As TextBox
    Friend WithEvents DTPicker1 As DateTimePicker
    Friend WithEvents txtpartno As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents DTPicker2 As DateTimePicker
    Friend WithEvents txtsubject As TextBox
    Friend WithEvents dgvAddComments As DataGridView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents cmdExit As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents cmdnew As Button
    Friend WithEvents clComments As DataGridViewTextBoxColumn
    Friend WithEvents lblNotVisible As Label
End Class
