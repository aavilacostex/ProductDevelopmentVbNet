<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChangeVendor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChangeVendor))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.lblVendorName = New System.Windows.Forms.Label()
        Me.txtCode = New System.Windows.Forms.TextBox()
        Me.cmbvendor = New System.Windows.Forms.ComboBox()
        Me.cmdsearch = New System.Windows.Forms.Button()
        Me.cmdsearch1 = New System.Windows.Forms.Button()
        Me.pnBottom = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdchange = New System.Windows.Forms.Button()
        Me.txtsearch1 = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.pnBottom.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 140.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 81.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 156.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblCode, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblVendorName, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCode, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbvendor, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdsearch, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdsearch1, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.pnBottom, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtsearch1, 1, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(24, 22)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 92.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 10.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(407, 137)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'lblCode
        '
        Me.lblCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCode.Location = New System.Drawing.Point(3, 8)
        Me.lblCode.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.lblCode.Name = "lblCode"
        Me.lblCode.Size = New System.Drawing.Size(118, 13)
        Me.lblCode.TabIndex = 0
        Me.lblCode.Text = "Search by Vendor #"
        '
        'lblVendorName
        '
        Me.lblVendorName.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorName.Location = New System.Drawing.Point(3, 36)
        Me.lblVendorName.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.lblVendorName.Name = "lblVendorName"
        Me.lblVendorName.Size = New System.Drawing.Size(134, 23)
        Me.lblVendorName.TabIndex = 1
        Me.lblVendorName.Text = "Search by Vendor  Name"
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(143, 3)
        Me.txtCode.Multiline = True
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(96, 18)
        Me.txtCode.TabIndex = 2
        '
        'cmbvendor
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.cmbvendor, 3)
        Me.cmbvendor.FormattingEnabled = True
        Me.cmbvendor.Location = New System.Drawing.Point(143, 63)
        Me.cmbvendor.Name = "cmbvendor"
        Me.cmbvendor.Size = New System.Drawing.Size(258, 21)
        Me.cmbvendor.TabIndex = 4
        '
        'cmdsearch
        '
        Me.cmdsearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsearch.Location = New System.Drawing.Point(245, 3)
        Me.cmdsearch.Name = "cmdsearch"
        Me.cmdsearch.Size = New System.Drawing.Size(75, 23)
        Me.cmdsearch.TabIndex = 5
        Me.cmdsearch.Text = "Search"
        Me.cmdsearch.UseVisualStyleBackColor = True
        '
        'cmdsearch1
        '
        Me.cmdsearch1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsearch1.Location = New System.Drawing.Point(326, 32)
        Me.cmdsearch1.Name = "cmdsearch1"
        Me.cmdsearch1.Size = New System.Drawing.Size(75, 23)
        Me.cmdsearch1.TabIndex = 6
        Me.cmdsearch1.Text = "Search"
        Me.cmdsearch1.UseVisualStyleBackColor = True
        '
        'pnBottom
        '
        Me.pnBottom.BackColor = System.Drawing.SystemColors.GrayText
        Me.TableLayoutPanel1.SetColumnSpan(Me.pnBottom, 4)
        Me.pnBottom.Controls.Add(Me.TableLayoutPanel2)
        Me.pnBottom.Location = New System.Drawing.Point(3, 91)
        Me.pnBottom.Name = "pnBottom"
        Me.pnBottom.Size = New System.Drawing.Size(398, 36)
        Me.pnBottom.TabIndex = 7
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 105.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 51.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.cmdExit, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.cmdchange, 0, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(242, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(156, 30)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'cmdExit
        '
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.Location = New System.Drawing.Point(108, 3)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(43, 24)
        Me.cmdExit.TabIndex = 7
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdchange
        '
        Me.cmdchange.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdchange.Location = New System.Drawing.Point(3, 3)
        Me.cmdchange.Name = "cmdchange"
        Me.cmdchange.Size = New System.Drawing.Size(99, 23)
        Me.cmdchange.TabIndex = 8
        Me.cmdchange.Text = "Change Vendor"
        Me.cmdchange.UseVisualStyleBackColor = True
        '
        'txtsearch1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtsearch1, 2)
        Me.txtsearch1.Location = New System.Drawing.Point(143, 32)
        Me.txtsearch1.Multiline = True
        Me.txtsearch1.Name = "txtsearch1"
        Me.txtsearch1.Size = New System.Drawing.Size(177, 20)
        Me.txtsearch1.TabIndex = 3
        '
        'frmChangeVendor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "frmChangeVendor"
        Me.Text = "frmChangeVendor"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.pnBottom.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents lblCode As Label
    Friend WithEvents lblVendorName As Label
    Friend WithEvents txtCode As TextBox
    Friend WithEvents txtsearch1 As TextBox
    Friend WithEvents cmbvendor As ComboBox
    Friend WithEvents cmdsearch As Button
    Friend WithEvents cmdsearch1 As Button
    Friend WithEvents pnBottom As Panel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents cmdExit As Button
    Friend WithEvents cmdchange As Button
End Class
