<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmpasschange
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmpasschange))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtcurpass = New System.Windows.Forms.TextBox()
        Me.txtnewpass = New System.Windows.Forms.TextBox()
        Me.txtnewpass2 = New System.Windows.Forms.TextBox()
        Me.cmdok = New System.Windows.Forms.Button()
        Me.cmdcancel = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 37.82383!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.08808!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.82902!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtcurpass, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtnewpass, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtnewpass2, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdok, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.cmdcancel, 2, 3)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(12, 12)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(386, 141)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 11, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter Current Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 46)
        Me.Label2.Margin = New System.Windows.Forms.Padding(3, 11, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Enter New Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 81)
        Me.Label3.Margin = New System.Windows.Forms.Padding(3, 11, 3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Enter New Password Again"
        '
        'txtcurpass
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtcurpass, 2)
        Me.txtcurpass.Location = New System.Drawing.Point(149, 3)
        Me.txtcurpass.Multiline = True
        Me.txtcurpass.Name = "txtcurpass"
        Me.txtcurpass.Size = New System.Drawing.Size(234, 29)
        Me.txtcurpass.TabIndex = 3
        '
        'txtnewpass
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtnewpass, 2)
        Me.txtnewpass.Location = New System.Drawing.Point(149, 38)
        Me.txtnewpass.Multiline = True
        Me.txtnewpass.Name = "txtnewpass"
        Me.txtnewpass.Size = New System.Drawing.Size(234, 29)
        Me.txtnewpass.TabIndex = 4
        '
        'txtnewpass2
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtnewpass2, 2)
        Me.txtnewpass2.Location = New System.Drawing.Point(149, 73)
        Me.txtnewpass2.Multiline = True
        Me.txtnewpass2.Name = "txtnewpass2"
        Me.txtnewpass2.Size = New System.Drawing.Size(234, 29)
        Me.txtnewpass2.TabIndex = 5
        '
        'cmdok
        '
        Me.cmdok.Location = New System.Drawing.Point(149, 108)
        Me.cmdok.Name = "cmdok"
        Me.cmdok.Size = New System.Drawing.Size(114, 30)
        Me.cmdok.TabIndex = 6
        Me.cmdok.Text = "OK"
        Me.cmdok.UseVisualStyleBackColor = True
        '
        'cmdcancel
        '
        Me.cmdcancel.Location = New System.Drawing.Point(269, 108)
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.Size = New System.Drawing.Size(114, 30)
        Me.cmdcancel.TabIndex = 7
        Me.cmdcancel.Text = "Cancel"
        Me.cmdcancel.UseVisualStyleBackColor = True
        '
        'frmpasschange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(415, 171)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmpasschange"
        Me.Text = "frmpasschange"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtcurpass As TextBox
    Friend WithEvents txtnewpass As TextBox
    Friend WithEvents txtnewpass2 As TextBox
    Friend WithEvents cmdok As Button
    Friend WithEvents cmdcancel As Button
End Class
