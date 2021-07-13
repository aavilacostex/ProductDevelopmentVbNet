<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LoadingExcel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoadingExcel))
        Me.ExcelLoader = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.ExcelLoader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ExcelLoader
        '
        Me.ExcelLoader.Image = CType(resources.GetObject("ExcelLoader.Image"), System.Drawing.Image)
        Me.ExcelLoader.Location = New System.Drawing.Point(12, 37)
        Me.ExcelLoader.Name = "ExcelLoader"
        Me.ExcelLoader.Size = New System.Drawing.Size(108, 82)
        Me.ExcelLoader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.ExcelLoader.TabIndex = 0
        Me.ExcelLoader.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Processing ..."
        '
        'LoadingExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(131, 131)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ExcelLoader)
        Me.Name = "LoadingExcel"
        Me.Text = "Loading.."
        CType(Me.ExcelLoader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ExcelLoader As PictureBox
    Friend WithEvents Label1 As Label
End Class
