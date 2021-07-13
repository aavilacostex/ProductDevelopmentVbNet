<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Loading
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Loading))
        Me.pcloader = New System.Windows.Forms.PictureBox()
        CType(Me.pcloader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pcloader
        '
        Me.pcloader.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pcloader.Image = CType(resources.GetObject("pcloader.Image"), System.Drawing.Image)
        Me.pcloader.InitialImage = Nothing
        Me.pcloader.Location = New System.Drawing.Point(23, 10)
        Me.pcloader.Name = "pcloader"
        Me.pcloader.Size = New System.Drawing.Size(85, 66)
        Me.pcloader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pcloader.TabIndex = 4
        Me.pcloader.TabStop = False
        '
        'Loading
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(129, 88)
        Me.ControlBox = False
        Me.Controls.Add(Me.pcloader)
        Me.Name = "Loading"
        Me.Text = "Loading . . . "
        CType(Me.pcloader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pcloader As PictureBox
End Class
