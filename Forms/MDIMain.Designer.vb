<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MDIMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIMain))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.PurchasingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomerClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductsDevelopmentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductsDevelopmentToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Test1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Test2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Test3ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PurchasingToolStripMenuItem, Me.Test1ToolStripMenuItem, Me.Test2ToolStripMenuItem, Me.Test3ToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'PurchasingToolStripMenuItem
        '
        Me.PurchasingToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClaimsToolStripMenuItem, Me.ProductsDevelopmentToolStripMenuItem})
        Me.PurchasingToolStripMenuItem.Name = "PurchasingToolStripMenuItem"
        Me.PurchasingToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.PurchasingToolStripMenuItem.Text = "Purchasing"
        '
        'ClaimsToolStripMenuItem
        '
        Me.ClaimsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SupplierClaimsToolStripMenuItem, Me.CustomerClaimsToolStripMenuItem})
        Me.ClaimsToolStripMenuItem.Name = "ClaimsToolStripMenuItem"
        Me.ClaimsToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.ClaimsToolStripMenuItem.Text = "Claims"
        Me.ClaimsToolStripMenuItem.Visible = False
        '
        'SupplierClaimsToolStripMenuItem
        '
        Me.SupplierClaimsToolStripMenuItem.Name = "SupplierClaimsToolStripMenuItem"
        Me.SupplierClaimsToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.SupplierClaimsToolStripMenuItem.Text = "Supplier Claims"
        '
        'CustomerClaimsToolStripMenuItem
        '
        Me.CustomerClaimsToolStripMenuItem.Name = "CustomerClaimsToolStripMenuItem"
        Me.CustomerClaimsToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.CustomerClaimsToolStripMenuItem.Text = "Customer Claims"
        '
        'ProductsDevelopmentToolStripMenuItem
        '
        Me.ProductsDevelopmentToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductsDevelopmentToolStripMenuItem1})
        Me.ProductsDevelopmentToolStripMenuItem.Name = "ProductsDevelopmentToolStripMenuItem"
        Me.ProductsDevelopmentToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.ProductsDevelopmentToolStripMenuItem.Text = "Products Development"
        '
        'ProductsDevelopmentToolStripMenuItem1
        '
        Me.ProductsDevelopmentToolStripMenuItem1.Name = "ProductsDevelopmentToolStripMenuItem1"
        Me.ProductsDevelopmentToolStripMenuItem1.Size = New System.Drawing.Size(195, 22)
        Me.ProductsDevelopmentToolStripMenuItem1.Text = "Products Development"
        '
        'Test1ToolStripMenuItem
        '
        Me.Test1ToolStripMenuItem.Name = "Test1ToolStripMenuItem"
        Me.Test1ToolStripMenuItem.Size = New System.Drawing.Size(75, 20)
        Me.Test1ToolStripMenuItem.Text = "Load Excel"
        '
        'Test2ToolStripMenuItem
        '
        Me.Test2ToolStripMenuItem.Name = "Test2ToolStripMenuItem"
        Me.Test2ToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.Test2ToolStripMenuItem.Text = "Test"
        Me.Test2ToolStripMenuItem.Visible = False
        '
        'Test3ToolStripMenuItem
        '
        Me.Test3ToolStripMenuItem.Name = "Test3ToolStripMenuItem"
        Me.Test3ToolStripMenuItem.Size = New System.Drawing.Size(45, 20)
        Me.Test3ToolStripMenuItem.Text = "Test3"
        Me.Test3ToolStripMenuItem.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(280, 115)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(217, 173)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "40th-ctp-logo.png")
        '
        'MDIMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MDIMain"
        Me.Text = "CTP INFORMATION SYSTEM"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents PurchasingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SupplierClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CustomerClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductsDevelopmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductsDevelopmentToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents Test1ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Test2ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Test3ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents ImageList1 As ImageList
End Class
