Public Class test
    Private Sub test_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        'SplitContainer1.Text = "splitContainer2"
        SplitContainer1.BorderStyle = BorderStyle.FixedSingle
        'SplitContainer1.Panel1Collapsed = True
        'SplitContainer1.Panel2Collapsed = False
        SplitContainer1.Panel1.AutoSize = True
        SplitContainer1.Panel2.AutoSize = True
        'SplitContainer1.Dock = DockStyle.Fill
    End Sub

End Class