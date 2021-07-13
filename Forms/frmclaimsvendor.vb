Public Class frmclaimsvendor
    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub frmclaimsvendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SSTab1.ItemSize = (New Size(SSTab1.Width / SSTab1.TabCount, 0))
        SSTab1.Padding = New System.Drawing.Point(300, 10)
        SSTab1.Appearance = TabAppearance.FlatButtons
        'TabControl1.ItemSize = New Size(0, 1)
        SSTab1.SizeMode = TabSizeMode.Fixed

        DataGridView1.RowHeadersVisible = False
        DataGridView2.RowHeadersVisible = False

        cmdnew.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\doc.PNG")
        cmdnew.ImageAlign = ContentAlignment.MiddleRight
        cmdnew.TextAlign = ContentAlignment.MiddleLeft

        cmdSave.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\save.PNG")
        cmdSave.ImageAlign = ContentAlignment.MiddleRight
        cmdSave.TextAlign = ContentAlignment.MiddleLeft

        cmdExit.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\exit.PNG")
        cmdExit.ImageAlign = ContentAlignment.MiddleRight
        cmdExit.TextAlign = ContentAlignment.MiddleLeft
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class