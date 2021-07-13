Imports System.Runtime.InteropServices

Public Class MyComboBox
    Inherits ComboBox

    Public Sub New()
        DrawMode = DrawMode.OwnerDrawFixed
        DropDownStyle = ComboBoxStyle.DropDownList
        'Dim ctor As New MyComboBox
    End Sub

    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        MyBase.OnDrawItem(e)

        Dim textColor = e.ForeColor

        If (e.State And DrawItemState.Focus) <> (DrawItemState.Focus) Then
            e.DrawBackground()
        Else
            textColor = Color.Green

        End If

        e.DrawFocusRectangle()

        Dim Index = e.Index
        If (Index < 0 Or Index >= Items.Count) Then Return
        Dim item = CType(Items(Index), DataRowView)


        Dim Text As String = If(item.Row.ItemArray(2).ToString() Is Nothing, Nothing, item.Row.ItemArray(2).ToString())
        Using brush = New SolidBrush(textColor)
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
            e.Graphics.DrawString(Text, e.Font, brush, e.Bounds)
        End Using

        'e.DrawBackground()
        'TextRenderer.DrawText(e.Graphics, text, Font, e.Bounds, e.ForeColor, TextFormatFlags.TextBoxControl)
    End Sub
End Class
