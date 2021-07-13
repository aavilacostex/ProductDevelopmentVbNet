Public Class btntextox
    Inherits TextBox

    Private _button As Button

    Public Sub btntextox(value As String)
        _button = New Button()
        'AddHandler myBtn.Click, AddressOf Me.myBtn_Click
    End Sub

    Private btn As Button
    Public Property myBtn() As Button
        Get
            Return btn
        End Get
        Set(ByVal value As Button)
            btn = value
        End Set
    End Property

    Private Sub myBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Handle your Button clicks here
    End Sub

End Class
