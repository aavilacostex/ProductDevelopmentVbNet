Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class ButtonTextBox
    Inherits TextBox

    Private ReadOnly _button As Button
    Private Shared ReadOnly SEARCH_BUTTON_WIDTH As Integer = 25

    Public Sub New()
        _button = New Button()
        _button.Cursor = Cursors.Default
        _button.Size = New Size(SEARCH_BUTTON_WIDTH, Me.ClientSize.Height + 2)
        _button.Dock = DockStyle.Right
        _button.Cursor = Cursors.Default
        '_button.Image = Properties.Resources.Find_5650
        _button.FlatStyle = FlatStyle.Flat
        _button.ForeColor = Color.White
        _button.FlatAppearance.BorderSize = 0
        Me.Controls.Add(_button)
        'SendMessage(Me.Handle, &H1501, (New IntPtr(2)), (New IntPtr(_button.Width << 16)))
        'SendMessage(Me.Handle, &H1501, (New IntPtr(1)), mCue)(New IntPtr(_button.Width << 16))
        ' Me.AcceptButton = _button

    End Sub

    Private btnButton As Button

    <Localizable(True)>
    Public Property customBtn() As Button
        Get
            Return btnButton
        End Get
        Set(ByVal value As Button)
            btnButton = value
            updateBtn()
        End Set
    End Property

    Private Sub updateBtn()
        If Me.IsHandleCreated AndAlso btnButton IsNot Nothing Then
            SendMessage(frmProductsDevelopment.Handle, &H1501, (New IntPtr(2)), (New IntPtr(_button.Width << 16)))
        End If
    End Sub

    Protected Overridable Sub OnResize(e As EventArgs)
        Me.OnResize(e)
        _button.Size = New Size(_button.Width, Me.ClientSize.Height + 2)
        _button.Location = New Point(Me.ClientSize.Width - _button.Width, -1)
        updateBtn()
        'SendMessage(this.Handle, 0xd3, (IntPtr)2, (IntPtr)(_button.Width << 16));
    End Sub

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    'Private Static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp)



    '    {
    '    base.OnResize(e);
    '    _button.Size = New Size(_button.Width, this.ClientSize.Height + 2);
    '    _button.Location = New Point(this.ClientSize.Width - _button.Width, -1);
    '    // Send EM_SETMARGINS to prevent text from disappearing underneath the button
    '    SendMessage(this.Handle, 0xd3, (IntPtr)2, (IntPtr)(_button.Width << 16));
    '}

    'Public Sub() {
    '    _button = New Button {Cursor = Cursors.Default};
    '    _button.SizeChanged += (o, e) => OnResize(e);
    '    this.Controls.Add(_button); 
    '}



End Class
