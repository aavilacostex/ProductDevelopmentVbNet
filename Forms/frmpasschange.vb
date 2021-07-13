Public Class frmpasschange
    Dim gnr As Gn1

    Dim KeyAscii As Integer
    Dim userid As String = gnr.userid
    Dim passcomm As String = gnr.passcomm
    Dim check As String


    Private Sub frmpasschange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub cmdcancel_Click()
        passcomm = ""
        'Unload frmpasschange
    End Sub
    Private Sub cmdok_Click()
        If Len(Trim(txtcurpass.Text)) = 0 And check = " " Then
            MsgBox("Current Password cannot be empty", vbOKOnly + vbInformation, "CTP System")
            SendKeys.Send("{tab}") ' Set the focus to the next control.
            SendKeys.Send("{tab}") ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
            'Cancel = True
            'Exit Sub
        Else
            If Len(Trim(txtnewpass.Text)) = 0 Then
                MsgBox("New Password cannot be empty", vbOKOnly + vbInformation, "CTP System")
                SendKeys.Send("{tab}") ' Set the focus to the next control.
                SendKeys.Send("{tab}") ' Set the focus to the next control.
                SendKeys.Send("{tab}") ' Set the focus to the next control.
                KeyAscii = 0        ' Ignore this key.
                'Cancel = True
                'Exit Sub
            Else
                If Len(Trim(txtnewpass2.Text)) = 0 Then
                    MsgBox("New Password cannot be empty", vbOKOnly + vbInformation, "CTP System")
                    SendKeys.Send("{tab}") ' Set the focus to the next control.
                    SendKeys.Send("{tab}") ' Set the focus to the next control.
                    SendKeys.Send("{tab}") ' Set the focus to the next control.
                    SendKeys.Send("{tab}") ' Set the focus to the next control.
                    KeyAscii = 0        ' Ignore this key.
                    'Cancel = True
                    'Exit Sub
                Else
                    'Call gotosetpass(Trim(UCase(userid)), Trim(txtnewpass.Text))
                    passcomm = Trim(txtnewpass.Text)
                    'Unload frmpasschange
                End If
            End If
        End If
    End Sub

    Private Sub Form_Load()
        check = " "
    End Sub

    Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

    End Sub

    Private Sub txtcurpass_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys.Send("{tab}") ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
        End If
    End Sub
    Private Sub txtcurpass_GotFocus()
        txtcurpass.SelectionStart = 0
        txtcurpass.SelectionLength = Len(Trim(txtcurpass.Text))
    End Sub
    Private Sub txtcurpass_LostFocus()
        On Error GoTo errhandler
        Exit Sub
errhandler:
        'Call gotoerror("frmpasschange", "txtcurpass_lostfocus", Err.Number, Err.Description, Err.Source)
    End Sub
    'Private Sub txtcurpass_Validate(Cancel As Boolean)
    'On Error GoTo errhandler
    'If Len(Trim(txtcurpass.Text)) > 0 Then
    '    check = checkusr(Trim(UCase(userid)), Trim(UCase(txtcurpass.Text)))
    '    If check <> "0" And check <> "E" Then
    '        MsgBox "Current Password is wrong", vbOKOnly + vbInformation, "CTP System"
    '        Cancel = True
    '        Exit Sub
    '    End If
    'End If
    'Exit Sub
    'errhandler:
    'Call gotoerror("frmpasschange", "txtcurpass_validate", Err.Number, Err.description, Err.Source)
    'End Sub

    Private Sub txtnewpass_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys.Send("{tab}")  ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
        End If
    End Sub
    Private Sub txtnewpass_GotFocus()
        txtnewpass.SelectionStart = 0
        txtnewpass.SelectionLength = Len(Trim(txtnewpass.Text))
    End Sub
    Private Sub txtnewpass_LostFocus()
        On Error GoTo errhandler
        Exit Sub
errhandler:
        'Call gotoerror("frmpasschange", "txtnewpass_lostfocus", Err.Number, Err.Description, Err.Source)
    End Sub
    Private Sub txtnewpass_Validate(Cancel As Boolean)
        Dim intrespond
        On Error GoTo errhandler
        If Len(Trim(txtnewpass.Text)) > 0 Then
            If check = " " Then
                'check = checkusr(Trim(UCase(userid)), Trim(UCase(txtcurpass.Text)))
                check = 0
                If check <> "0" And check <> "E" Then
                    MsgBox("Current Password is wrong", vbOKOnly + vbInformation, "CTP System")
                    SendKeys.Send("{tab}") ' Set the focus to the next control.
                    KeyAscii = 0        ' Ignore this key.
                    Cancel = True
                    Exit Sub
                End If
            End If
            If Len(Trim(txtnewpass.Text)) < 6 Then
                intrespond = MsgBox("Minimum length of Password is 6", vbOKOnly + vbInformation, "CTP System")
                Cancel = True
                Exit Sub
            Else
                If Len(Trim(txtnewpass.Text)) > 8 Then
                    intrespond = MsgBox("Maximum length of Password is 8", vbOKOnly + vbInformation, "CTP System")
                    Cancel = True
                    Exit Sub
                Else
                    If InStr(1, Trim(txtnewpass.Text), " ") Or Not gnr.checkstring(Trim(txtnewpass.Text)) Then
                        MsgBox("New Password cannot have spaces or special characters", vbOKOnly + vbInformation, "CTP System")
                        Cancel = True
                        Exit Sub
                    Else
                        If Trim(txtnewpass.Text) = Trim(txtcurpass.Text) Then
                            MsgBox("Your New Password must be different than the Current one", vbOKOnly + vbInformation, "CTP System")
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        Exit Sub
errhandler:
        'Call gotoerror("frmpasschange", "txtnewpass_validate", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Sub txtnewpass2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            Call cmdok_Click()
            'SendKeys "{tab}"    ' Set the focus to the next control.
            'KeyAscii = 0        ' Ignore this key.
        End If
    End Sub
    Private Sub txtnewpass2_GotFocus()
        txtnewpass2.SelectionStart = 0
        txtnewpass2.SelectionLength = Len(Trim(txtnewpass2.Text))
    End Sub
    Private Sub txtnewpass2_LostFocus()
        On Error GoTo errhandler
        Exit Sub
errhandler:
        'Call gotoerror("frmpasschange", "txtnewpass2_lostfocus", Err.Number, Err.Description, Err.Source)
    End Sub
    Private Sub txtnewpass2_Validate(Cancel As Boolean)
        On Error GoTo errhandler
        If Len(Trim(txtnewpass2.Text)) > 0 Then
            If Trim(txtnewpass2.Text) <> Trim(txtnewpass.Text) Then
                MsgBox("You must enter the same New Password", vbOKOnly + vbInformation, "CTP System")
                Cancel = True
                Exit Sub
            End If
        End If
        Exit Sub
errhandler:
        'Call gotoerror("frmpasschange", "txtnewpass2_validate", Err.Number, Err.Description, Err.Source)
    End Sub



End Class