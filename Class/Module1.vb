Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Public Module TextBoxExtensions

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function SendMessage(ByVal hWnd As HandleRef,
                                        ByVal msg As UInteger,
                                        ByVal wParam As IntPtr,
                                        ByVal lParam As String) As IntPtr
    End Function

    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub SetWatermark(ByVal ctl As Control, ByVal text As String)
        Const EM_SETCUEBANNER As Int32 = &H1501
        Const CB_SETCUEBANNER As Int32 = &H1703

        Dim retainOnFocus As IntPtr = New IntPtr(1)
        Dim msg As UInteger = EM_SETCUEBANNER

        If TypeOf ctl Is ComboBox Then
            msg = CB_SETCUEBANNER
        End If

        SendMessage(New HandleRef(ctl, ctl.Handle), msg, retainOnFocus, text)
    End Sub

    <DebuggerStepThrough()>
    <Runtime.CompilerServices.Extension()>
    Public Sub SetBtnTexbox(ByVal ctl As Control, ByVal ctlImg As ImageList, Optional ByVal text As String = Nothing)
        Dim btn As System.Windows.Forms.Button = New System.Windows.Forms.Button()
        btn.Size = New Size(25, ctl.ClientSize.Height + 2)
        btn.Location = New Point(ctl.ClientSize.Width - btn.Width - 1, -1)
        btn.FlatStyle = FlatStyle.Flat
        btn.Cursor = Cursors.Default
        'btn.Image = System.Windows.Forms.image Image. FromFile("C:\ansoft\Soljica\texture\tone.png")
        btn.Image = ctlImg.Images(0)
        btn.FlatAppearance.BorderSize = 0
        ctl.Controls.Add(btn)

        SendMessage(New HandleRef(ctl, ctl.Handle), &HD3, CType(2, IntPtr), CType((btn.Width << 16), IntPtr))

        SetWatermark(ctl, text)
    End Sub

End Module