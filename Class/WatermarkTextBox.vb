Imports System.ComponentModel
Imports System.Runtime.InteropServices


Public Class WatermarkTextBox
    Inherits TextBox

    <Localizable(True)>
    Public Property Cue() As String
        Get
            Return mCue
        End Get
        Set
            mCue = Value
            updateCue()
        End Set
    End Property

    Private Sub updateCue()
        If Me.IsHandleCreated AndAlso mCue IsNot Nothing Then
            SendMessage(Me.Handle, &H1501, (New IntPtr(1)), mCue)     'this line get the error msg
        End If
    End Sub
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        updateCue()
    End Sub
    Private mCue As String

    ' PInvoke
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

End Class
