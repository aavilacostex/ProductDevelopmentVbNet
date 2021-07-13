Imports System.Threading

Public Class PleaseWait
    Private mSplash As Form
    Private mLocation As Point

    Sub MySub(location As Point)
        mLocation = location
        Dim t As Thread = New Thread(New ThreadStart(Sub()
                                                         workerThread()
                                                     End Sub))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
    End Sub

    Public Sub Dispose()
        mSplash.Invoke(New MethodInvoker(Sub()
                                             stopThread()
                                         End Sub))
    End Sub

    Private Sub stopThread()
        mSplash.Close()
    End Sub

    Private Sub workerThread()
        mSplash = New LoadingExcel()   ' Substitute this With your own
        mSplash.StartPosition = FormStartPosition.Manual
        mSplash.Location = mLocation
        mSplash.TopMost = True
        Application.Run(mSplash)
    End Sub

End Class
