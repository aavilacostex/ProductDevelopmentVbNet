Imports System.Reflection

Public Class frmLogin

#Region "Variables"

    Public intrespond As Long
    Public LoginSucceeded As Boolean
    Public s As String
    Public sql As String

    Dim frm As frmLogin
    Dim gnr As Gn1 = New Gn1()

    Public Conn As New ADODB.Connection
    Public ConnSql As New ADODB.Connection
    Public codloginctp As Long
    Public Versionctp As String '= gnr.Versionctp
    Public rs As ADODB.Recordset '= gnr.rs
    Dim CurrentCTPVersion As Version = My.Application.Info.Version
    Dim userid As String '= gnr.userid
    Dim passcomm As String '= gnr.passcomm
    Dim colorbackcolor As Integer
    Dim BackColorValue As Color
    Dim initialwindow As String

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)

#End Region

    Private Sub Form_Load()
        Dim IpAddrs
        Dim find1 As Integer
        Dim find2 As Integer
        Dim find3 As Long
        Dim find4 As String
        Dim printpath As String
        Dim exMessage As String = " "

        'On Error GoTo errhandler
        Try
            Conn.ConnectionString = gnr.Conexion
            Conn.Open()

            Dim dsControlData = gnr.GetDataByPartMix1()
            If Not dsControlData Is Nothing Then
                If dsControlData.Tables(0).Rows.Count > 0 Then
                    For Each tt As DataRow In dsControlData.Tables(0).Rows
                        If tt.ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "SVR" Then
                            gnr.primaryservername = Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                                Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                        End If
                        If tt.ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "PIC" Then
                            gnr.pathpicture = Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                                Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                        End If
                        If tt.ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "GEN" Then
                            gnr.Path = Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                                Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                        End If
                        If tt.ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "EMA" Then
                            gnr.emailspath = Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                                Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                        End If
                        If tt.ItemArray(dsControlData.Tables(0).Columns("cnt03").Ordinal) = "REP" Then
                            gnr.JiraPathBaseValue = Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde1").Ordinal)) &
                                Trim(tt.ItemArray(dsControlData.Tables(0).Columns("cntde2").Ordinal))
                        End If
                    Next
                End If
            End If
            'Public Const pathpicture = "\\Dellserver\CTP_System\images\Employee ID Pictures\"
            'Public Const pathgeneral = "\\Dellserver\Inetpub_D\"
            'Public Const emailspath = "\\Dellserver\Inetpub_D\CTP_System\Emails"

            'IpAddrs = gnr.GetIpAddrTable
            'IpAddrs = gnr.LocalIPAddress
            'IpAddrs = gnr.GetARPTablr()
            IpAddrs = gnr.GetIPv4Address()
            gnr.ipaddresslocal = IpAddrs
            Versionctp = CurrentCTPVersion.Build & " - " & Strings.Right(IpAddrs, 5)
            'Versionctp = Version & " - " & Right(IpAddrs(1), 5)

            'revisar este codigo con alejandro
            'find1 = InStr(1, Trim(IpAddrs(1)), ".")
            'find2 = InStr(find1 + 1, Trim(IpAddrs(1)), ".")
            'find3 = InStr(find2 + 1, Trim(IpAddrs(1)), ".")
            'find4 = Mid(Trim(IpAddrs(1)), find2 + 1, find3 - find2 - 1)
            'If find4 = 12 Then
            '    printpath = "\\Dalsvr\CTP_System\Reports"
            'End If

            cmdok.FlatStyle = FlatStyle.Flat
            cmdcancel.FlatStyle = FlatStyle.Flat

            Dim maxValue = gnr.getmax("loginctp", "codlogin")
            If Not String.IsNullOrEmpty(maxValue) Then
                maxValue += 1
            Else
                maxValue = 1 'preguntar duda
            End If

            'codloginctp = gnr.getmax("loginctp", "codlogin")
            Dim userName = If(String.IsNullOrEmpty(Environment.UserName), "", Environment.UserName)
            If Not String.IsNullOrEmpty(userName) Then
                Dim rsInsLoginTcp = gnr.InsertIntoLoginTcp(maxValue, userName, Versionctp)
                If rsInsLoginTcp <> 0 Then
                    'error message
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#Region "Key Events"

    Private Sub txtUserName_GotFocus(sender As Object, e As EventArgs) Handles txtUserName.GotFocus
        txtUserName.SelectionStart = 0
        txtUserName.SelectionLength = Len(txtUserName.Text)
    End Sub

    Private Sub txtpassword_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            'Call cmdok_Click
        End If
    End Sub

    Private Sub txtusername_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys.Send("{tab}")
            'SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
        End If
    End Sub

    Private Sub cmdok_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress, cmdok.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            cmdok_Click()
        End If

        'If KeyAscii = 13 Then  ' The ENTER key.
        '    SendKeys.Send("{tab}")
        'SendKeys "{tab}"    ' Set the focus to the next control.
        '    KeyAscii = 0        ' Ignore this key.
        'End If
    End Sub

    Private Sub cmdcancel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress, cmdcancel.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            cmdcancel_Click()
        End If
    End Sub

    Private Sub txtpassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress, txtPassword.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            cmdok_Click()
            txtPassword.Text = txtPassword.Text.Replace(Environment.NewLine, "")
        End If
    End Sub

#End Region

#Region "Control Events"

    Private Sub cmdcancel_Click()
        Dim exMessage As String = Nothing
        Try
            'sql = "CALL CTPINV.RECLAIM"
            'Conn.Execute (sql)
            Dim rsDeletionLoginTcp = gnr.DeleteRecorFromLoginTcp(codloginctp)
            If rsDeletionLoginTcp < 0 Then
                Log.Error("Error deleting record from Login CTP Table")
                'error eliminacion mensaje
            End If

            'sql = "delete from loginctp where codlogin = " & codloginctp
            'Conn.Execute(sql)
            If Conn.State = 1 Then
                Conn.Close()
            End If
            If ConnSql.State = 1 Then
                ConnSql.Close()
            End If
            LoginSucceeded = False
            Me.Close()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Log.Error(exMessage)
            LoginSucceeded = False
            Me.Close()
        End Try

    End Sub

    Private Sub cmdok_Click()
        'Dim WinWnd As Long
        Dim check As String
        Dim totaldays As Integer
        Dim exMessage As String = " "
        Try
            'WinWnd = FindWindow(vbNullString, "CTPSystem " & Version)
            'If WinWnd <> 0 Then
            '    End
            '    Exit Sub
            'End If

            colorbackcolor = 0
            initialwindow = "Main Menu"

            'servername = Conn.DefaultDatabase
            Dim servername As String = ""
            check = gnr.checkusr(Trim(UCase(txtUserName.Text)), Trim(UCase(txtPassword.Text)))
            If check = "U" Then
                MsgBox("Username not valid.", vbInformation + vbOKOnly, "CTP System")
                'txtUserName.SetFocus
            Else
                If check = "N" Then
                    MsgBox("User not authorized.", vbOKOnly + vbInformation, "CTP System")
                    Exit Sub
                End If
                If check = "E" Then
                    userid = Trim(UCase(txtUserName.Text))
                    MsgBox("Your password has expired; please change it.", vbOKOnly + vbInformation, "CTP System")
                    'frmpasschange.Show 1
                    If passcomm = "" Then
                        MsgBox("Password expired.", vbOKOnly + vbInformation, "CTP System")
                        Exit Sub
                    Else
                        check = "0"
                    End If
                End If
                If check = "0" And Len(Trim(UCase(txtPassword.Text))) = 0 Then
                    userid = Trim(UCase(txtUserName.Text))
                    MsgBox("You need to set a new password; please change it.", vbOKOnly + vbInformation, "CTP System")
                    'frmpasschange.Show 1
                    If passcomm = "" Then
                        MsgBox("Password not set.", vbOKOnly + vbInformation, "CTP System")
                        Exit Sub
                    Else
                        check = "0"
                    End If
                End If
                If check = "0" Then
                    userid = Trim(UCase(txtUserName.Text))
                    'pass = Trim(UCase(txtpassword.Text)))

                    Dim dsUsrData = gnr.getUserDataByUsername(userid)
                    If Not dsUsrData Is Nothing Then
                        If dsUsrData.Tables(0).Rows.Count > 0 Then
                            initialwindow = "Main Menu"
                            LoginSucceeded = True
                            Me.Hide()
                            MDIMain.Show()
                            'MDIMain.toolbar1.Visible = True
                            If dsUsrData.Tables(0).Rows(0).ItemArray(dsUsrData.Tables(0).Columns("DECODE").Ordinal) = 14 Then 'Or userid = "JDMERCADO" Then
                                Dim dsMarktData = gnr.getMarketingDataByDate()
                                If Not dsMarktData Is Nothing Then
                                    If dsMarktData.Tables(0).Rows.Count > 0 Then
                                        Dim dbDate = CDate(dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACABD").Ordinal))
                                        Dim todayDate = CDate(Now.ToShortDateString())
                                        totaldays = (dbDate - todayDate).TotalDays
                                        Dim macana = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACANA").Ordinal)
                                        Dim macabd = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACABD").Ordinal)
                                        Dim macaco = dsMarktData.Tables(0).Rows(0).ItemArray(dsMarktData.Tables(0).Columns("MACACO").Ordinal)

                                        If totaldays <= dsUsrData.Tables(0).Rows(0).ItemArray(dsUsrData.Tables(0).Columns("MACADY").Ordinal) Then
                                            Dim rsMessage As DialogResult = MessageBox.Show(" " & Trim(macana) & " Date : " & Format(macabd, "mm/dd/yyyy"), "CTP System", MessageBoxButtons.OKCancel)
                                            If rsMessage = DialogResult.Cancel Then
                                                Dim qryResult = gnr.UpdateMarktCampaignData(macaco)
                                                If qryResult <> 0 Then
                                                    'error actualizacion mensaje
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            'sql = "delete from loginctp where codlogin = " & codloginctp
                            Dim rsDeletionLoginTcp = gnr.DeleteRecorFromLoginTcp(codloginctp)
                            If rsDeletionLoginTcp < 0 Then
                                Log.Error("Error deleting record from Login CTP Table")
                                'error eliminacion mensaje
                            End If

                            Versionctp = CurrentCTPVersion.Build & " - " & Strings.Right(gnr.ipaddresslocal, 5)

                            'codloginctp = gnr.getmax("loginctp", "codlogin")
                            'sql = "INSERT INTO LOGINCTP VALUES(" & codloginctp & ",'" & userid & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','" & Versionctp & "')"
                            'Conn.Execute(sql)

                            Dim maxValue = gnr.getmax("loginctp", "codlogin")
                            If Not String.IsNullOrEmpty(maxValue) Then
                                maxValue += 1
                            Else
                                maxValue = 1 'preguntar duda
                            End If
                            Dim rsInsLoginTcp = gnr.InsertIntoLoginTcp(maxValue, Trim(UCase(txtUserName.Text)), Versionctp)
                            If rsInsLoginTcp <> 0 Then
                                'error message
                            End If
                            Call amenu()

                            Dim allowedAdminUser = gnr.getFlagAllow(Trim(UCase(txtUserName.Text)))
                            LikeSession.flagAccessAllow = allowedAdminUser

                            If userid = "CARLOS" Or userid = "JDMERCADO" Or userid = "MVELEZ" Or userid = "KRODRIGUEZ" Or userid = "JDMIRA" Or userid = "HOLIVEROS" Or userid = "LARIAS" Or userid = "AAVILA" Then
                                'ConnSql.ConnectionString = gnr.SQLCon
                                'ConnSql.Open()

                                'gnr.ConnSqlNOVA.ConnectionString = gnr.NOVASQLCon
                                'gnr.ConnSqlNOVA.Open()
                            End If
                        End If
                    End If
                Else
                    Dim rsMessageInvalid As DialogResult = MessageBox.Show("Invalid Password, try again!", "CTP System", MessageBoxButtons.OK)
                    'MsgBox("Invalid Password, try again!", vbOKOnly + vbInformation, "CTP System")
                    'txtPassword.SetFocus
                    'SendKeys.Send("{Home}+{End}")
                End If
            End If
            Exit Sub
        Catch ex As Exception
            If gnr.ConnSqlNOVA.State = 0 Then
                Dim rsMessageNova As DialogResult = MessageBox.Show("Connection to NOVATIME failed!", "CTP System", MessageBoxButtons.OK)
                'MsgBox("Connection to NOVATIME failed!", vbOKOnly + vbInformation, "CTP System")
            End If
            Exit Sub
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

    Private Sub atoolbar(dkey)
        'On Error GoTo errhandler
        'j = 1
        'For j = 1 To MDIMain.toolbar1.Buttons.Count
        '    If MDIMain.toolbar1.Buttons.Item(j).Key = dkey Then
        '        MDIMain.toolbar1.Buttons.Item(j).Enabled = False
        '        j = MDIMain.toolbar1.Buttons.Count + 1
        '    End If
        'Next j
        'Exit Sub
        'errhandler:
        'Call gotoerror("frmlogin", "atoolbar", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Load()
    End Sub

    Private Sub amenu()
        Dim errormark As Long
        Dim exMessage As String = " "
        Try
            errormark = 0
            MDIMain.Show()
            'MDIMain.toolbar1.Visible = True
            errormark = 1
            If colorbackcolor <> 0 Then
                errormark = 2
                MDIMain.BackColor = ColorTranslator.FromWin32(colorbackcolor)
            End If
            errormark = 3
            Dim dsGetUserMenuByUserId = gnr.GetMenuByUser(userid)
            errormark = 4
            Dim myMenu As MenuStrip
            If dsGetUserMenuByUserId IsNot Nothing Then
                If dsGetUserMenuByUserId.Tables(0).Rows.Count > 0 And MDIMain.Controls.Count > 0 Then
                    errormark = 5

                    'version produccion
                    For Each tt As DataRow In dsGetUserMenuByUserId.Tables(0).Rows
                        For Each ttt As Control In MDIMain.Controls
                            If Trim(ttt.Name = tt.Item("DMDIMAIN")) Then
                                errormark = 6
                                If tt.Item("AMENU") = 1 Then
                                    errormark = 7
                                    ttt.Enabled = True
                                Else
                                    ttt.Enabled = False
                                End If
                                Exit For
                            End If
                        Next
                    Next




                    'version simple de prueba
                    'myMenu = MDIMain.MenuStrip1
                    'For Each ttt As DataRow In dsGetUserMenuByUserId.Tables(0).Rows
                    '    For Each tt As ToolStripMenuItem In myMenu.Items
                    '        If tt.Name = ttt.ItemArray(dsGetUserMenuByUserId.Tables(0).Columns("dmdimain").Ordinal) Then
                    '            errormark = 6
                    '            If ttt.ItemArray(dsGetUserMenuByUserId.Tables(0).Columns("amenu").Ordinal) = 1 Then
                    '                errormark = 7
                    '                tt.Enabled = True
                    '            Else
                    '                tt.Enabled = False
                    '            End If
                    '        End If
                    '    Next
                    'Next
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub cmdok_Click(sender As Object, e As EventArgs) Handles cmdok.Click
        cmdok_Click()
    End Sub

    Private Sub cmdcancel_Click(sender As Object, e As EventArgs) Handles cmdcancel.Click
        cmdcancel_Click()
    End Sub

End Class