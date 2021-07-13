Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports outlook = Microsoft.Office.Interop.Outlook
Imports System.Reflection
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Excel = Microsoft.Office.Interop.Excel
Imports ClosedXML.Excel
Imports System.Windows.Threading
Imports System.Windows.Threading.Dispatcher

Public Class frmProductsDevelopment
    Public flagdeve As Long '1 is new
    'Public filepicture As New clsReadWrite
    Public strwhere As String
    Public strToUnion As String
    Public strToUnionTab2 As String
    Public userid As String
    Public flagnewpart As Integer
    Public flagallow As Integer = 0
    Public puragent As Integer
    Dim sql As String
    Dim requireValidation As Integer = 0
    Dim partstoshow As String
    Dim toemails As String = ""
    Dim gnr As Gn1 = New Gn1()
    Dim vblog As VBLog = New VBLog()
    Dim wm As WatermarkTextBox = New WatermarkTextBox()
    Dim bt As ButtonTextBox = New ButtonTextBox()
    Dim dspCall As Dispatcher

    private strLogCadenaCabecera As string = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Excel03ConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"
    Private Excel07ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX={2}'"

    'Public Const PageSize = 10
    'Public Property TotalRecords() As Integer
    'Dim urlPathBase As String = "https://costex.atlassian.net/browse/"
    Dim pathpictureparts As String

    Dim bs As BindingSource = New BindingSource()
    Dim Tables = New BindingList(Of DataTable)()

    Dim bs1 As BindingSource = New BindingSource()
    Dim Tables1 = New BindingList(Of DataTable)()

    Public Event PositionChanged(sender As Object, e As EventArgs)

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")


    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' MsgBox("The application is terminating.")
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub frmProductsDevelopment_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not gnr.CheckForInternetConnection Then
            Log.Warn("There is an internet connection issue.Please try in a while!")
            MessageBox.Show("There is an internet connection issue.Please try in a while!", "CTP System", MessageBoxButtons.OK)
        Else
            LoadCombos(sender, e)
            frmProductsDevelopment_load()
            dspCall = CurrentDispatcher

            If gnr.FlagCloseMDIForm.Equals("0") Then

                '    If MDIMain.Visible Then
                '        MDIMain.Hide()
                '    End If

                BackgroundWorker2.RunWorkerAsync()
            End If
        End If

    End Sub

    Private Sub frmProductsDevelopment_load()
        Dim exMessage As String = " "
        Try

            'gnr.killBackgroundProcess()

            If CInt(gnr.FlagProductionMethod).Equals(1) Then
                userid = UCase(LikeSession.retrieveUser)
                'userid = "CMONTILVA"
            Else

                If frmLogin.txtUserName.Text IsNot Nothing Then
                    userid = UCase(LikeSession.retrieveUser)
                Else
                    userid = UCase(gnr.AuthorizatedTestUser)
                End If
            End If

            'test purpose
            'userid = "LREDONDO"

            LikeSession.userid = userid

            If gnr.getFlagAllow(userid) = 1 Then
                flagallow = 1
            Else
                cmbPrpech.Visible = False
                'cmddelete.Visible = False
                cmbuser2.Visible = False
            End If

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - PD Start", "")

            'Dim btn As System.Windows.Forms.Button = New System.Windows.Forms.Button()
            'btn.Size = New Size(25, txtsearchcode.ClientSize.Height + 2)
            'btn.Location = New Point(txtsearchcode.ClientSize.Width - btn.Width - 1, -1)
            'btn.FlatStyle = FlatStyle.Flat
            'btn.Cursor = Cursors.Default
            ''btn.Image = System.Windows.Forms.image Image. FromFile("C:\ansoft\Soljica\texture\tone.png")
            'btn.FlatAppearance.BorderSize = 0
            'txtsearchcode.Controls.Add(btn)
            'SendMessage(txtsearchcode.Handle, &HD3, CType(2, IntPtr), CType((btn.Width << 16), IntPtr))

            ResizeTabs()
            FillDDlUser()
            SetValues()

            FillDDLStatus1()
            FillDDlPrPech()
            FillDDlPrPech1()

            ToolTip2.SetToolTip(LinkLabel6, "Info")

            'testMethod()
            'test purpose
            'gnr.sendEmail()
            'Dim dss = gnr.GetPOQotaData()
            'dropdownlist default fill section
            'Dim varvar = 1439
            'Dim dstest = gnr.DeleteDataBynrojectNo(varvar)

            'Dim toemailsww = prepareEmailsToSend(1)
            'Dim rsResult = gnr.sendEmail(toemailsww, txtpartno.Text)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()


            Log.Info("test message")

            'writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub frmProductsDevelopment_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Dim exMessage As String = Nothing
        Try
            ToolTip2.SetToolTip(LinkLabel6, "Info")
            gnr.killBackgroundProcess()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, "Exception: ", exMessage)
            writeComputerEventLog()
        End Try

    End Sub

    Private Sub frmProductsDevelopment_Closing(sender As Object, e As EventArgs) Handles MyBase.FormClosing

        Dim exMessage As String = Nothing
        Try
            Application.Exit()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, "Exception: ", exMessage)
            writeComputerEventLog()
        End Try

    End Sub

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Public Sub testMethod()

    End Sub

#Region "Threads"

#Region "Thread 1"

    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
        Handles BackgroundWorker1.RunWorkerCompleted
        Loading.Close()
    End Sub

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) _
        Handles BackgroundWorker1.DoWork
        ExecuteCombos(sender, e)
    End Sub

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) _
        Handles BackgroundWorker1.ProgressChanged
        'txtMfrNoSearch.Text = e.ProgressPercentage.ToString()
    End Sub

#End Region

#Region "Thread 2"

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) _
        Handles BackgroundWorker2.RunWorkerCompleted
        If e.Cancelled Then
            'Label1.Text = "cancelled"
        ElseIf e.Error IsNot Nothing Then
            'Label1.Text = e.Error.Message
        Else
            'Label1.Text = "Sum = " & e.Result.ToString()
        End If
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        execute_delegate_MDIClose()
    End Sub

#End Region

#End Region

#Region "Combobox load Region"

    Private Sub FillDDlUser()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next


            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            cmbuser1.DataSource = dsUser.Tables(0)
            cmbuser1.DisplayMember = "FullValue"
            cmbuser1.ValueMember = "USUSER"

            'cmbuser1.SelectedIndex = cmbuser.FindString(Trim(UCase(userid)))

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            'Log.Error(exMessage)
        End Try
    End Sub

    Private Sub FillDDlUser1()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next


            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            cmbuser.DataSource = dsUser.Tables(0)
            cmbuser.DisplayMember = "FullValue"
            cmbuser.ValueMember = "USUSER"


        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            'Log.Error(exMessage)
        End Try
    End Sub

    Private Sub FillDDlUser2()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            Dim newRow1 As DataRow = dsUser.Tables(0).NewRow
            newRow1("USUSER") = ""
            newRow1("USNAME") = ""
            newRow1("FullValue") = ""
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow1, 0)

            cmbuser2.DataSource = dsUser.Tables(0)
            cmbuser2.DisplayMember = "FullValue"
            cmbuser2.ValueMember = "USUSER"
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            'Log.Error(exMessage)
        End Try
    End Sub

    Private Sub FillDDlPrPech1()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next


            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            Dim newRow1 As DataRow = dsUser.Tables(0).NewRow
            newRow1("USUSER") = ""
            newRow1("USNAME") = ""
            newRow1("FullValue") = ""
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow1, 0)

            cmbPrpech.DataSource = dsUser.Tables(0)
            cmbPrpech.DisplayMember = "FullValue"
            cmbPrpech.ValueMember = "USUSER"
            cmbPrpech.SelectedIndex = -1

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            'Log.Error(exMessage)
        End Try
    End Sub

    Private Sub FillDDlPrPech()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsUser = gnr.FillDDLUser()

            dsUser.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsUser.Tables(0).Rows.Count - 1
                If dsUser.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsUser.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsUser.Tables(0).Rows(i).Item(1).ToString()
                    CleanUser = Trim(dsUser.Tables(0).Rows(i).Item(0).ToString())
                    dsUser.Tables(0).Rows(i).Item(2) = fllValueName
                    dsUser.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next


            Dim newRow As DataRow = dsUser.Tables(0).NewRow
            newRow("USUSER") = "N/A"
            newRow("USNAME") = "NO NAME"
            newRow("FullValue") = "N/A -- NO NAME"
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow, 0)

            Dim newRow1 As DataRow = dsUser.Tables(0).NewRow
            newRow1("USUSER") = ""
            newRow1("USNAME") = ""
            newRow1("FullValue") = ""
            'dsUser.Tables(0).Rows.Add(newRow)
            dsUser.Tables(0).Rows.InsertAt(newRow1, 0)

            'MyComboBox1.DataSource = dsUser.Tables(0)
            'MyComboBox1.DisplayMember = "FullValue"
            'MyComboBox1.ValueMember = "USUSER"
            'MyComboBox1.SelectedIndex = -1

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub FillDDLStatus()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsStatuses = gnr.GetAllStatuses()

            dsStatuses.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsStatuses.Tables(0).Rows.Count - 1
                If dsStatuses.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsStatuses.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsStatuses.Tables(0).Rows(i).Item(1).ToString()
                    'CleanUser = Trim(dsStatuses.Tables(0).Rows(i).Item(0).ToString())
                    dsStatuses.Tables(0).Rows(i).Item(2) = fllValueName
                    'dsStatuses.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            cmbstatus.DataSource = dsStatuses.Tables(0)
            cmbstatus.DisplayMember = "FullValue"
            cmbstatus.ValueMember = "CNT03"

            'cmbstatus1.DataSource = dsStatuses.Tables(0)
            'cmbstatus1.DisplayMember = "FullValue"
            'cmbstatus1.ValueMember = "CNT03"

            'cmbstatus1.SelectedIndex = -1
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub FillDDLStatus1()
        Dim exMessage As String = " "
        Dim CleanUser As String
        Try
            Dim dsStatuses = gnr.GetAllStatuses()

            dsStatuses.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsStatuses.Tables(0).Rows.Count - 1
                If dsStatuses.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsStatuses.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsStatuses.Tables(0).Rows(i).Item(1).ToString()
                    'CleanUser = Trim(dsStatuses.Tables(0).Rows(i).Item(0).ToString())
                    dsStatuses.Tables(0).Rows(i).Item(2) = fllValueName
                    'dsStatuses.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            Dim newRow As DataRow = dsStatuses.Tables(0).NewRow
            newRow("CNT03") = ""
            newRow("CNTDE1") = ""
            newRow("FullValue") = ""
            'dsUser.Tables(0).Rows.Add(newRow)
            dsStatuses.Tables(0).Rows.InsertAt(newRow, 0)

            cmbstatus1.DataSource = dsStatuses.Tables(0)
            cmbstatus1.DisplayMember = "FullValue"
            cmbstatus1.ValueMember = "CNT03"

            'cmbstatus1.DataSource = dsStatuses.Tables(0)
            'cmbstatus1.DisplayMember = "FullValue"
            'cmbstatus1.ValueMember = "CNT03"

            'cmbstatus1.SelectedIndex = -1
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub FillDDlMinorCode()
        Dim exMessage As String = " "
        Try
            Dim dsMinCodes = gnr.FillDDlMinorCode()

            dsMinCodes.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsMinCodes.Tables(0).Rows.Count - 1
                If dsMinCodes.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsMinCodes.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsMinCodes.Tables(0).Rows(i).Item(1).ToString()
                    'dsMinCodes = Trim(dsMinCodes.Tables(0).Rows(i).Item(0).ToString())
                    dsMinCodes.Tables(0).Rows(i).Item(2) = fllValueName
                    'dsMinCodes.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            cmbminorcode.DataSource = dsMinCodes.Tables(0)
            cmbminorcode.DisplayMember = "FullValue"
            cmbminorcode.ValueMember = "CNT03"

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub FillDDlMajorCode()
        Dim exMessage As String = " "
        Try
            Dim dsMMajCodes = gnr.FillDDlMajorCode()

            dsMMajCodes.Tables(0).Columns.Add("FullValue", GetType(String))

            For i As Integer = 0 To dsMMajCodes.Tables(0).Rows.Count - 1
                If dsMMajCodes.Tables(0).Rows(i).Table.Columns("FullValue").ToString = "FullValue" Then
                    Dim fllValueName = dsMMajCodes.Tables(0).Rows(i).Item(0).ToString() + " -- " + dsMMajCodes.Tables(0).Rows(i).Item(1).ToString()
                    'dsMinCodes = Trim(dsMinCodes.Tables(0).Rows(i).Item(0).ToString())
                    dsMMajCodes.Tables(0).Rows(i).Item(2) = fllValueName
                    'dsMinCodes.Tables(0).Rows(i).Item(0) = CleanUser
                    'do something
                End If
            Next

            cmbmajorcode.DataSource = dsMMajCodes.Tables(0)
            cmbmajorcode.DisplayMember = "FullValue"
            cmbmajorcode.ValueMember = "CNT03"

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub SSTab1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
    Handles SSTab1.SelectedIndexChanged
        SSTab1_Selected(sender, Nothing)
    End Sub

    Private Sub SSTab1_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) _
    Handles SSTab1.Selected
        Dim exMessage As String = Nothing
        Try
            If SSTab1.SelectedIndex = 0 Then
                Panel4.Enabled = True
                cmdSave1.Enabled = False
                cmdcvendor.Enabled = False
                cmdchange.Enabled = False
                cmdmpartno.Enabled = False
                cmdunitcost.Enabled = False
                cmdexit1.Enabled = True
                cmdnew1.Enabled = True

                'If Not String.IsNullOrEmpty(txtCode.Text) Then
                '    checkPendingReferences(txtCode.Text)
                'Else
                '    If Not String.IsNullOrEmpty(txtsearchcode.Text) Then
                '        checkPendingReferences(txtsearchcode.Text)
                '    End If
                'End If


            ElseIf SSTab1.SelectedIndex = 1 Then
                Panel4.Enabled = True
                If flagdeve = 1 Then
                    cmdnew2.Enabled = False
                Else
                    cmdnew2.Enabled = True
                End If

                showTab2FilterPanel(dgvProjectDetails)

            ElseIf SSTab1.SelectedIndex = 2 Then
                Dim rsValue As Integer = -1
                Panel4.Enabled = True
                rsValue = mandatoryFields("new", SSTab1.SelectedIndex)
                If rsValue = 0 Then
                    flagdeve = 0
                    flagnewpart = 1
                Else
                    Dim rsMessage As DialogResult = MessageBox.Show("All the fields in the Project Tab must be filled before add parts!", "CTP System", MessageBoxButtons.OK)
                    If rsMessage = DialogResult.OK Then
                        SSTab1.SelectedIndex = 1
                    End If
                End If
            End If
            TableLayoutPanel15.Enabled = True
            cmdchange.Enabled = True
            cmdunitcost.Enabled = True
            cmdmpartno.Enabled = True
            cmdcvendor.Enabled = True
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    'Private Sub pScrollForm()
    '    Dim ctl As Control


    '    If SSTab1.SelectedIndex = 1 Then
    '        Dim myScrollBar As VScrollBar = Me.Controls.Find("vScrollBar1", True).FirstOrDefault()
    '        If (myScrollBar IsNot Nothing) Then
    '            For Each ctl In Me.Controls
    '                If Not (TypeOf ctl Is VScrollBar) Then
    '                    ctl.Top = ctl.Top + oldPos - myScrollBar.Value
    '                End If
    '            Next
    '        End If
    '        ol = myScrollBar.Value
    '    End If
    'End Sub

    Private Sub vScrollBar1_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
        Panel1.VerticalScroll.Value = e.NewValue
    End Sub

#End Region

#Region "Grid Events"
    Private Sub dgvProjectDetails_ColumnHeaderMouseClick(ByVal sender As Object,
        ByVal e As DataGridViewCellMouseEventArgs) _
        Handles dgvProjectDetails.ColumnHeaderMouseClick

        'REVISAR REVISAR SE PUEDE HACER

        'Dim ds = LikeSession.dsDgvProjectDetails
        'Dim dt = If(ds IsNot Nothing, ds.Tables(0), Nothing)

        'If dt IsNot Nothing Then

        '    If dt.Rows.Count > 10 Then
        '        toPaginateDs(dgvProjectDetails, ds)
        '    Else
        '        dgvProjectDetails.DataSource = dt
        '        dgvProjectDetails.Refresh()
        '    End If

        '    'dgvProjectDetails.DataSource = dt
        '    'dgvProjectDetails.Refresh()
        'End If

        'Dim newColumn As DataGridViewColumn =
        '    dgvProjectDetails.Columns(e.ColumnIndex)
        'Dim oldColumn As DataGridViewColumn = dgvProjectDetails.SortedColumn
        'Dim direction As ListSortDirection

        '' If oldColumn is null, then the DataGridView is not currently sorted. 
        'If oldColumn IsNot Nothing Then

        '    ' Sort the same column again, reversing the SortOrder. 
        '    If oldColumn Is newColumn AndAlso dgvProjectDetails.SortOrder =
        '        SortOrder.Ascending Then
        '        direction = ListSortDirection.Descending
        '        ' Msgbox HERE
        '    Else

        '        ' Sort a new column and remove the old SortGlyph.
        '        direction = ListSortDirection.Ascending
        '        oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None
        '        ' Msgbox HERE
        '    End If
        'Else
        '    direction = ListSortDirection.Ascending
        '    ' Msgbox HERE
        'End If

        '' Sort the selected column.
        'dgvProjectDetails.Sort(newColumn, direction)
        'If direction = ListSortDirection.Ascending Then
        '    newColumn.HeaderCell.SortGlyphDirection = SortOrder.Ascending
        'Else
        '    newColumn.HeaderCell.SortGlyphDirection = SortOrder.Descending
        'End If

    End Sub

    Private Sub fillcell1(strwhere As String, flag As Integer)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRDATE DESC"

            'get the query results
            ds = gnr.FillGrid(sql)
            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 6

                    'Add Columns
                    DataGridView1.Columns(0).Name = "ProjectNo"
                    DataGridView1.Columns(0).HeaderText = "Project No."
                    DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    DataGridView1.Columns(1).Name = "ProjectName"
                    DataGridView1.Columns(1).HeaderText = "Project Name"
                    DataGridView1.Columns(1).DataPropertyName = "PRNAME"

                    DataGridView1.Columns(2).Name = "DateEnt"
                    DataGridView1.Columns(2).HeaderText = "Date Entered"
                    DataGridView1.Columns(2).DataPropertyName = "PRDATE"

                    DataGridView1.Columns(3).Name = "PersonInCharge"
                    DataGridView1.Columns(3).HeaderText = "Person In Charge"
                    DataGridView1.Columns(3).DataPropertyName = "PRPECH"

                    DataGridView1.Columns(4).Name = "Status"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

                    DataGridView1.Columns(5).Name = "hasDoc"
                    DataGridView1.Columns(5).HeaderText = "Has Documents"
                    DataGridView1.Columns(5).DataPropertyName = ""


                    'fill second tab if one record in datagrid
                    If ds.Tables(0).Rows.Count = 1 Then

                        If Not String.IsNullOrEmpty(txtsearchcode.Text) Then
                            If GetAmountOfProjectReferences(txtsearchcode.Text) = 1 Then
                                fillSecondTabUpp(txtsearchcode.Text)
                                fillcell2(txtsearchcode.Text)
                                'x = dgvName.Rows(yourRowIndex).Cells(yourColumnIndex).Value
                                fillTab3(txtsearchcode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                                SSTab1.SelectedIndex = 2
                            Else
                                fillSecondTabUpp(txtsearchcode.Text)
                                fillcell2(txtsearchcode.Text)
                                SSTab1.SelectedIndex = 1
                            End If
                        Else
                            Dim grvCode = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                            If GetAmountOfProjectReferences(grvCode) = 1 Then
                                fillSecondTabUpp(grvCode)
                                fillcell2(grvCode)
                                fillTab3(grvCode, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                                SSTab1.SelectedIndex = 2
                            Else
                                fillSecondTabUpp(grvCode)
                                fillcell2(grvCode)
                                SSTab1.SelectedIndex = 1
                            End If
                        End If
                    End If

                    'FILL GRID
                    LikeSession.dsDatagridview1 = ds
                    If ds.Tables(0).Rows.Count > 10 Then
                        toPaginateDs(DataGridView1, ds)
                    Else
                        DataGridView1.DataSource = ds.Tables(0)
                        DataGridView1.Refresh()
                    End If
                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    If flag = 0 Then
                        Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    End If
                    Exit Sub
                End If
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                If flag = 0 Then
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub fillcell1LastOne(strwhere)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT * FROM PRDVLH " & strwhere & " ORDER BY PRHCOD DESC FETCH FIRST 1 ROW ONLY"   'DELETE BURNED REFERENCE
            'get the query results

            ds = gnr.FillGrid(sql)
            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 5

                    'Add Columns
                    DataGridView1.Columns(0).Name = "ProjectNo"
                    DataGridView1.Columns(0).HeaderText = "Project No."
                    DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    DataGridView1.Columns(1).Name = "ProjectName"
                    DataGridView1.Columns(1).HeaderText = "Project Name"
                    DataGridView1.Columns(1).DataPropertyName = "PRNAME"

                    DataGridView1.Columns(2).Name = "DateEnt"
                    DataGridView1.Columns(2).HeaderText = "Date Entered"
                    DataGridView1.Columns(2).DataPropertyName = "PRDATE"

                    DataGridView1.Columns(3).Name = "PersonInCharge"
                    DataGridView1.Columns(3).HeaderText = "Person In Charge"
                    DataGridView1.Columns(3).DataPropertyName = "PRPECH"

                    DataGridView1.Columns(4).Name = "Status"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

                    'FILL GRID
                    'DataGridView1.DataSource = ds.Tables(0)
                    LikeSession.dsDatagridview1 = ds
                    If ds.Tables(0).Rows.Count > 10 Then
                        toPaginateDs(DataGridView1, ds)
                    Else
                        DataGridView1.DataSource = ds.Tables(0)
                        DataGridView1.Refresh()
                    End If
                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub fillcell2(code As String, Optional ByVal strwhere As String = Nothing, Optional ByVal status As String = Nothing, Optional ByVal sessionFlag As Boolean = False)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture
            sql = "SELECT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
                    Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2
                    ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A1.PRHCOD = " & code & strwhere & " "

            'sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS,PRDJIRA,PRDUSR FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "  'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    dgvProjectDetails.DataSource = Nothing
                    dgvProjectDetails.Refresh()
                    dgvProjectDetails.AutoGenerateColumns = False
                    dgvProjectDetails.ColumnCount = 9


                    'Add Columns
                    dgvProjectDetails.Columns(0).Name = "Date"
                    dgvProjectDetails.Columns(0).HeaderText = "Date"
                    dgvProjectDetails.Columns(0).DataPropertyName = "PRDDAT"

                    dgvProjectDetails.Columns(1).Name = "PartNo"
                    dgvProjectDetails.Columns(1).HeaderText = "Part#"
                    dgvProjectDetails.Columns(1).DataPropertyName = "PRDPTN"

                    dgvProjectDetails.Columns(2).Name = "CTPNo"
                    dgvProjectDetails.Columns(2).HeaderText = "CTP#"
                    dgvProjectDetails.Columns(2).DataPropertyName = "PRDCTP"

                    dgvProjectDetails.Columns(3).Name = "MFRNo"
                    dgvProjectDetails.Columns(3).HeaderText = "MFR#"
                    dgvProjectDetails.Columns(3).DataPropertyName = "PRDMFR#"

                    dgvProjectDetails.Columns(4).Name = "Vendor"
                    dgvProjectDetails.Columns(4).HeaderText = "Vendor"
                    dgvProjectDetails.Columns(4).DataPropertyName = "VMVNUM"

                    dgvProjectDetails.Columns(5).Name = "VendorName"
                    dgvProjectDetails.Columns(5).HeaderText = "Vendor Name"
                    dgvProjectDetails.Columns(5).DataPropertyName = "VMNAME"

                    dgvProjectDetails.Columns(6).Name = "Status"
                    dgvProjectDetails.Columns(6).HeaderText = "Status"
                    dgvProjectDetails.Columns(6).DataPropertyName = "PRDSTS"

                    dgvProjectDetails.Columns(7).Name = "JiraTaskColumn"
                    dgvProjectDetails.Columns(7).HeaderText = "JiraTask"
                    dgvProjectDetails.Columns(7).DataPropertyName = "PRDJIRA"

                    dgvProjectDetails.Columns(8).Name = "hasDoc2"
                    dgvProjectDetails.Columns(8).HeaderText = "Has Documents?"
                    dgvProjectDetails.Columns(8).DataPropertyName = ""

                    'dgvProjectDetails.DataSource = ds.Tables(0)
                    'LikeSession.dsDgvProjectDetails = ds
                    If ds.Tables(0).Rows.Count > 10 Then
                        toPaginateDs(dgvProjectDetails, ds)
                    Else
                        dgvProjectDetails.DataSource = ds.Tables(0)
                        dgvProjectDetails.Refresh()
                    End If

                    If sessionFlag Then
                        LikeSession.dsDgvProjectDetails = ds
                    End If

                    'fill third tab if one record in datagrid
                    If ds.Tables(0).Rows.Count = 1 Then
                        fillTab3(code, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                        SSTab1.SelectedIndex = 2
                    End If
                Else
                    If SSTab1.SelectedIndex = 0 Then
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                        Exit Sub
                    Else
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        Dim resultAlert As DialogResult = MessageBox.Show("This project does not have parts.", "CTP System", MessageBoxButtons.YesNo)
                        If resultAlert = DialogResult.Yes Then
                            SSTab1.SelectedIndex = 2
                        End If
                        Exit Sub
                    End If
                End If
            Else
                If SSTab1.SelectedIndex = 0 Then
                    dgvProjectDetails.DataSource = Nothing
                    dgvProjectDetails.Refresh()
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                Else
                    dgvProjectDetails.DataSource = Nothing
                    dgvProjectDetails.Refresh()
                    Dim resultAlert As DialogResult = MessageBox.Show("This search reference is not present in the current project.", "CTP System", MessageBoxButtons.OK)
                    'If resultAlert = DialogResult.Yes Then
                    '    SSTab1.SelectedIndex = 2
                    'End If
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            dgvProjectDetails.DataSource = Nothing
            dgvProjectDetails.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub fillcell22(code As String)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            sql = "SELECT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(PRDVLD.VMVNUM) as VMVNUM,Trim(VMNAME) as VMNAME,
                    Trim(PRDSTS) as PRDSTS FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "
            ds = gnr.FillGrid(sql)

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    dgvProjectDetails.DataSource = Nothing
                    dgvProjectDetails.Refresh()
                    dgvProjectDetails.AutoGenerateColumns = False
                    'dgvProjectDetails.ColumnCount = 8

                    'Add Columns
                    dgvProjectDetails.Columns(0).Name = "Date"
                    dgvProjectDetails.Columns(0).HeaderText = "Date"
                    dgvProjectDetails.Columns(0).DataPropertyName = "PRDDAT"

                    dgvProjectDetails.Columns(1).Name = "PartNo"
                    dgvProjectDetails.Columns(1).HeaderText = "Part#"
                    dgvProjectDetails.Columns(1).DataPropertyName = "PRDPTN"

                    dgvProjectDetails.Columns(2).Name = "CTPNo"
                    dgvProjectDetails.Columns(2).HeaderText = "CTP#"
                    dgvProjectDetails.Columns(2).DataPropertyName = "PRDCTP"

                    dgvProjectDetails.Columns(3).Name = "MFRNo"
                    dgvProjectDetails.Columns(3).HeaderText = "MFR#"
                    dgvProjectDetails.Columns(3).DataPropertyName = "PRDMFR#"

                    dgvProjectDetails.Columns(4).Name = "Vendor"
                    dgvProjectDetails.Columns(4).HeaderText = "Vendor"
                    dgvProjectDetails.Columns(4).DataPropertyName = "VMVNUM"

                    dgvProjectDetails.Columns(5).Name = "VendorName"
                    dgvProjectDetails.Columns(5).HeaderText = "Vendor Name"
                    dgvProjectDetails.Columns(5).DataPropertyName = "VMNAME"

                    dgvProjectDetails.Columns(6).Name = "Status"
                    dgvProjectDetails.Columns(6).HeaderText = "Status"
                    dgvProjectDetails.Columns(6).DataPropertyName = "PRDSTS"

                    dgvProjectDetails.Columns(8).Name = "hasDoc2"
                    dgvProjectDetails.Columns(8).HeaderText = "Has Documents?"
                    dgvProjectDetails.Columns(8).DataPropertyName = ""

                    'FILL GRID
                    dgvProjectDetails.DataSource = ds.Tables(0)
                    'dgvProjectDetails_DataBindingComplete(Nothing, Nothing)
                Else
                    dgvProjectDetails.DataSource = Nothing
                    dgvProjectDetails.Refresh()
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    Exit Sub
                End If
            Else
                dgvProjectDetails.DataSource = Nothing
                dgvProjectDetails.Refresh()
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                Exit Sub
            End If
        Catch ex As Exception
            dgvProjectDetails.DataSource = Nothing
            dgvProjectDetails.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub fillcelldetail(strwhere As String, flag As Integer, Optional ByVal strExtraUnion As String = Nothing, Optional ByVal sessionFlag As Boolean = False)
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim ds1 As New DataSet()
        ds1.Locale = CultureInfo.InvariantCulture

        Try
            sql = "SELECT distinct(A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD " & strwhere & " ORDER BY 1 DESC"

            ds = gnr.FillGrid(sql)

            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 6

                    'Add Columns
                    DataGridView1.Columns(0).Name = "ProjectNo"
                    DataGridView1.Columns(0).HeaderText = "Project No."
                    DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    DataGridView1.Columns(1).Name = "ProjectName"
                    DataGridView1.Columns(1).HeaderText = "Project Name"
                    DataGridView1.Columns(1).DataPropertyName = "PRNAME"

                    DataGridView1.Columns(2).Name = "DateEnt"
                    DataGridView1.Columns(2).HeaderText = "Date Entered"
                    DataGridView1.Columns(2).DataPropertyName = "PRDATE"

                    DataGridView1.Columns(3).Name = "PersonInCharge"
                    DataGridView1.Columns(3).HeaderText = "Person In Charge"
                    DataGridView1.Columns(3).DataPropertyName = "PRPECH"

                    DataGridView1.Columns(4).Name = "Status"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

                    DataGridView1.Columns(5).Name = "hasDoc"
                    DataGridView1.Columns(5).HeaderText = "Has Documents"
                    DataGridView1.Columns(5).DataPropertyName = ""

                    cleanFormValues("TabPage2", 0)
                    cleanFormValues("TabPage3", 1)

                    'fill second tab if one record in datagrid
                    If ds.Tables(0).Rows.Count = 1 Then

                        Dim Regex = New Regex("WHERE")
                        Dim newStrWhere = If(strwhere <> Nothing, Regex.Replace(UCase(strwhere), " AND ", 1), Nothing)
                        Dim newstrExtraTab2 = If(strExtraUnion <> Nothing, Regex.Replace(UCase(strExtraUnion), " AND ", 1), Nothing)

                        If Not String.IsNullOrEmpty(txtsearchcode.Text) Then
                            If GetAmountOfProjectReferences(txtsearchcode.Text) = 1 Then
                                fillSecondTabUpp(txtsearchcode.Text)
                                fillcell2(txtsearchcode.Text, newstrExtraTab2, Nothing, sessionFlag)
                                If dgvProjectDetails.DataSource IsNot Nothing Then
                                    fillTab3(txtsearchcode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                                    flagdeve = 0
                                    flagnewpart = 0
                                    SSTab1.SelectedIndex = 2
                                End If
                            Else
                                fillSecondTabUpp(txtsearchcode.Text)
                                SSTab1.SelectedIndex = 1
                                fillcell2(txtsearchcode.Text, newstrExtraTab2, Nothing, sessionFlag)
                                showTab2FilterPanel(dgvProjectDetails)

                                'Dim sql1 = "SELECT A2.* FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD " & strwhere & " ORDER BY A1.PRDATE DESC"
                                'ds1 = gnr.FillGrid(sql1)

                                'If ds1 IsNot Nothing Then
                                '    If ds1.Tables(0).Rows.Count >= 2 Then
                                '        SSTab1.SelectedIndex = 1
                                '    ElseIf ds1.Tables(0).Rows.Count > 0 Then
                                '        Dim ProjectNo = ds1.Tables(0).Rows(0).ItemArray(0).ToString()
                                '        Dim partNo = ds1.Tables(0).Rows(0).ItemArray(1).ToString()
                                '        fillTab3(ProjectNo, partNo)
                                '        SSTab1.SelectedIndex = 2
                                '    End If
                                'End If
                            End If
                        Else
                            Dim grvCode = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                            If GetAmountOfProjectReferences(grvCode) = 1 Then
                                fillSecondTabUpp(grvCode)
                                fillcell2(grvCode, newstrExtraTab2, Nothing, sessionFlag)
                                If dgvProjectDetails.DataSource IsNot Nothing Then
                                    fillTab3(txtsearchcode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                                    flagdeve = 0
                                    flagnewpart = 0
                                    SSTab1.SelectedIndex = 2
                                End If
                            Else
                                fillSecondTabUpp(grvCode)
                                SSTab1.SelectedIndex = 1
                                fillcell2(grvCode, newstrExtraTab2, Nothing, sessionFlag)
                                showTab2FilterPanel(dgvProjectDetails)

                                'AQUI REVISAR EN CASO DE purc EL UNION FALLA 
                                'Dim sql1 = "SELECT A2.* FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD " & strwhere & " ORDER BY A1.PRDATE DESC"
                                'ds1 = gnr.FillGrid(sql1)

                                'If ds1 IsNot Nothing Then
                                '    If ds1.Tables(0).Rows.Count >= 2 Then
                                '        SSTab1.SelectedIndex = 1
                                '    ElseIf ds1.Tables(0).Rows.Count > 0 Then
                                '        Dim projectNo = ds1.Tables(0).Rows(0).ItemArray(0).ToString()
                                '        Dim partNo = ds1.Tables(0).Rows(0).ItemArray(1).ToString()
                                '        fillTab3(projectNo, partNo)
                                '        SSTab1.SelectedIndex = 2
                                '    End If
                                'End If
                            End If
                        End If
                    End If

                    'FILL GRID
                    'DataGridView1.DataSource = ds.Tables(0)
                    LikeSession.dsDatagridview1 = ds
                    If ds.Tables(0).Rows.Count > 10 Then
                        toPaginateDs(DataGridView1, ds)
                    Else
                        DataGridView1.DataSource = ds.Tables(0)
                        'DataGridView1.Refresh()
                    End If
                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    If flag = 0 Then
                        Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    End If
                    Exit Sub
                End If
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()
                If flag = 0 Then
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        Exit Sub
    End Sub

    Private Function fillcelldetailOther(strwhere) As Data.DataSet
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Try
            sql = "SELECT distinct(prdvlh.prhcod),prdptn,prname,prdate,prpech,prstat FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD " & strwhere & " ORDER BY PRDATE DESC"

            ds = gnr.FillGrid(sql)

            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 5

                    'Add Columns
                    DataGridView1.Columns(0).Name = "ProjectNo"
                    DataGridView1.Columns(0).HeaderText = "Project No."
                    DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    DataGridView1.Columns(1).Name = "ProjectName"
                    DataGridView1.Columns(1).HeaderText = "Project Name"
                    DataGridView1.Columns(1).DataPropertyName = "PRNAME"

                    DataGridView1.Columns(2).Name = "DateEnt"
                    DataGridView1.Columns(2).HeaderText = "Date Entered"
                    DataGridView1.Columns(2).DataPropertyName = "PRDATE"

                    DataGridView1.Columns(3).Name = "PersonInCharge"
                    DataGridView1.Columns(3).HeaderText = "Person In Charge"
                    DataGridView1.Columns(3).DataPropertyName = "PRPECH"

                    DataGridView1.Columns(4).Name = "Status"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

                    'FILL GRID
                    DataGridView1.DataSource = ds.Tables(0)
                    Return ds
                Else
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()

                    ds = Nothing
                    Return ds
                    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()

                ds = Nothing
                Return ds
                Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())

            ds = Nothing
            Return ds
        End Try

    End Function

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    Handles DataGridView1.CellFormatting
        Dim exMessage As String = Nothing
        Try
            Dim CurrentState As String = ""
            If e.ColumnIndex = 4 Then
                If e.Value IsNot Nothing Then
                    CurrentState = e.Value.ToString
                    Dim strResult = checkPendingReferences(DataGridView1.Rows(e.RowIndex).Cells("ProjectNo").Value)
                    If strResult <> CurrentState Then
                        CurrentState = strResult
                    End If
                    If CurrentState = "I" Then
                        DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "In Process"
                    ElseIf CurrentState = "F" Then
                        e.CellStyle.ForeColor = Color.Red
                        e.Value = "Finished"
                        'DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "Finished"
                    End If
                End If
            ElseIf e.ColumnIndex = 5 Then
                If e.Value Is Nothing Then
                    Dim projectNo = DataGridView1.Rows(e.RowIndex).Cells("ProjectNo").Value
                    If checkIfDocsPresent(projectNo, 0) Then
                        e.CellStyle.ForeColor = Color.Green
                        DataGridView1.Rows(e.RowIndex).Cells("hasDoc").Value = "Yes"
                        'e.Value = checkIfDocsPresent(3221, 0).ToString()
                    Else
                        e.CellStyle.ForeColor = Color.Red
                        DataGridView1.Rows(e.RowIndex).Cells("hasDoc").Value = "No"
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub dgvProjectDetails_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) _
       Handles dgvProjectDetails.DataBindingComplete

        Dim CurrentState As String = " "
        Dim NewState As String = " "
        Dim exMessage As String = " "
        Try
            For Each row As DataGridViewRow In dgvProjectDetails.Rows
                CurrentState = If(row.Cells(6).Value IsNot Nothing, row.Cells(6).Value.ToString(), Nothing)
                If CurrentState IsNot Nothing Then
                    If CurrentState.Length <= 4 Then
                        NewState = gnr.GetProjectStatusDescription(CurrentState)
                        row.Cells(6).Value = NewState
                    Else
                        Exit For
                    End If
                End If
            Next

            For Each column As DataGridViewColumn In dgvProjectDetails.Columns
                column.SortMode = DataGridViewColumnSortMode.Programmatic
            Next

            dgvProjectDetails.AutoResizeColumns()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub fillTab3(code As String, partNoCustom As String)
        Dim exMessage As String = " "
        Dim partDescription As String
        Dim ds As New DataSet()
        Dim ds1 As New DataSet()
        Dim columnToChange As String = "DVPRMG"
        Dim columnToChange1 As String = "VMNAME"
        Try
            Dim part As String = partNoCustom
            ds = gnr.GetDataByCodeAndPartNo(code, part)
            partDescription = gnr.GetDataByPartNo(part)
            ds1 = gnr.GetDataByPartNo2(part)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    SSTab1.SelectedTab = TabPage3
                    Dim partNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                    SSTab1.TabPages(2).Text = "Part No." & Trim(partNo)
                    For Each RowDs In ds.Tables(0).Rows

                        Dim CleanDateString As String = Regex.Replace(RowDs.Item(2).ToString(), "/[^0-9a-zA-Z:]/g", "")
                        Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                        DTPicker2.Value = dtChange.ToShortDateString()

                        Dim CleanDateString1 As String = Regex.Replace(RowDs.Item(ds.Tables(0).Columns("PODATE").Ordinal).ToString(), "/[^0-9a-zA-Z:]/g", "")
                        Dim dtChange1 As DateTime = DateTime.Parse(CleanDateString1)
                        DTPicker3.Value = dtChange1.ToShortDateString()

                        Dim CleanDateString2 As String = Regex.Replace(RowDs.Item(ds.Tables(0).Columns("PRDEDD").Ordinal).ToString(), "/[^0-9a-zA-Z:]/g", "")
                        Dim dtChange2 As DateTime = DateTime.Parse(CleanDateString2)
                        DTPicker4.Value = dtChange2.ToShortDateString()

                        Dim CleanDateString3 As String = Regex.Replace(RowDs.Item(ds.Tables(0).Columns("PRDERD").Ordinal).ToString(), "/[^0-9a-zA-Z:]/g", "")
                        Dim dtChange3 As DateTime = DateTime.Parse(CleanDateString3)
                        DTPicker5.Value = dtChange3.ToShortDateString()

                        txtvendorno.Text = RowDs.Item(ds.Tables(0).Columns("VMVNUM").Ordinal).ToString()
                        txtvendorname.Text = RowDs.Item(ds.Tables(0).Columns("VMNAME").Ordinal).ToString()
                        txtpartno.Text = RowDs.Item(ds.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                        txtctpno.Text = RowDs.Item(ds.Tables(0).Columns("PRDCTP").Ordinal).ToString()
                        txtqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDQTY").Ordinal).ToString()
                        txtmfrno.Text = RowDs.Item(ds.Tables(0).Columns("PRDMFR#").Ordinal).ToString()
                        txtsampleqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDSQTY").Ordinal).ToString()
                        txtminqty.Text = If(String.IsNullOrEmpty(gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text)), 0, gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text))
                        'txtminqty.Text = RowDs.Item(ds.Tables(0).Columns("PQMIN").Ordinal).ToString()
                        Dim unitCostNew = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDCON").Ordinal).ToString()), 5)
                        txtunitcostnew.Text = If(unitCostNew <> 0, String.Format("{0:0.00}", unitCostNew), "0")
                        Dim unitCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDCOS").Ordinal).ToString()), 5)
                        txtunitcost.Text = If(unitCost <> 0, String.Format("{0:0.00}", unitCost), "0")
                        Dim sampleCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDSCO").Ordinal).ToString()), 5)
                        txtsample.Text = If(sampleCost <> 0, String.Format("{0:0.00}", sampleCost), "0")
                        Dim miscCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDTTC").Ordinal).ToString()), 5)
                        txttcost.Text = If(miscCost <> 0, String.Format("{0:0.00}", miscCost), "0")
                        Dim toolingCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDTCO").Ordinal).ToString()), 5)
                        txttoocost.Text = If(toolingCost <> 0, String.Format("{0:0.00}", toolingCost), "0")
                        txtpo.Text = RowDs.Item(ds.Tables(0).Columns("PRDPO#").Ordinal).ToString()
                        txtBenefits.Text = RowDs.Item(ds.Tables(0).Columns("PRDBEN").Ordinal).ToString()
                        txtcomm.Text = RowDs.Item(ds.Tables(0).Columns("PRDINF").Ordinal).ToString()

                        'prevetn to get min qty from database
                        'txtminqty.Text = gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text)
                        'Dim minQtyValue As Integer = 0
                        'txtminqty.Text = minQtyValue.ToString()

                        flagdeve = 0
                        flagnewpart = 0

                        If cmbuser.FindStringExact(Trim(RowDs.Item(18).ToString())) Then
                            cmbuser.SelectedIndex = cmbuser.FindString(Trim(RowDs.Item(18).ToString()))
                        End If

                        Dim posValue As Integer = 0
                        For Each obj As DataRowView In cmbstatus.Items
                            Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDSTS").Ordinal).ToString())
                            Dim VarCombo = Trim(obj.Item(0).ToString())
                            If VarQuery = VarCombo Then
                                cmbstatus.SelectedIndex = posValue
                                Exit For
                            Else
                                posValue += 1
                            End If
                        Next

                        Dim posValueMin As Integer = 0
                        For Each obj As DataRowView In cmbminorcode.Items
                            Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDMPC").Ordinal).ToString())
                            Dim VarCombo = Trim(obj.Item(2).ToString())
                            If VarQuery = VarCombo Then
                                cmbminorcode.SelectedIndex = posValueMin
                                Exit For
                            Else
                                posValueMin += 1
                            End If
                        Next

                        txtpartdescription.Text = partDescription

                        Dim rdValue = RowDs.Item(ds.Tables(0).Columns("PRDPTS").Ordinal).ToString()
                        If rdValue = "1" Then
                            optCTP.Checked = True
                            optVENDOR.Checked = False
                            optboth.Checked = False
                        ElseIf rdValue = "2" Then
                            optCTP.Checked = False
                            optVENDOR.Checked = True
                            optboth.Checked = False
                        ElseIf rdValue = "" Then
                            optCTP.Checked = False
                            optVENDOR.Checked = False
                            optboth.Checked = True
                        End If

                        searchpart()

                        'new item or new supplier
                        chknew.Checked = False
                        chkSupplier.Checked = False
                        chknew.Checked = If(itemCategory(txtpartno.Text, txtvendorno.Text) = 2, True, False)
                        chkSupplier.Checked = If(chknew.Checked, False, True)
                        'If chknew.Checked Then
                        '    chkSupplier.Checked = Not chknew.Checked
                        'End If

                    Next

                    If cmbuser.SelectedIndex = -1 Then
                        cmbuser.SelectedIndex = cmbuser1.Items.Count - 1
                    End If

                    If ds1 IsNot Nothing Then
                        If ds1.Tables(0).Rows.Count > 0 Then
                            Dim ctIndex = ds1.Tables(0).Columns(columnToChange).Ordinal
                            Dim ctIndex1 = ds1.Tables(0).Columns(columnToChange1).Ordinal
                            txtvendornoa.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex).ToString()
                            txtvendornamea.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex1).ToString()
                        End If
                    End If

                    Dim jiraValue = Trim(ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDJIRA").Ordinal).ToString())
                    If Not String.IsNullOrEmpty(jiraValue) Then
                        txtjiratask.Text = jiraValue
                        Cmdjira.Visible = True
                    Else
                        lbljiratask.Visible = False
                        txtjiratask.Visible = False
                        Cmdjira.Visible = False
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub dgvProjectDetails_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles dgvProjectDetails.DoubleClick
        Dim Index As Integer
        Dim ds As New DataSet()
        Dim ds1 As New DataSet()
        Dim RowDs As DataRow
        ds.Locale = CultureInfo.InvariantCulture
        ds1.Locale = CultureInfo.InvariantCulture
        Dim exMessage As String = " "
        Dim code As String = txtCode.Text
        Dim partDescription As String
        Dim dtSecondTb As DataTable = New DataTable()
        Dim columnToChange As String = "DVPRMG"
        Dim columnToChange1 As String = "VMNAME"

        Try
            For Each row As DataGridViewRow In dgvProjectDetails.SelectedRows
                Index = dgvProjectDetails.CurrentCell.RowIndex
                If dgvProjectDetails.Rows(Index).Selected = True Then
                    Dim part As String = row.Cells(1).Value.ToString()
                    ds = gnr.GetDataByCodeAndPartNo(code, part)
                    partDescription = gnr.GetDataByPartNo(part)
                    ds1 = gnr.GetDataByPartNo2(part)
                    If ds IsNot Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            SSTab1.SelectedTab = TabPage3
                            Dim partNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                            SSTab1.TabPages(2).Text = "Part No." & Trim(partNo)
                            For Each RowDs In ds.Tables(0).Rows

                                Dim CleanDateString As String = Regex.Replace(RowDs.Item(2).ToString(), "/[^0-9a-zA-Z:]/g", "")
                                Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                                DTPicker2.Value = dtChange.ToShortDateString()

                                txtvendorno.Text = RowDs.Item(ds.Tables(0).Columns("VMVNUM").Ordinal).ToString()
                                txtvendorname.Text = RowDs.Item(ds.Tables(0).Columns("VMNAME").Ordinal).ToString()
                                txtpartno.Text = RowDs.Item(ds.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                                txtctpno.Text = RowDs.Item(ds.Tables(0).Columns("PRDCTP").Ordinal).ToString()
                                txtqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDQTY").Ordinal).ToString()
                                'Dim qtyValue As Integer = 0
                                'txtqty.Text = qtyValue.ToString()
                                txtmfrno.Text = RowDs.Item(ds.Tables(0).Columns("PRDMFR#").Ordinal).ToString()
                                txtsampleqty.Text = RowDs.Item(ds.Tables(0).Columns("PRDSQTY").Ordinal).ToString()
                                'txtminqty.Text = RowDs.Item(ds.Tables(0).Columns("PQMIN").Ordinal).ToString()
                                Dim unitCostNew = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDCON").Ordinal).ToString()), 5)
                                txtunitcostnew.Text = If(unitCostNew <> 0, String.Format("{0:0.00}", unitCostNew), "0")
                                Dim unitCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDCOS").Ordinal).ToString()), 5)
                                txtunitcost.Text = If(unitCost <> 0, String.Format("{0:0.00}", unitCost), "0")
                                Dim sampleCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDSCO").Ordinal).ToString()), 5)
                                txtsample.Text = If(sampleCost <> 0, String.Format("{0:0.00}", sampleCost), "0")
                                Dim miscCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDTTC").Ordinal).ToString()), 5)
                                txttcost.Text = If(miscCost <> 0, String.Format("{0:0.00}", miscCost), "0")
                                Dim toolingCost = Math.Round(CDbl(RowDs.Item(ds.Tables(0).Columns("PRDTCO").Ordinal).ToString()), 5)
                                txttoocost.Text = If(toolingCost <> 0, String.Format("{0:0.00}", toolingCost), "0")
                                txtpo.Text = RowDs.Item(ds.Tables(0).Columns("PRDPO#").Ordinal).ToString()
                                txtBenefits.Text = RowDs.Item(ds.Tables(0).Columns("PRDBEN").Ordinal).ToString()

                                'prevent to get the min qty from database
                                txtminqty.Text = If(String.IsNullOrEmpty(gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text)), 0, gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text))
                                'txtminqty.Text = gnr.GetDataByVendorAndPartNo(txtvendorno.Text, txtpartno.Text)
                                'Dim minQtyValue As Integer = 0
                                'txtminqty.Text = minQtyValue.ToString()

                                'new item or new supplier
                                chknew.Checked = False
                                chkSupplier.Checked = False
                                chknew.Checked = If(itemCategory(txtpartno.Text, txtvendorno.Text) = 2, True, False)
                                chkSupplier.Checked = If(chknew.Checked, False, True)
                                'If chknew.Checked Then
                                '    chkSupplier.Checked = Not chknew.Checked
                                'End If

                                flagdeve = 0
                                flagnewpart = 0

                                If cmbuser.FindStringExact(Trim(RowDs.Item(18).ToString())) Then
                                    cmbuser.SelectedIndex = cmbuser.FindString(Trim(RowDs.Item(18).ToString()))
                                End If

                                Dim posValue As Integer = 0
                                For Each obj As DataRowView In cmbstatus.Items
                                    Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDSTS").Ordinal).ToString())
                                    Dim VarCombo = Trim(obj.Item(0).ToString())
                                    If VarQuery = VarCombo Then
                                        cmbstatus.SelectedIndex = posValue
                                        Exit For
                                    Else
                                        posValue += 1
                                    End If
                                Next

                                Dim posValueMin As Integer = 0
                                For Each obj As DataRowView In cmbminorcode.Items
                                    Dim VarQuery = Trim(RowDs.Item(ds.Tables(0).Columns("PRDMPC").Ordinal).ToString())
                                    Dim VarCombo = Trim(obj.Item(2).ToString())
                                    If VarQuery = VarCombo Then
                                        cmbminorcode.SelectedIndex = posValueMin
                                        Exit For
                                    Else
                                        posValueMin += 1
                                    End If
                                Next

                                txtpartdescription.Text = partDescription

                                Dim rdValue = RowDs.Item(ds.Tables(0).Columns("PRDPTS").Ordinal).ToString()
                                If rdValue = "1" Then
                                    optCTP.Checked = True
                                    optVENDOR.Checked = False
                                    optboth.Checked = False
                                ElseIf rdValue = "2" Then
                                    optCTP.Checked = False
                                    optVENDOR.Checked = True
                                    optboth.Checked = False
                                ElseIf rdValue = "" Then
                                    optCTP.Checked = False
                                    optVENDOR.Checked = False
                                    optboth.Checked = True
                                End If

                            Next

                            If cmbuser.SelectedIndex = -1 Then
                                cmbuser.SelectedIndex = cmbuser1.Items.Count - 1
                            End If

                            If ds1 IsNot Nothing Then
                                If ds1.Tables(0).Rows.Count > 0 Then
                                    Dim ctIndex = ds1.Tables(0).Columns(columnToChange).Ordinal
                                    Dim ctIndex1 = ds1.Tables(0).Columns(columnToChange1).Ordinal
                                    txtvendornoa.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex).ToString()
                                    txtvendornamea.Text = ds1.Tables(0).Rows(0).ItemArray(ctIndex1).ToString()
                                End If
                            End If

                            Dim jiraValue = Trim(ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDJIRA").Ordinal).ToString())
                            If Not String.IsNullOrEmpty(jiraValue) Then
                                txtjiratask.Text = jiraValue
                                Cmdjira.Visible = True
                            Else
                                txtjiratask.Visible = False
                                lbljiratask.Visible = False
                                Cmdjira.Visible = False
                            End If

                            searchpart()

                        End If
                    End If
                End If
            Next

            changeControlAccess(True)

            cmbminorcode.Enabled = False
            txtminor.Enabled = False
            cmbmajorcode.Enabled = False
            txtMajor.Enabled = False
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView1.DoubleClick
        Dim Index As Integer
        Dim ds As New DataSet
        Dim RowDs As DataRow
        ds.Locale = CultureInfo.InvariantCulture
        Dim exMessage As String = " "
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                Index = DataGridView1.CurrentCell.RowIndex
                If DataGridView1.Rows(Index).Selected = True Then
                    Dim code As String = row.Cells(0).Value.ToString()

                    ds = gnr.GetDataByPRHCOD(code)
                    If ds.Tables(0).Rows.Count = 1 Then
                        forceDbClick_Action(code)
                    Else
                        'message box warning
                    End If
                Else
                    'is is not selected
                End If
            Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub dgvProjectDetails_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
    Handles dgvProjectDetails.CellContentClick
        Dim exMessage As String = Nothing
        Try
            If e.ColumnIndex = 7 And e.RowIndex > 0 Then
                Dim UrlAddress = gnr.JiraPathBaseValue + dgvProjectDetails(e.ColumnIndex, e.RowIndex).Value.ToString()
                If System.Uri.IsWellFormedUriString(UrlAddress, UriKind.Absolute) Then
                    Process.Start(UrlAddress)
                Else
                    MessageBox.Show("The url has error.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                'Dim senderGrid = DirectCast(sender, DataGridView)
                If e.RowIndex > 0 Then
                    dgvProjectDetails_DoubleClick(sender, e)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub dgvProjectDetails_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) _
    Handles dgvProjectDetails.CellFormatting
        Dim exMessage As String = Nothing
        Try
            If e.ColumnIndex = 7 Then
                If Not String.IsNullOrEmpty(Trim(e.Value)) Then
                    'DataGridView1.Rows(e.RowIndex).Cells("Status").Value = "In Process"
                    'e.Value = "Go to Jiratask"
                    e.FormattingApplied = True
                Else
                    e.Value = ""
                End If
            ElseIf e.ColumnIndex = 8 Then
                If Not String.IsNullOrEmpty(txtCode.Text) Then
                    Dim projectNo = txtCode.Text
                    Dim partNo = If(dgvProjectDetails.Rows(e.RowIndex).Cells("PartNo").Value IsNot Nothing, dgvProjectDetails.Rows(e.RowIndex).Cells("PartNo").Value.ToString(), Nothing)
                    Dim DicRefDocs = getReferenceDocuments(projectNo, partNo)
                    If DicRefDocs IsNot Nothing Then
                        For Each pair As KeyValuePair(Of String, String) In DicRefDocs
                            If pair.Value.ToString() = "True" Then
                                dgvProjectDetails.Rows(e.RowIndex).Cells("hasDoc2").Value = "Yes"
                            Else
                                dgvProjectDetails.Rows(e.RowIndex).Cells("hasDoc2").Value = "No"
                            End If
                            'dgvProjectDetails.Rows(e.RowIndex).Cells("hasDoc2").Value = pair.Value.ToString()
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Protected Sub toPaginateDs(dgv As DataGridView, ds As DataSet)
        Dim exMessage As String = " "
        Try
            Dim dtGrid As New DataTable
            dtGrid = ds.Tables(0)

            Dim counter As Integer = 0
            Dim dt As DataTable = Nothing

            If dgv.Name.ToString().Equals("DataGridView1") Then

                DataGridView1.Visible = True
                If Tables.Count > 0 Then
                    Tables = New BindingList(Of DataTable)()
                    bs.MoveFirst()
                End If

                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator1.BindingSource = bs
                bs.DataSource = Tables
                AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
                bs_PositionChanged(bs, Nothing)
            Else
                dgvProjectDetails.Visible = True
                If Tables1.Count > 0 Then
                    Tables1 = New BindingList(Of DataTable)()
                    bs1.MoveFirst()
                End If

                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables1.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator2.BindingSource = bs1
                bs1.DataSource = Tables1
                AddHandler bs1.PositionChanged, AddressOf bs1_PositionChanged
                bs1_PositionChanged(bs1, Nothing)

            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            'Log.Error(exMessage)
        End Try
    End Sub

    Protected Sub toPaginate(dgv As DataGridView)
        Dim exMessage As String = " "
        Try
            'dim tables as BindingList<DataTable>  = new BindingList<DataTable>()
            Dim dtGrid As New DataTable
            dtGrid = (DirectCast(dgv.DataSource, DataTable))

            Dim counter As Integer = 0
            Dim dt As DataTable = Nothing

            If dgv.Name.ToString().Equals("DataGridView1") Then
                dgvProjectDetails.Visible = True
                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator1.BindingSource = bs
                bs.DataSource = Tables
                AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
                bs_PositionChanged(bs, Nothing)
            Else
                dgvProjectDetails.Visible = True
                For Each item As DataRow In dtGrid.Rows
                    If counter = 0 Then
                        dt = dtGrid.Clone()
                        Tables1.Add(dt)
                    End If

                    dt.Rows.Add(item.ItemArray)
                    counter += 1

                    If counter > 9 Then
                        counter = 0
                    End If
                Next

                BindingNavigator2.BindingSource = bs1
                bs1.DataSource = Tables1
                AddHandler bs1.PositionChanged, AddressOf bs1_PositionChanged
                bs1_PositionChanged(bs1, Nothing)

            End If


        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub bs_PositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        'If DataGridView1.DataSource IsNot Nothing Then
        Dim exMessage As String = Nothing
        Try
            If bs.Position <> -1 Then
                DataGridView1.DataSource = Tables(bs.Position)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'End If
    End Sub

    Private Sub bs1_PositionChanged(ByVal sender As Object, ByVal e As EventArgs)
        'If dgvProjectDetails.DataSource IsNot Nothing Then
        Dim exMessage As String = Nothing
        Try
            If bs1.Position <> -1 Then
                dgvProjectDetails.DataSource = Tables1(bs1.Position)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'End If
    End Sub

#End Region

#Region "Textbox events"

    'Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles txtsearch.GotFocus
    '    txtsearch.SelectionStart = 0
    '    txtsearch.SelectionLength = Len(Trim(txtsearch.Text))
    'End Sub

    'Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles txtsearch1.GotFocus
    '    txtsearch1.SelectionStart = 0
    '    txtsearch1.SelectionLength = Len(Trim(txtsearch.Text))
    'End Sub

    Private Sub txtname_TextChanged(sender As Object, e As EventArgs) Handles txtname.TextChanged
        cmdSave2.Enabled = True
    End Sub

    Private Sub txtpartno_TextChanged(sender As Object, e As EventArgs) Handles txtpartno.TextChanged
        If Not String.IsNullOrEmpty(txtpartno.Text) Then
            TabPage3.Name = "Part No. " & txtpartno.Text
        End If
    End Sub

    Private Sub txtsearchcode_TextChanged(sender As Object, e As EventArgs) Handles txtsearchcode.TextChanged
        txtsearchcode.Text = txtsearchcode.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        txtsearch.Text = txtsearch.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtJiratasksearch_TextChanged(sender As Object, e As EventArgs) Handles txtJiratasksearch.TextChanged
        txtJiratasksearch.Text = txtJiratasksearch.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtsearch1_TextChanged(sender As Object, e As EventArgs) Handles txtsearch1.TextChanged
        txtsearch1.Text = txtsearch1.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtsearchpart_TextChanged(sender As Object, e As EventArgs) Handles txtsearchpart.TextChanged
        txtsearchpart.Text = txtsearchpart.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub cmbPrpech_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPrpech.SelectedIndexChanged
        txtsearchcode.Text = txtsearchcode.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub cmbstatus1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbstatus1.SelectedIndexChanged
        txtsearchcode.Text = txtsearchcode.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtsearchctp_TextChanged(sender As Object, e As EventArgs) Handles txtsearchctp.TextChanged
        txtsearchctp.Text = txtsearchctp.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub txtMfrNoSearch_TextChanged(sender As Object, e As EventArgs) Handles txtMfrNoSearch.TextChanged
        txtMfrNoSearch.Text = txtMfrNoSearch.Text.Replace(Environment.NewLine, "")
    End Sub

    Private Sub TextBox_Focus(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles txtsearch.GotFocus, txtsearch1.GotFocus, txtsearchcode.GotFocus, txtsearchctp.GotFocus, txtsearchpart.GotFocus,
        txtMfrNoSearch.GotFocus, txtJiratasksearch.GotFocus, cmbstatus1.GotFocus, cmbPrpech.GotFocus

        Dim exMessage As String = Nothing
        Try
            Dim controlSender As Object
            Dim currTextBox As System.Windows.Forms.TextBox
            Dim currComboBox As System.Windows.Forms.ComboBox
            Dim flag As Integer = 0

            Dim sender_type = sender.GetType().ToString()
            If sender_type.Equals("System.Windows.Forms.TextBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.TextBox)
            ElseIf sender_type.Equals("System.Windows.Forms.ComboBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.ComboBox)
                flag = 1
            Else
                controlSender = Nothing
            End If

            If flag = 0 Then
                currTextBox = sender
                Dim castedTextBox As Control = Nothing
                If currTextBox.Equals(txtsearch) Then
                    castedTextBox = DirectCast(txtsearch, Control)
                    LikeSession.focussedControl = castedTextBox
                ElseIf currTextBox.Equals(txtsearch1) Then
                    castedTextBox = DirectCast(txtsearch1, Control)
                    LikeSession.focussedControl = txtsearch1
                ElseIf currTextBox.Equals(txtsearchcode) Then
                    castedTextBox = DirectCast(txtsearchcode, Control)
                    LikeSession.focussedControl = txtsearchcode
                ElseIf currTextBox.Equals(txtsearchctp) Then
                    castedTextBox = DirectCast(txtsearchctp, Control)
                    LikeSession.focussedControl = txtsearchctp
                ElseIf currTextBox.Equals(txtsearchpart) Then
                    castedTextBox = DirectCast(txtsearchpart, Control)
                    LikeSession.focussedControl = txtsearchpart
                ElseIf currTextBox.Equals(txtMfrNoSearch) Then
                    castedTextBox = DirectCast(txtMfrNoSearch, Control)
                    LikeSession.focussedControl = txtMfrNoSearch
                ElseIf currTextBox.Equals(txtJiratasksearch) Then
                    castedTextBox = DirectCast(txtJiratasksearch, Control)
                    LikeSession.focussedControl = txtJiratasksearch
                End If
            Else
                currComboBox = sender
                Dim castedComboBox As Control = Nothing
                If currComboBox.Equals(cmbstatus1) Then
                    castedComboBox = DirectCast(cmbstatus1, Control)
                    LikeSession.focussedControl = cmbstatus1
                ElseIf currComboBox.Equals(cmbPrpech) Then
                    castedComboBox = DirectCast(cmbPrpech, Control)
                    LikeSession.focussedControl = cmbPrpech
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

#End Region

#Region "Button Events"

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Dim exMessage As String = Nothing
        Try
            If Application.OpenForms.OfType(Of frmLoadExcel).Any() Then
                'MessageBox.Show("The Form is already opened")
                Dim rsDialog As DialogResult = MessageBox.Show("The requeted form is already open. Do you want to reload it?", "CTP System", MessageBoxButtons.YesNo)
                If rsDialog = DialogResult.Yes Then
                    frmLoadExcel.Close()
                    frmLoadExcel.Show()
                Else
                    frmLoadExcel.BringToFront()
                End If
            Else
                frmLoadExcel.Show()
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        cmdClearFilters1_Click(sender, Nothing)
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        customMessageBox1.ShowDialog()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        cmdClearFilters_Click(sender, Nothing)
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim exMessage As String = " "
        Dim flagReference As Boolean = False
        Try
            If Not String.IsNullOrEmpty(txtsearch1.Text) Or cmbstatus1.SelectedIndex > 0 Then

                Dim lstReferenceUsers As String() = gnr.ReferenceUsersReport.Split(",")
                For Each item As String In lstReferenceUsers
                    If item = userid Then
                        flagReference = True
                        Exit For
                    End If
                Next

                'test purpose
                'flagReference = True

                Dim dsMassiveData = If(String.IsNullOrEmpty(txtsearch1.Text), gnr.GetMassiveReferences(cmbstatus1.SelectedValue, userid, flagReference), gnr.GetMassiveReferences(cmbstatus1.SelectedValue, userid, flagReference, Trim(txtsearch1.Text)))
                'gnr.GetMassiveReferences(cmbstatus1.SelectedValue, txtsearch1.Text)
                If dsMassiveData IsNot Nothing Then
                    If dsMassiveData.Tables(0).Rows.Count() > 0 Then
                        Dim vendorNo = txtsearch1.Text
                        Dim Status = Trim(cmbstatus1.GetItemText(cmbstatus1.SelectedItem(1)))
                        Dim title = "Status Report for vendor " & vendorNo & " And Status " & Status & " requested by " & userid & " running at "
                        prodDevExcelGeneration(dsMassiveData, vendorNo, Status, title)
                    Else
                        MessageBox.Show("There is not results with this vendor number and project status selected.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is not results with this vendor number and project status selected.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("To run this report you must select the vendor number and project status.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Dim exMessage As String = " "
        Dim created As String = False
        Try
            If userid IsNot Nothing Then
                If True Then

                End If
                Dim dsAlertInactives = If(String.IsNullOrEmpty(txtsearch1.Text), gnr.GetInactiveAlertByUser(userid), gnr.GetInactiveAlertByUser(userid, Trim(txtsearch1.Text)))
                'Dim dsAlertInactives = gnr.GetInactiveAlertByUser(userid)
                If dsAlertInactives IsNot Nothing Then
                    If dsAlertInactives.Tables(0).Rows.Count > 0 Then
                        Dim newDsAlertInactive = New DataSet()
                        Dim newDtAlertInactive = New DataTable()
                        newDtAlertInactive = dsAlertInactives.Tables(0).Clone()

                        For Each dtw As DataRow In dsAlertInactives.Tables(0).Rows
                            If UCase(Trim(dtw.ItemArray(8).ToString())) = userid Then
                                newDtAlertInactive.ImportRow(dtw)
                            End If
                        Next
                        newDsAlertInactive.Tables.Add(newDtAlertInactive)

                        Dim title As String
                        title = "Inactivity report for " & userid & " running at "
                        InactiveQotaAlertExcelGeneration(newDsAlertInactive, userid, created, title)
                        If Not created Then
                            MessageBox.Show("There is an error in the creation of the report.", "CTP System", MessageBoxButtons.OK)
                            'Dim result As DialogResult = MessageBox.Show("Did you want to receive an email with the oldest quotation without activity?", "CTP System", MessageBoxButtons.YesNo)
                            'If result = DialogResult.Yes Then
                            '    Dim customtoemails = prepareEmailsToSendReport(1)
                            '    Dim rsResult = gnr.sendEmail(customtoemails, userid)
                            '    If rsResult < 0 Then
                            '        MessageBox.Show("Ann error ocurred sending emails.", "CTP System", MessageBoxButtons.OK)
                            '    End If
                            'End If
                        End If
                    Else
                        MessageBox.Show("There is not results with this user.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is not results with this user.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("There is an error reaching the current logged user.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub AddNewProjectToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AddNewProjectToolStripMenuItem1.Click
        Dim exMessage As String = Nothing
        Try
            Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                flagdeve = 1
                flagnewpart = 1
                cleanFormValues("TabPage2", 2)
                cleanFormValues("TabPage3", 2)
                TableLayoutPanel15.Enabled = True
                cmdchange.Enabled = True
                cmdunitcost.Enabled = True
                cmdmpartno.Enabled = True
                cmdcvendor.Enabled = True
                cmdSave2.Enabled = True
                gotonew()
                SSTab1.SelectedIndex = 1
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub AddNewPartToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AddNewPartToolStripMenuItem1.Click
        'tab 3
        Dim exMessage As String = Nothing
        Try
            Dim result As DialogResult = MessageBox.Show("Do you want to add a new part to the project?", "CTP System", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                flagdeve = 0
                flagnewpart = 1
                cleanFormValues("TabPage3", 1)
                TableLayoutPanel15.Enabled = False
                cmdchange.Enabled = False
                cmdunitcost.Enabled = False
                cmdmpartno.Enabled = False
                cmdcvendor.Enabled = False
                setVendorValues()
                'gotonew()
                'frmProductsDevelopment_load()
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdnew3_Click(sender As Object, e As EventArgs) Handles cmdnew3.Click
        Dim exMessage As String = Nothing
        Try
            Dim screenPoint As Point = cmdnew3.PointToScreen(New Point(cmdnew3.Left, cmdnew3.Bottom))
            If screenPoint.Y & ContextMenuStrip2.Size.Height > Screen.PrimaryScreen.WorkingArea.Height Then
                ContextMenuStrip2.Show(cmdnew3, New Point(0, -ContextMenuStrip2.Size.Height))
            Else
                ContextMenuStrip2.Show(cmdnew3, New Point(0, cmdnew3.Height))
            End If

            cmdSave2.Enabled = True
            'ContextMenuStrip1.Show(cmdSplit, New Point(0, cmdSplit.Height))
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try


        'Dim validationResult = mandatoryFields("new", SSTab1.SelectedIndex)
        'If validationResult.Equals(0) Then
        '    Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.OK)
        '    flagdeve = 1
        '    flagnewpart = 1
        '    cleanFormValues("TabPage2", 2)
        '    cleanFormValues("TabPage3", 2)
        '    gotonew()

        'Else
        '    Dim resultNew As DialogResult = MessageBox.Show("You have data in the form. You could missing if continue. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo)
        '    If resultNew = DialogResult.Yes Then
        '        flagdeve = 1
        '        flagnewpart = 1
        '        cleanFormValues("TabPage2", 2)
        '        cleanFormValues("TabPage3", 2)
        '        gotonew()
        '    End If
        'End If

    End Sub

    Private Sub AddNewProjectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewProjectToolStripMenuItem.Click
        Dim exMessage As String = Nothing
        Try
            Dim validationResult = mandatoryFields("new", SSTab1.SelectedIndex, 1)
            If validationResult.Equals(0) Then
                Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    flagdeve = 1
                    flagnewpart = 1
                    cleanFormValues("TabPage2", 2)
                    TableLayoutPanel15.Enabled = True
                    cmdchange.Enabled = True
                    cmdunitcost.Enabled = True
                    cmdmpartno.Enabled = True
                    cmdcvendor.Enabled = True
                    gotonew()
                    showTab2FilterPanel(dgvProjectDetails)
                    'frmProductsDevelopment_load()
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub AddNewPartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewPartToolStripMenuItem.Click
        'tab 2
        Dim exMessage As String = " "
        Try
            flagdeve = 0
            flagnewpart = 1
            cleanFormValues("TabPage3", 1)
            TableLayoutPanel15.Enabled = False
            cmdchange.Enabled = False
            cmdunitcost.Enabled = False
            cmdmpartno.Enabled = False
            cmdcvendor.Enabled = False
            setVendorValues()
            gotonew()
            SSTab1.SelectedIndex = 2
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdnew2_Click(sender As Object, e As EventArgs) Handles cmdnew2.Click
        Dim exMessage As String = Nothing
        Try
            Dim screenPoint As Point = cmdnew2.PointToScreen(New Point(cmdnew2.Left, cmdnew2.Bottom))
            If screenPoint.Y & ContextMenuStrip1.Size.Height > Screen.PrimaryScreen.WorkingArea.Height Then
                If flagdeve = 0 Then
                    ContextMenuStrip1.Show(cmdnew2, New Point(0, -ContextMenuStrip1.Size.Height))
                End If
            Else
                If flagdeve = 0 Then
                    ContextMenuStrip1.Show(cmdnew2, New Point(0, cmdnew2.Height))
                End If
            End If
            cmdSave2.Enabled = True
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdnew1_Click(sender As Object, e As EventArgs) Handles cmdnew1.Click
        Dim exMessage As String = Nothing
        Try
            If flagdeve = 1 Then
                gotonew()
            Else
                flagdeve = 1
                flagnewpart = 1
                cleanFormValues("TabPage2", 2)
                cleanFormValues("TabPage3", 2)
                TableLayoutPanel15.Enabled = True
                cmdchange.Enabled = True
                cmdunitcost.Enabled = True
                cmdmpartno.Enabled = True
                cmdcvendor.Enabled = True
                cmdSave2.Enabled = True
                gotonew()
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
        'If result = DialogResult.No Then
        '    'MessageBox.Show("No pressed")
        'ElseIf result = DialogResult.Yes Then
        '    'MessageBox.Show("Yes pressed")
        '    flagdeve = 1
        '    flagnewpart = 1
        '    cleanFormValues("TabPage2", 2)
        '    gotonew()
        'End If
    End Sub

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
    End Sub

    Private Sub cmdexit2_Click(sender As Object, e As EventArgs) Handles cmdexit2.Click
        Me.Close()
    End Sub

    Private Sub cmdexit3_Click(sender As Object, e As EventArgs) Handles cmdexit3.Click
        Me.Close()
    End Sub

    Private Sub gotonew()
        Dim exMessage As String = Nothing
        Try
            SSTab1.SelectedTab = TabPage2
            requireValidation = 1
            cmbprstatus.SelectedIndex = 1
            'pathpictureparts = pathgeneral & "CTPPictures\pic_not_av.jpg"
            Dim pathPictures = gnr.Path & "CTPPictures\"
            If Not Directory.Exists(pathPictures) Then
                System.IO.Directory.CreateDirectory(pathPictures)
            End If
            pathpictureparts = pathPictures & "avatar-ctp.PNG"

            Dim existsFile As Boolean = File.Exists(pathpictureparts)
            If existsFile Then
                PictureBox1.Load(pathpictureparts)
            End If

            cmbuser1.SelectedIndex = If(cmbuser.FindString(Trim(UCase(userid))) <> -1,
                                    cmbuser.FindString(Trim(UCase(userid))), 0)
            'pathpictureparts = gnr.pathgeneral & "CTPPictures\pic_not_av.jpg"
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub PoQotaFunction(Status2 As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        'Dim Status2 As String = ""

        Try
            statusquote = "D-" & Status2
            Dim mpnopo As String = String.Empty
            Dim spacepoqota As String = String.Empty
            Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text)

            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    mpnopo = Trim(UCase(txtmfrno.Text))
                    Dim maxValue = 0
                    Dim dsUpdatedData As Integer

                    Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                    If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtvendorno.Text, txtpartno.Text)
                        If dsUpdatedData <> 0 Then
                            MessageBox.Show("An error ocurred updating fields.", "CTP System", MessageBoxButtons.OK)
                        End If
                    Else
                        Dim arrayCheck As New List(Of String)
                        arrayCheck = strCheckPoQoteIns.Split(",").ToList()
                        For Each item As String In arrayCheck
                            If item = "Sequencial" Then
                                'show error message
                                Exit For
                            ElseIf item = "Vendor Number" Then
                                txtvendorno.Text = "0" 'ask for vendor??
                            ElseIf item = "Unit Cost New" Then
                                txtunitcostnew.Text = "0"
                            ElseIf item = "Min Quantity" Then
                                txtminqty.Text = "0"
                            End If
                        Next
                        dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, txtunitcost.Text, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                            txtvendorno.Text, txtpartno.Text)

                        If dsUpdatedData <> 0 Then
                            'show message error
                        End If
                    End If
                Else
                    'warning message
                End If
            Else
                Dim maxValue = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                If Not String.IsNullOrEmpty(maxValue) Then
                    maxValue += 1
                Else
                    maxValue = 1
                End If
                spacepoqota = "                               DEV"
                mpnopo = Trim(UCase(txtmfrno.Text))
                Dim ResultQuery As String = String.Empty

                Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                        DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                    ResultQuery = gnr.InsertNewPOQota(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                    If ResultQuery <> 0 Then
                        'show message error
                    End If
                Else
                    Dim arrayCheck As New List(Of String)
                    arrayCheck = strCheckPoQoteIns.Split(",").ToList()
                    For Each item As String In arrayCheck
                        If item = "Sequencial" Then
                            'show error message
                            Exit For
                        ElseIf item = "Vendor Number" Then
                            txtvendorno.Text = "0"
                        ElseIf item = "Unit Cost New" Then
                            txtunitcostnew.Text = "0"
                        ElseIf item = "Min Qty" Then
                            txtminqty.Text = "0"
                        End If
                    Next

                    ResultQuery = gnr.InsertNewPOQota(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                       DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                    If ResultQuery <> 0 Then
                        'show message error
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub InsertProductDetails(projectNo As String, partstoshow As String)
        Dim dtTime As DateTimePicker = New DateTimePicker()
        Dim dtTime1 As DateTimePicker = New DateTimePicker()
        Dim dtTime2 As DateTimePicker = New DateTimePicker()
        Dim dtTime3 As DateTimePicker = New DateTimePicker()
        Dim dtTime4 As DateTimePicker = New DateTimePicker()
        Dim dtTime5 As DateTimePicker = New DateTimePicker()
        Dim QueryDetailResult As Integer = -1
        Dim exMessage As String = " "
        Try
            dtTime5.Value = New DateTime(1900, 1, 1)
            dtTime5.CustomFormat = "yyyy/MM/dd/"

            Dim strCheck = gnr.checkFields(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, txtqty.Text,
                                                                "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                                                cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
                                                                dtTime5.Value.ToShortDateString(), txtsampleqty.Text)
            If String.IsNullOrEmpty(strCheck) Then
                QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, txtqty.Text,
                                    "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                    cmbuser.SelectedValue, chknew, dtTime3, txtsample.Text, txttcost.Text, txtvendorno.Text, partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, dtTime4,
                                    dtTime5, CInt(txtsampleqty.Text))
                If QueryDetailResult <> 0 Then
                    'show message error
                End If
            Else
                Dim arrayCheck As New List(Of String)
                arrayCheck = strCheck.Split(",").ToList()
                For Each item As String In arrayCheck
                    If item = "Project Number" Then
                        'show error message must have data
                        Exit For
                    ElseIf item = "Quantity" Then
                        txtqty.Text = "0"
                    ElseIf item = "Unit Cost" Then
                        txtunitcost.Text = "0"
                    ElseIf item = "Unit Cost New" Then
                        txtunitcostnew.Text = "0"
                    ElseIf item = "Sample Cost" Then
                        txtsample.Text = "0"
                    ElseIf item = "Misc. Cost" Then
                        txttcost.Text = "0"
                    ElseIf item = "Vendor Number" Then
                        Exit For
                        'txtvendorno.Text = "0"  must have data
                    ElseIf item = "Tooling Cost" Then
                        txttoocost.Text = "0"
                    ElseIf item = "Sample Quantity" Then
                        txtsampleqty.Text = "0"
                    End If
                Next

                If txtvendorno.Text <> "" And projectNo <> 0 Then
                    QueryDetailResult = gnr.InsertProductDetail(projectNo, txtpartno.Text, DTPicker2, userid, dtTime, userid, dtTime1, txtctpno.Text, CInt(txtqty.Text),
                                    "", txtmfrno.Text, CInt(txtunitcost.Text), CInt(txtunitcostnew.Text), txtpo.Text, dtTime2, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                    cmbuser.SelectedValue, chknew, dtTime3, CInt(txtsample.Text), CInt(txttcost.Text), CInt(txtvendorno.Text), partstoshow, cmbminorcode.SelectedValue, CInt(txttoocost.Text), dtTime4,
                                    dtTime5, CInt(txtsampleqty.Text))
                Else
                    QueryDetailResult = -1
                    MessageBox.Show("The project number an d vendor number must have value.", "CTP System", MessageBoxButtons.OK)
                End If

                If QueryDetailResult < 0 Then
                    MessageBox.Show("Ann error ocurred inserting data in database.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub InsertProdWishList(userId As String, partNo As String)
        Dim exMessage As String = " "
        Try
            Dim rsAddPart As DialogResult = MessageBox.Show("Do you want to add this part to the Wish List?", "CTP System", MessageBoxButtons.YesNo)
            If rsAddPart = DialogResult.Yes Then
                Dim dsGetWLByPartNo = gnr.GetWLDataByPartNo(txtpartno.Text)

                If dsGetWLByPartNo IsNot Nothing Then
                    If dsGetWLByPartNo.Tables(0).Rows.Count > 0 Then
                        Dim rsPartExist As DialogResult = MessageBox.Show("This part # is already included in the wish list.", "CTP System", MessageBoxButtons.OK)
                    Else
                        Dim maxItemWL = gnr.getmax("PRDWL", "WHLCODE")
                        Dim rsInsWishListPart = gnr.InsertWishListProduct(maxItemWL, userId, txtpartno.Text)
                        If rsInsWishListPart < 0 Then
                            MessageBox.Show("Ann error ocurred inserting data in WishList.", "CTP System", MessageBoxButtons.OK)
                        End If
                    End If
                Else
                    Dim maxItemWL = gnr.getmax("PRDWL", "WHLCODE") + 1
                    Dim rsInsWishListPart = gnr.InsertWishListProduct(maxItemWL, userId, txtpartno.Text)
                    If rsInsWishListPart < 0 Then
                        MessageBox.Show("Ann error ocurred inserting data in WishList.", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub ProdDetailAndAllCommentHelper(strUser As String, flag As Integer)
        Dim exMessage As String = " "
        Try
            If flag = 0 Then
                Dim queryProdDetail = gnr.UpdateProductDetail(txtCode.Text, txtpartno.Text)
                If queryProdDetail <> 0 Then
                    MessageBox.Show("Ann error ocurred updating data in Prodcut Detail.", "CTP System", MessageBoxButtons.OK)
                End If
            End If

            Dim codComment = gnr.getmax("PRDCMH", "PRDCCO") + 1
            Dim queryProdComments = gnr.InsertProductComment(txtCode.Text, txtpartno.Text, codComment, userid)
            If queryProdComments <> 0 Then
                MessageBox.Show("Ann error ocurred inserting data in Product Comment Header database.", "CTP System", MessageBoxButtons.OK)
            End If
            Dim codDetComment = 1
            'Dim messcomm = "Person in charge changed assigned " & Trim(cmbuser.SelectedValue)
            Dim messcomm = strUser
            Dim queryProdCommentsDet = gnr.InsertProductCommentDetail(txtCode.Text, txtpartno.Text, codComment, codDetComment, messcomm)
            If queryProdCommentsDet <> 0 Then
                MessageBox.Show("Ann error ocurred inserting data in Product Comment Detail database.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub save()
        Dim exMessage As String = " "
        Try

            Dim insertYear As String = (Year(Now()) - 2000)
            'Dim test As String
            'insertYear = insertYear.Substring(1, 2)
            insertYear = CInt(insertYear)
            Dim insertMonth = Date.Today.Month
            Dim insertDay = Date.Today.Day
            Dim flagustatus As Integer
            partstoshow = displayPart()
            Dim QueryDetailResult As Integer = -1
            Dim statusquote As String

            If flagdeve = 1 Then 'new
                'assign logged user if not selected
                Dim validUser As String = If(cmbuser1.SelectedIndex < 1, userid, "")

                If Not String.IsNullOrEmpty(validUser) Then
                    cmbuser1.SelectedIndex = If(cmbuser1.FindString(Trim(UCase(validUser))) > 0, cmbuser1.FindString(Trim(validUser)), 0)
                End If

                Dim ProjectNo = gnr.getmax("PRDVLH", "PRHCOD") + 1

                'ALTERACION #1
                'Dim queryResult = gnr.InsertNewProject(ProjectNo, userid, DTPicker1, txtainfo.Text, txtname.Text, cmbprstatus, validUser)
                Dim queryResult = 1
                If queryResult < 0 Then
                    MessageBox.Show("Ann error ocurred inserting data in Product Header database.", "CTP System", MessageBoxButtons.OK)
                Else
                    txtCode.Text = ProjectNo
                    flagdeve = 0
                    'strwhere = CustomStrWhereResult()
                    'fillcell1LastOne(strwhere)

                    If flagnewpart = 1 Then
                        If Trim(txtpartno.Text) <> "" Then '?????
                            Dim Status2 As String = ""
                            'Status2 = If(gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()) <> "", gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()), "")
                            Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())
                            'Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo("1527554") 'burned
                            Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo(txtpartno.Text)
                            Dim strProjectNo = If(String.IsNullOrEmpty(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()), 0, CInt(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()))
                            Dim strProjectName = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(1).ToString())

                            If dsProjectNoResult.Tables(0).Rows.Count > 0 Then
                                If (ProjectNo = strProjectNo) Then
                                    Dim resultAlert As DialogResult = MessageBox.Show("This part no. already exists in this project. :" & ProjectNo & " - " & txtname.Text & "", "CTP System", MessageBoxButtons.OK)
                                Else
                                    Dim result As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & ProjectNo & " - " & strProjectName & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                    If result = DialogResult.No Then
                                        MessageBox.Show("No pressed")
                                    ElseIf result = DialogResult.Yes Then
                                        InsertProductDetails(ProjectNo, partstoshow)

                                        If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                            'send email
                                            gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                            'Dim result1 As Integer = gnr.sendEmail("")
                                        End If

                                        PoQotaFunction(Status2)

                                        If cmbuser.SelectedValue <> "N/A " Then
                                            ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                        End If

                                        'Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                        'If resultMsgUser = DialogResult.Yes Then
                                        '    'save files
                                        '    copyProjecFiles(strProjectNo)
                                        'End If
                                    End If
                                End If
                            Else
                                InsertProductDetails(ProjectNo, partstoshow)
                                If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                    'send email
                                    gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                    'Dim result As Integer = gnr.sendEmail("")
                                End If

                                PoQotaFunction(Status2)

                                If cmbuser.SelectedValue <> "N/A" Then
                                    ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                End If

                                'rectificado probar

                                'Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                'If resultMsgUser = DialogResult.Yes Then
                                '    'save files
                                '    copyProjecFiles(strProjectNo)
                                'End If
                            End If
                        End If
                    End If
                    SSTab1.TabPages(1).Text = "Project No." & Trim(txtCode.Text)
                    'SSTab1.tex = "Project No." & Trim(txtCode.Text)
                    txtsearchcode.Text = Trim(txtCode.Text)
                    'cmdsearchcode_Click(1)
                    'Dim resultDone As DialogResult = MessageBox.Show("The project is ready to add parts.", "CTP System", MessageBoxButtons.OK)
                    'flagdeve = 0
                    'flagnewpart = 0
                    requireValidation = 0
                    'SSTab1.SelectedIndex = 2
                End If
            Else 'modify
                Dim Status2 As String = ""
                If Not (String.IsNullOrEmpty(gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString()))) Then
                    Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())
                End If

                If Not String.IsNullOrEmpty(txtCode.Text) Then
                    cmdSave2.Enabled = False
                End If

                'Dim ProjectNoExists = gnr.GetDataByPRHCOD(Trim(txtCode.Text))
                'If ProjectNoExists IsNot Nothing Then
                Dim rsProdClosedParts As Integer
                If cmbprstatus.Text.IndexOf("F") = 1 Then

                    'End If
                    'If cmbprstatus.FindString("F") Then
                    Dim dsProdDet = gnr.GetProdDetByCodeAndExc(txtCode.Text)
                    If Not dsProdDet Is Nothing Then
                        If dsProdDet.Tables(0).Rows.Count <= 0 Then 'todos los parts# estan cerrados
                            rsProdClosedParts = gnr.UpdateProdClosedParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                                  Trim(cmbprstatus.Text.ToString().Substring(0, 1)), Trim(txtCode.Text))
                            If Not rsProdClosedParts.Equals(0) Then
                                MessageBox.Show("Ann error ocurred updating Closed Parts.", "CTP System", MessageBoxButtons.OK)
                            End If
                        Else 'hay # de parte abiertos
                            Dim resultOpenParts As DialogResult = MessageBox.Show("All Items must be closed if you want to finish the project.", "CTP System", MessageBoxButtons.OK)
                            Dim rsProdOpenParts = gnr.UpdateProdOpenParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                              Trim(txtCode.Text))
                            If Not rsProdOpenParts.Equals(0) Then
                                MessageBox.Show("Ann error ocurred updating Open Parts.", "CTP System", MessageBoxButtons.OK)
                            End If

                            'Dim resultError As DialogResult = MessageBox.Show("An error ocurred. Call to an administrator.", "CTP System", MessageBoxButtons.OK)

                        End If
                    End If
                Else
                    rsProdClosedParts = gnr.UpdateProdClosedParts(userid, DTPicker1.Value.ToString(), Trim(cmbuser1.SelectedValue.ToString()), Trim(txtainfo.Text), Trim(txtname.Text),
                                                                              Trim(cmbprstatus.Text.ToString().Substring(0, 1)), Trim(txtCode.Text))
                    If Not rsProdClosedParts.Equals(0) Then
                        MessageBox.Show("Ann error ocurred updating closed parts.", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
                'End If

                flagdeve = 0
                If flagnewpart = 1 Then
                    If Trim(txtpartno.Text) <> "" And Trim(txtvendorno.Text) <> "" Then

                        'validacion para custom insert in prdvlh
                        Dim ProjectNoDB = gnr.getmax("PRDVLH", "PRHCOD")
                        Dim ProjectNoCurrent = CInt(Trim(txtCode.Text))

                        Dim strPosition As Integer = cmbuser1.Text.IndexOf("N/A")
                        Dim validUser As String = If(strPosition = 0, userid, cmbuser1.Text)

                        If CInt(ProjectNoDB) + 1 = ProjectNoCurrent Then
                            Dim queryResult = gnr.InsertNewProject(ProjectNoCurrent, userid, DTPicker1, txtainfo.Text, txtname.Text, cmbprstatus, validUser)
                            If queryResult < 0 Then
                                MessageBox.Show("An error ocurred in the creation of a new project. Please check the input info!", "CTP System", MessageBoxButtons.OK)
                                Log.Error("An error ocurred in the creation of a new project")
                                Exit Sub
                            End If
                        End If

                        'Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo("1527554") 'burned
                        Dim dsProjectNoResult As DataSet = gnr.GetCodeAndNameByPartNo(txtpartno.Text)

                        Dim strProjectNo As String
                        Dim strProjectName As String
                        Dim strVendorNo As String

                        Dim ProjectNo = txtCode.Text
                        Dim codeTemp As String
                        Dim nameTemp As String
                        Dim validation As Integer = 0

                        If dsProjectNoResult IsNot Nothing Then
                            strProjectNo = If(String.IsNullOrEmpty(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()), 0, CInt(dsProjectNoResult.Tables(0).Rows(0).ItemArray(0).ToString()))
                            strProjectName = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(1).ToString())
                            strVendorNo = Trim(dsProjectNoResult.Tables(0).Rows(0).ItemArray(2).ToString())
                        Else
                            strProjectNo = 0
                            strProjectName = ""
                            strVendorNo = ""
                        End If

                        If dsProjectNoResult IsNot Nothing Then
                            If dsProjectNoResult.Tables(0).Rows.Count = 1 Then
                                If ((ProjectNo = strProjectNo) And (txtvendorno.Text = strVendorNo)) Then
                                    Dim resultAlert As DialogResult = MessageBox.Show("This part no. already exists in this project. :" & ProjectNo & " - " & txtname.Text & "", "CTP System", MessageBoxButtons.OK)
                                Else
                                    Dim result As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & strProjectNo & " - " & strProjectName & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                    If result = DialogResult.No Then
                                        Exit Sub
                                    ElseIf result = DialogResult.Yes Then

                                        InsertProductDetails(ProjectNo, partstoshow)
                                        If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                            'send email
                                            gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                            'Dim result1 As Integer = gnr.sendEmail("")
                                        End If

                                        InsertProdWishList(userid, txtpartno.Text)

                                        PoQotaFunction(Status2)

                                        If cmbuser.SelectedValue <> "N/A " Then
                                            ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                        End If

                                        'Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                        'If resultMsgUser = DialogResult.Yes Then
                                        '    'save files
                                        '    copyProjecFiles(strProjectNo)
                                        'End If

                                        MessageBox.Show("Reference Added Successfully.", "CTP System", MessageBoxButtons.OK)
                                    End If
                                End If
                            ElseIf dsProjectNoResult.Tables(0).Rows.Count > 1 Then
                                For Each ttt As DataRow In dsProjectNoResult.Tables(0).Rows
                                    If ((txtCode.Text = ttt.ItemArray(0).ToString()) And
                                            (txtvendorno.Text = ttt.ItemArray(2))) Then
                                        Dim result1 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & txtCode.Text & " - " & Trim(txtname.Text) & " with the vendor: " & Trim(txtvendorname.Text), "CTP System", MessageBoxButtons.OK)
                                        validation = 1
                                        Exit Sub
                                        'Exit For
                                    Else
                                        codeTemp = ttt.ItemArray(0).ToString()
                                        nameTemp = ttt.ItemArray(1).ToString()
                                    End If
                                Next
                                If (Not String.IsNullOrEmpty(codeTemp) And Not String.IsNullOrEmpty(nameTemp)) And validation = 0 Then
                                    Dim result2 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & codeTemp & " - " & Trim(nameTemp) & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                    If result2 = DialogResult.No Then
                                        Exit Sub
                                    ElseIf result2 = DialogResult.Yes Then
                                        InsertProductDetails(ProjectNo, partstoshow)
                                        If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                            'send email
                                            gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                            'Dim result As Integer = gnr.sendEmail("")
                                        End If

                                        InsertProdWishList(userid, txtpartno.Text)

                                        PoQotaFunction(Status2)

                                        If cmbuser.SelectedValue <> "N/A " Then
                                            ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                                        End If

                                        'Dim resultMsgUser As DialogResult = MessageBox.Show("Do you want to add the files in project no. : " & ProjectNo & " - " & strProjectName & "", "CTP System", MessageBoxButtons.YesNo)
                                        'If resultMsgUser = DialogResult.Yes Then
                                        '    'save files
                                        '    copyProjecFiles(strProjectNo)
                                        'End If

                                        MessageBox.Show("Reference Added Successfully.", "CTP System", MessageBoxButtons.OK)
                                    End If
                                End If
                            End If
                        Else
                            InsertProductDetails(ProjectNo, partstoshow)

                            If Trim(Status2) = "Technical Documentation" Or Trim(Status2) = "Analysis of Samples" Or Trim(Status2) = "Pending from Supplier" Then
                                'send email
                                gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                'Dim result As Integer = gnr.sendEmail("")
                            End If

                            PoQotaFunction(Status2)

                            If cmbuser.SelectedValue <> "N/A " Then
                                ProdDetailAndAllCommentHelper(cmbuser.SelectedValue, 0)
                            End If

                            MessageBox.Show("Reference Added Successfully.", "CTP System", MessageBoxButtons.OK)
                        End If
                    End If
                Else 'update
                    If Trim(txtpartno.Text) <> "" And Trim(txtvendorno.Text) <> "" Then
                        Dim dsGetProdDesc = gnr.GetDataByCodeAndPartNoProdDesc(txtCode.Text, txtpartno.Text)
                        If dsGetProdDesc.Tables(0).Rows.Count > 0 Then
                            If Trim(cmbuser.SelectedValue) <> Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal)) Then
                                Dim messcomm = "Person in charge changed from " & Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal)) & " to " & Trim(cmbuser.SelectedValue)
                                ProdDetailAndAllCommentHelper(messcomm, 0)
                            End If
                            If cmbstatus.SelectedValue = "CS" Or cmbstatus.SelectedValue = "CN" Then
                                InsertProdWishList(userid, txtpartno.Text)
                                'Dim rsAddPart As DialogResult = MessageBox.Show("Do you want to add this part to the Wish List?", "CTP System", MessageBoxButtons.YesNo)
                                'If rsAddPart = DialogResult.Yes Then
                                '    Dim dsGetWLByPartNo = gnr.GetWLDataByPartNo(txtpartno.Text)
                                '    If dsGetWLByPartNo.Tables(0).Rows.Count > 0 Then
                                '        Dim rsPartExist As DialogResult = MessageBox.Show("This part # is already included in the wish list.", "CTP System", MessageBoxButtons.OK)
                                '    Else
                                '        Dim maxItem = gnr.getmax("PRDWL", "PRWCOD")
                                '        Dim rsInsWishListPart = gnr.InsertWishListProduct(maxItem, userid, txtpartno.Text)
                                '        If rsInsWishListPart < 0 Then

                                '            'error message
                                '        End If
                                '    End If
                                'End If
                            End If
                            Dim status1 = ""
                            status1 = gnr.GetProjectStatusDescription(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDUSR").Ordinal))
                            Status2 = gnr.GetProjectStatusDescription(cmbstatus.SelectedValue.ToString())

                            flagustatus = 1

                            If Trim(cmbstatus.SelectedValue) <> dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) Then
                                If (Trim(Status2) = "Closed w/o negotiation") Or (Trim(Status2) = "Closed (Demand/cost/material)") Then
                                    Dim rsEnterComm As DialogResult = MessageBox.Show("Enter Comment.", "CTP System", MessageBoxButtons.OK)
                                    'gnr.seeaddprocomments = 5
                                    'frmproductsdevelopmentcomments.ShowDialog()
                                    cmdcomments_Click(Nothing, Nothing)
                                End If
                                If (Trim(Status2) = "Approved") Or (Trim(Status2) = "Approved with advice") Then
                                    Dim rsAssignVendor As DialogResult = MessageBox.Show("Do you want to change the assigned vendor?", "CTP System", MessageBoxButtons.YesNo)
                                    If rsAssignVendor = DialogResult.Yes Then
                                        Dim dsGetPartVendor = gnr.GetDataByPartVendor(txtpartno.Text)
                                        If dsGetPartVendor.Tables(0).Rows.Count > 0 Then
                                            gnr.changeVendor(txtpartno.Text, txtvendorno.Text, userid)
                                        Else
                                            Dim dsGetPartNoVendor = gnr.GetDataByPartNoVendor(txtpartno.Text)
                                            If dsGetPartNoVendor.Tables(0).Rows.Count > 0 Then
                                                Dim rsInsertNewInv = gnr.InsertNewInv("", txtpartno.Text, dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc1").Ordinal),
                                                                                      dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc2").Ordinal), "", txtunitcostnew.Text, "", "", txtvendorno.Text)
                                                If rsInsertNewInv <> 0 Then
                                                    MessageBox.Show("Ann error ocurred inserting data in Inventory.", "CTP System", MessageBoxButtons.OK)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                'paso de receiving of first production a rejected y esto implica cambio de vendor asignado
                                If (Trim(cmbstatus.SelectedValue) = "R ") And (dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) = "RP") Then
                                    Dim flagchangevendor = 1
                                    frmChangeVendor.ShowDialog()
                                End If
                                If Trim(Status2) = "Closed Successfully" Then

                                    gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                    Dim result As Integer = gnr.sendEmail("", txtpartno.Text)

                                    'toemails = prepareEmailsToSend(1)
                                    'Dim rsResult = gnr.sendEmail(toemails, txtpartno.Text)
                                    'If rsResult < 0 Then
                                    '    MessageBox.Show("Ann error ocurred sending emails.", "CTP System", MessageBoxButtons.OK)
                                    'End If

                                    Dim rsAssignVendor As DialogResult = MessageBox.Show("Do you want to change the assigned vendor?", "CTP System", MessageBoxButtons.YesNo)
                                    If rsAssignVendor = DialogResult.Yes Then
                                        Dim dsGetPartVendor = gnr.GetDataByPartVendor(txtpartno.Text)
                                        If dsGetPartVendor.Tables(0).Rows.Count > 0 Then
                                            gnr.changeVendor(txtpartno.Text, txtvendorno.Text, userid)
                                        Else
                                            Dim dsGetPartNoVendor = gnr.GetDataByPartNoVendor(txtpartno.Text)
                                            If dsGetPartNoVendor.Tables(0).Rows.Count > 0 Then
                                                Dim rsInsertNewInv = gnr.InsertNewInv("", txtpartno.Text, dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc1").Ordinal),
                                                                                      dsGetPartNoVendor.Tables(0).Rows(0).ItemArray(dsGetPartNoVendor.Tables(0).Columns("impc2").Ordinal), "", txtunitcostnew.Text, "", "", txtvendorno.Text)
                                                If rsInsertNewInv <> 0 Then
                                                    MessageBox.Show("Ann error ocurred inserting data in Inventory.", "CTP System", MessageBoxButtons.OK)
                                                End If
                                            End If
                                        End If
                                    End If

                                End If
                                If (Trim(Status2) = "Technical Documentation") Or (Trim(Status2) = "Analysis of Samples") Or (Trim(Status2) = "Pending from Supplier") Then
                                    'send email
                                    gnr.OpenOutlookMessage(txtname.Text, txtpartno.Text, Status2)
                                    'Dim result As Integer = gnr.sendEmail("")
                                End If
                                'remove condition to prevent the update the statues when analysis of sample yto others
                                'If (Trim(dsGetProdDesc.Tables(0).Rows(0).ItemArray(dsGetProdDesc.Tables(0).Columns("PRDSTS").Ordinal) = "AS") And (Trim(cmbstatus.SelectedValue) <> "AS")) Then
                                '    If (Trim(cmbstatus.SelectedValue) = "R") Or Trim(cmbstatus.SelectedValue) = "A" Or Trim(cmbstatus.SelectedValue) = "AA" Then
                                '        flagustatus = 1
                                '    Else
                                '        flagustatus = 0
                                '        Dim rsStatusGet As DialogResult = MessageBox.Show("Status can not be changed. You must change this item to Approved with advice, Approved or Rejected.", "CTP System", MessageBoxButtons.OK)
                                '    End If
                                'Else
                                '    flagustatus = 1
                                'End If
                                If flagustatus = 1 Then
                                    Dim messcomm = "Status changed from " & status1 & " to " & Status2
                                    ProdDetailAndAllCommentHelper(messcomm, flagustatus)

                                    PoQotaFunction(Status2)

                                End If
                            End If
                        End If

                        Dim chkValue As Integer = If(chknew.Checked, 1, 0)

                        If flagustatus = 1 Then
                            Dim rsUpdProdDet = gnr.UpdateProductDetail1(partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, DTPicker5.Value, "", txtvendorno.Text, chkValue,
                                                                        DTPicker4.Value, txtsample.Text, txttcost.Text, cmbuser.SelectedValue, DTPicker2.Value, userid,
                                                                        txtctpno.Text, txtsampleqty.Text, txtqty.Text, "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text,
                                                                        DTPicker3.Value, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text, txtCode.Text, txtpartno.Text)
                            If rsUpdProdDet <> 0 Then
                                MessageBox.Show("Ann error ocurred updating data in Product Detail database.", "CTP System", MessageBoxButtons.OK)
                            End If
                        Else
                            Dim rsUpdProdDet = gnr.UpdateProductDetail2(partstoshow, cmbminorcode.SelectedValue, txttoocost.Text, DTPicker5.Value, txtvendorno.Text, chkValue,
                                                                       DTPicker4.Value, txtsample.Text, txttcost.Text, cmbuser.SelectedValue, DTPicker2.Value, userid,
                                                                       txtctpno.Text, txtsampleqty.Text, txtqty.Text, "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text,
                                                                       DTPicker3.Value, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text, txtpartno.Text)
                            If rsUpdProdDet <> 0 Then
                                MessageBox.Show("Ann error ocurred updating data in Product Detail database.", "CTP System", MessageBoxButtons.OK)
                            End If
                        End If

                        Dim mpnopo As String = String.Empty
                        Dim spacepoqota As String = String.Empty
                        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                        mpnopo = Trim(UCase(txtmfrno.Text))
                        Dim maxValue = 0
                        Dim dsUpdatedData As Integer
                        statusquote = "D-" & Status2

                        Dim strCheckPoQoteIns = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo,
                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota, txtunitcostnew.Text, txtminqty.Text)
                        Dim newUnitCost = If(txtunitcostnew.Text <> "0", txtunitcostnew.Text, 0)

                        If String.IsNullOrEmpty(strCheckPoQoteIns) Then
                            dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, newUnitCost, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                txtvendorno.Text, txtpartno.Text)
                            If dsUpdatedData <> 0 Then
                                'show message error
                            End If
                        Else
                            Dim arrayCheck As New List(Of String)
                            arrayCheck = strCheckPoQoteIns.Split(",").ToList()
                            For Each item As String In arrayCheck
                                If item = "Sequencial" Then
                                    'show error message
                                    Exit For
                                ElseIf item = "Vendor Number" Then
                                    txtvendorno.Text = "0" 'ask for vendor??
                                ElseIf item = "Unit Cost New" Then
                                    txtunitcostnew.Text = "0"
                                ElseIf item = "Min Quantity" Then
                                    txtminqty.Text = "0"
                                End If
                            Next
                            dsUpdatedData = gnr.UpdatePoQoraRow(mpnopo, txtminqty.Text, newUnitCost, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                txtvendorno.Text, txtpartno.Text)

                            If dsUpdatedData <> 0 Then
                                'show message error
                            End If
                        End If
                        MessageBox.Show("Reference Updated Successfully.", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
                txtsearchcode.Text = Trim(txtCode.Text)
                'cmdsearchcode_Click(1)

                Dim dsGetProdDetByCodeAndExc = gnr.GetProdDetByCodeAndExc(txtCode.Text)
                If Not dsGetProdDetByCodeAndExc Is Nothing Then
                    If dsGetProdDetByCodeAndExc.Tables(0).Rows.Count = 0 Then 'todos los parts# estan cerrados
                        Dim dspMsg As DialogResult = MessageBox.Show("All parts for this project are closed. Do you want to finish the project?", "CTP System", MessageBoxButtons.YesNo)
                        If dspMsg = DialogResult.Yes Then
                            Dim rsUpdProdDevHeader = gnr.UpdateProductDevHeader(txtCode.Text)
                            If rsUpdProdDevHeader <> 0 Then
                                MessageBox.Show("Ann error ocurred updating data in Product Header datatable.", "CTP System", MessageBoxButtons.OK)
                            Else
                                If flagnewpart = 0 Then
                                    Dim dspUpdMess As DialogResult = MessageBox.Show("Project Updated Succesfully.", "CTP System", MessageBoxButtons.OK)
                                Else
                                    Dim dspCreatMess As DialogResult = MessageBox.Show("Project Closed Succesfully.", "CTP System", MessageBoxButtons.OK)
                                End If
                            End If
                        End If
                    End If
                End If
                requireValidation = 0
            End If

            If SSTab1.SelectedIndex = 2 Then
                If Trim(txtpartno.Text) <> "" Then
                    'fillcell2(txtCode.Text)
                    forceDbClick_Action(txtCode.Text, 3, True)

                    Dim dspNewPart As DialogResult = MessageBox.Show("Do you want to add other part to the project?", "CTP System", MessageBoxButtons.YesNo)
                    If dspNewPart = DialogResult.No Then
                        SSTab1.SelectedTab = TabPage2

                        'check if all references are closed to update the general status
                        Dim strResult = checkPendingReferences(txtCode.Text)
                        If strResult = "I" Then
                            cmbprstatus.SelectedIndex = 1
                        Else
                            cmbprstatus.SelectedIndex = 2
                        End If

                    Else
                        SSTab1.SelectedTab = TabPage3
                        flagnewpart = 1
                    End If
                    cleanFormValues("TabPage3", 0)
                    setVendorValues()
                End If
            Else
                fillcell1LastOne("")
            End If


        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdSave2_Click(sender As Object, e As EventArgs) Handles cmdSave2.Click
        Dim exMessage As String = Nothing
        Try
            Dim validationResult = mandatoryFields("save", SSTab1.SelectedIndex, 1)
            If validationResult.Equals(0) Then
                'Dim result As DialogResult = MessageBox.Show("Do you want to create a new project?", "CTP System", MessageBoxButtons.YesNo)
                'If result = DialogResult.No Then
                'MessageBox.Show("No pressed")
                'ElseIf result = DialogResult.Yes Then
                'MessageBox.Show("Yes pressed")
                save()
                If SSTab1.SelectedIndex = 1 Then
                    If flagnewpart = 1 Then
                        Dim result1 As DialogResult = MessageBox.Show("The project is ready to add parts. Please proceed to the project tab to add parts?", "CTP System", MessageBoxButtons.OK)
                        If result1 = DialogResult.OK Then
                            cmbuser.SelectedIndex = cmbuser1.SelectedIndex
                            SSTab1.SelectedTab = TabPage3
                        End If
                    End If
                End If
                'End If
            Else
                Dim resultSave As DialogResult = MessageBox.Show("Error in Data Validation. Mandatory fields must be filled!!", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdSave3_Click(sender As Object, e As EventArgs) Handles cmdSave3.Click
        Dim exMessage As String = Nothing
        Try
            Dim DtUseTime = New DateTimePicker()
            DtUseTime.Value = DateTime.Now
            Dim rsValidation = gnr.checkFields(txtCode.Text, txtpartno.Text, DTPicker2, userid, DtUseTime, userid, DtUseTime, txtctpno.Text, txtqty.Text,
                                                                    "", txtmfrno.Text, txtunitcost.Text, txtunitcostnew.Text, txtpo.Text, DtUseTime, cmbstatus.SelectedValue, txtBenefits.Text, txtcomm.Text,
                                                                    cmbuser.SelectedValue, chknew, DtUseTime, txtsample.Text, txttcost.Text, txtvendorno.Text, 0, cmbminorcode.SelectedValue, txttoocost.Text, DtUseTime,
                                                                    DateTime.Now.ToShortDateString(), txtsampleqty.Text)

            Dim validationResult = mandatoryFields("save", SSTab1.SelectedIndex, 1, rsValidation)
            If validationResult.Equals(0) Then
                Dim result As DialogResult = If(flagnewpart = 1, MessageBox.Show("If click yes the part will be added to the project. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo),
                                                    MessageBox.Show("If click yes will updated this part data. Do you want to proceed?", "CTP System", MessageBoxButtons.YesNo))
                If result = DialogResult.No Then
                    'MessageBox.Show("No pressed")
                ElseIf result = DialogResult.Yes Then
                    'MessageBox.Show("Yes pressed")
                    save()
                End If
            Else
                Dim resultSave As DialogResult = MessageBox.Show("Error in Data Validation. Mandatory fields must be filled!!", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click

        Dim resultNew As DialogResult = MessageBox.Show("Operation not allowed from this tab screen!!! ", "CTP System", MessageBoxButtons.OK)

    End Sub

    Private Function getDatetimeValue(strDate As String) As DateTime
        Dim exMessage As String = " "
        Try
            Dim CleanDateString As String = Regex.Replace(strDate, "/[^0-9a-zA-Z:]/g", "")
            Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
            Return dtChange
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Sub cmddelete_Click(sender As Object, e As EventArgs) Handles cmddelete.Click
        Dim exMessage As String = " "
        Dim rsValidationDev As Boolean = False
        Dim rsValidationPOq As Boolean = False
        Try
            If Trim(txtCode.Text) <> "" And Trim(txtpartno.Text) <> "" Then
                Dim resultAlert As DialogResult = MessageBox.Show("Are you sure, you want to delete this item??.", "CTP System", MessageBoxButtons.YesNo)
                If resultAlert = DialogResult.Yes Then
                    Dim dsResult = gnr.GetDataByCodeAndVendorAndPart(txtCode.Text, txtvendorno.Text, txtpartno.Text)
                    If Not dsResult Is Nothing Then
                        If dsResult.Tables(0).Rows.Count > 0 Then
                            Dim code = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                            Dim partNo = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDPTN").Ordinal).ToString()
                            Dim prdDat = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDDAT").Ordinal).ToString())
                            Dim crUser = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("CRUSER").Ordinal).ToString()
                            Dim crDate = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("CRDATE").Ordinal).ToString())
                            Dim moUser = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("MOUSER").Ordinal).ToString()
                            Dim moDate = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("MODATE").Ordinal).ToString())
                            Dim ctpNo = Trim(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDCTP").Ordinal).ToString())
                            Dim qty = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDQTY").Ordinal).ToString()
                            Dim mfr = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDMFR").Ordinal).ToString()
                            Dim mfrNo = Trim(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDMFR#").Ordinal).ToString())
                            Dim prdCos = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDCOS").Ordinal).ToString()
                            Dim prdCon = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDCON").Ordinal).ToString()
                            Dim poNo = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDPO#").Ordinal).ToString()
                            Dim poDate = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PODATE").Ordinal).ToString())
                            Dim status = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDSTS").Ordinal).ToString()
                            Dim benefits = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDBEN").Ordinal).ToString()
                            Dim info = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDINF").Ordinal).ToString()
                            Dim prdUsr = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDUSR").Ordinal).ToString()
                            Dim chkNew = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDNEW").Ordinal).ToString()
                            Dim prdEdd = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDEDD").Ordinal).ToString())
                            Dim prdSco = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDSCO").Ordinal).ToString()
                            Dim miscCost = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDTTC").Ordinal).ToString()
                            Dim vendorNo = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("VMVNUM").Ordinal).ToString()
                            Dim prdPts = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDPTS").Ordinal).ToString()
                            Dim prdMpc = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDMPC").Ordinal).ToString()
                            Dim toolCost = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDTCO").Ordinal).ToString()
                            Dim prdErd = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDERD").Ordinal).ToString())
                            Dim prdPda = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDPDA").Ordinal).ToString())
                            Dim prdSqty = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRDSQTY").Ordinal).ToString()
                            Dim prwLda = getDatetimeValue(dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRWLDA").Ordinal).ToString())
                            Dim prwLfl = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PRWLFL").Ordinal).ToString()
                            Dim partNoo = dsResult.Tables(0).Rows(0).ItemArray(dsResult.Tables(0).Columns("PARTNO").Ordinal).ToString()

                            Dim rsInsertion = gnr.InsertIntoLogProdDetail(userid, code, partNo, prdDat, crUser, crDate, moUser, moDate, ctpNo, qty, mfr, mfrNo, prdCos, prdCon, poNo, poDate, status,
                                                                            benefits, info, prdUsr, chkNew, prdEdd, prdSco, miscCost, vendorNo, prdPts, prdMpc, toolCost, prdErd, prdPda,
                                                                            prdSqty, prwLda, prwLfl, partNo)
                            If rsInsertion < 0 Then
                                Log.Error(exMessage)
                                MessageBox.Show("An error ocurred inserting data in Log Product Detail datatable.", "CTP System", MessageBoxButtons.OK)
                            End If
                        End If
                    End If

                    Dim prodDetDeletion = gnr.DeleteDataFromProdDet(txtCode.Text, txtpartno.Text)
                    If prodDetDeletion = 1 Then
                        Dim prodCommHeaderDeletion = gnr.DeleteDataFromProdCommHeader(txtCode.Text, txtpartno.Text)
                        If prodCommHeaderDeletion >= 0 Then
                            Dim prodCommDetDeletion = gnr.DeleteDataFromProdCommDet(txtCode.Text, txtpartno.Text)
                            If prodCommDetDeletion >= 0 Then
                                Log.Info("Deletion process succeed.")
                                rsValidationDev = True
                                ' MessageBox.Show("Deletion process succeed.", "CTP System", MessageBoxButtons.OK)
                            End If
                        End If
                    End If

                    Dim dsData = gnr.GetDataFromProdHeaderAndDetail2(txtCode.Text, txtpartno.Text, txtvendorno.Text)
                    If dsData Is Nothing Then
                        Dim rsPoqotaDeletion = gnr.DeleteDataFromPoQota(txtvendorno.Text, txtpartno.Text)
                        If rsPoqotaDeletion < 0 Then
                            MessageBox.Show("Ann error ocurred deleting data from POQOTA.", "CTP System", MessageBoxButtons.OK)
                        Else
                            rsValidationPOq = True
                        End If
                    Else
                        If dsData.Tables(0).Rows.Count = 0 Then
                            Dim rsPoqotaDeletion = gnr.DeleteDataFromPoQota(txtvendorno.Text, txtpartno.Text)
                            If rsPoqotaDeletion < 0 Then
                                MessageBox.Show("Ann error ocurred deleting data from POQOTA.", "CTP System", MessageBoxButtons.OK)
                            Else
                                rsValidationPOq = True
                            End If
                        End If
                    End If

                    If rsValidationPOq = True And rsValidationDev = True Then
                        Dim refAmount As Integer = GetAmountOfProjectReferences(txtCode.Text)
                        If refAmount = 0 Then
                            Dim rsDel = gnr.DeleteDataFromProdHead(txtCode.Text)
                            If rsDel = 1 Then
                                MessageBox.Show("The deletion process completed successfully.", "CTP System", MessageBoxButtons.OK)
                                txtCode.Text = String.Empty
                                cmdall_Click("cmdall2", Nothing)
                                SSTab1.SelectedIndex = 0
                            End If
                        Else
                            fillcell2(txtCode.Text)
                            MessageBox.Show("Record Deleted.", "CTP System", MessageBoxButtons.OK)
                            SSTab1.SelectedIndex = 1
                        End If
                    Else
                        MessageBox.Show("The deletion process does not complete successfully.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    Dim resultAlert1 As DialogResult = MessageBox.Show("Select Project and Part # to see files.", "CTP System", MessageBoxButtons.OK)
                End If

            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdpartno_Click(sender As Object, e As EventArgs) Handles cmdpartno.Click
        Dim exMessage As String = " "
        Try
            If (flagdeve = 0 And flagnewpart = 1) Or (flagdeve) = 1 Then
                If Trim(txtvendorno.Text) <> "" Then
                    Dim partno = InputBox("Enter Part No. :", "Select Part No.")
                    If Trim(partno) <> "" Then
                        cleanFormValues(SSTab1.SelectedTab.Name, 1)
                        SSTab1.TabPages(2).Text = "Part No." & Trim(partno)
                        Dim dsGetDataFromProdHeadAndDet = gnr.GetDataFromProdHeaderAndDetail(partno)
                        Dim codeTemp As String
                        Dim nameTemp As String
                        Dim validation As Integer = 0
                        If Not dsGetDataFromProdHeadAndDet Is Nothing Then
                            If dsGetDataFromProdHeadAndDet.Tables(0).Rows.Count = 1 Then
                                codeTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                nameTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                If txtCode.Text = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString() Then
                                    Dim result1 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & codeTemp & " - " & Trim(nameTemp), "CTP System", MessageBoxButtons.OK)
                                Else
                                    codeTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                    nameTemp = dsGetDataFromProdHeadAndDet.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                    Dim result2 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & codeTemp & " - " & Trim(nameTemp), "CTP System", MessageBoxButtons.OK)
                                End If
                            ElseIf dsGetDataFromProdHeadAndDet.Tables(0).Rows.Count > 1 Then
                                For Each ttt As DataRow In dsGetDataFromProdHeadAndDet.Tables(0).Rows
                                    If txtCode.Text = ttt.ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal.ToString()).ToString() Then
                                        Dim result1 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & txtCode.Text & " - " & Trim(txtname.Text), "CTP System", MessageBoxButtons.OK)
                                        validation = 1
                                        Exit Sub
                                        'Exit For
                                    Else
                                        codeTemp = ttt.ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRHCOD").Ordinal.ToString())
                                        nameTemp = ttt.ItemArray(dsGetDataFromProdHeadAndDet.Tables(0).Columns("PRNAME").Ordinal.ToString())
                                    End If
                                Next
                                If (Not String.IsNullOrEmpty(codeTemp) And Not String.IsNullOrEmpty(nameTemp)) And validation = 0 Then
                                    Dim result2 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & codeTemp & " - " & Trim(nameTemp), "CTP System", MessageBoxButtons.OK)
                                End If

                            End If
                        Else
                            MessageBox.Show("This part is not present in other projects. We are looking for it in the inventory.", "CTP System", MessageBoxButtons.OK)
                        End If

                        Dim dsGetDataFromDualInv = gnr.GetDataFromDualInventory(partno)
                        If Not dsGetDataFromDualInv Is Nothing Then
                            If dsGetDataFromDualInv.Tables(0).Rows.Count > 0 Then
                                txtpartno.Text = partno
                                txtpartdescription.Text = Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMDSC").Ordinal).ToString())

                                If cmbminorcode.FindStringExact(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString())) Then
                                    cmbminorcode.SelectedIndex = cmbminorcode.FindString(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC2").Ordinal).ToString()))
                                End If

                                If cmbmajorcode.FindStringExact(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC1").Ordinal).ToString())) Then
                                    cmbmajorcode.SelectedIndex = cmbmajorcode.FindString(Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("IMPC1").Ordinal).ToString()))
                                End If

                                If Trim(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()) <> "" Then
                                    Dim dsGetVendorQuey = gnr.GetVendorQuey(dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString())
                                    If Not dsGetVendorQuey Is Nothing Then
                                        If dsGetVendorQuey.Tables(0).Rows.Count > 0 Then
                                            txtvendornoa.Text = dsGetDataFromDualInv.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInv.Tables(0).Columns("DVPRMG").Ordinal).ToString()
                                            txtvendornamea.Text = Trim(dsGetVendorQuey.Tables(0).Rows(0).ItemArray(dsGetVendorQuey.Tables(0).Columns("VMNAME").Ordinal).ToString())
                                        Else
                                            txtvendornoa.Text = ""
                                            txtvendornamea.Text = ""
                                        End If
                                    End If
                                Else
                                    txtvendornoa.Text = ""
                                    txtvendornamea.Text = ""
                                End If

                                Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partno)
                                If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                                    txtctpno.Text = dsGetCTPPartRef
                                    txtmfrno.Text = dsGetCTPPartRef
                                Else
                                    txtctpno.Text = ""
                                    txtmfrno.Text = ""
                                End If

                                If Trim(txtvendornoa.Text) <> "" Then
                                    Dim dsGetAssignedVendor = gnr.GetAssignedVendor(txtvendornoa.Text, partno)
                                    If Not dsGetAssignedVendor Is Nothing Then
                                        If dsGetAssignedVendor.Tables(0).Rows.Count > 0 Then
                                            txtunitcost.Text = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQPRC").Ordinal).ToString()
                                            txtminqty.Text = 0
                                            'txtminqty.Text = dsGetAssignedVendor.Tables(0).Rows(0).ItemArray(dsGetAssignedVendor.Tables(0).Columns("PQMIN").Ordinal).ToString()
                                        Else
                                            txtunitcost.Text = 0
                                            txtminqty.Text = 0
                                        End If
                                    End If
                                Else
                                    txtunitcost.Text = 0
                                    txtminqty.Text = 0
                                End If

                                searchpart()
                                'Call searchpart
                                'txtctpno.SetFocus
                                chknew.Checked = False
                                chknew.Enabled = False
                            Else
                                chknew.Enabled = True
                            End If
                        Else
                            Dim dsGetDataFromDualInventory1 = gnr.GetDataByPartNoVendor(partno)
                            If Not dsGetDataFromDualInventory1 Is Nothing Then
                                If dsGetDataFromDualInventory1.Tables(0).Rows.Count > 0 Then
                                    txtpartno.Text = partno
                                    txtpartdescription.Text = Trim(dsGetDataFromDualInventory1.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInventory1.Tables(0).Columns("IMDSC").Ordinal).ToString())

                                    If cmbminorcode.FindStringExact(Trim(dsGetDataFromDualInventory1.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInventory1.Tables(0).Columns("IMPC2").Ordinal).ToString())) Then
                                        cmbminorcode.SelectedIndex = cmbminorcode.FindString(Trim(dsGetDataFromDualInventory1.Tables(0).Rows(0).ItemArray(dsGetDataFromDualInventory1.Tables(0).Columns("IMPC2").Ordinal).ToString()))
                                    End If

                                    txtvendornoa.Text = ""
                                    txtvendornamea.Text = ""

                                    Dim dsGetCTPPartRef = gnr.GetCTPPartRef(partno)
                                    If Not String.IsNullOrEmpty(dsGetCTPPartRef) Then
                                        txtctpno.Text = dsGetCTPPartRef
                                        txtmfrno.Text = dsGetCTPPartRef
                                    Else
                                        txtctpno.Text = ""
                                        txtmfrno.Text = ""
                                    End If
                                    searchpart()
                                    'txtctpno.SetFocus
                                    chknew.Checked = False
                                    chknew.Enabled = False
                                Else
                                    chknew.Enabled = True
                                    Dim result3 As DialogResult = MessageBox.Show("Part No. not found.", "CTP System", MessageBoxButtons.OK)
                                End If
                            Else
                                chknew.Enabled = True
                                MessageBox.Show("This part does not exists in our inventary. Please add to the inventary before trying to use.", "CTP System", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                        End If

                        'test purpose
                        'Dim testPartNo = "5257106"
                        Dim dsGetPartInWishList = gnr.GetPartInWishList(partno)
                        'Dim dsGetPartInWishList = gnr.GetPartInWishList(testPartNo)
                        If Not dsGetPartInWishList Is Nothing Then
                            If dsGetPartInWishList.Tables(0).Rows.Count > 0 Then
                                chknew.Enabled = False
                                chknew.Checked = False
                                Dim wlcode = dsGetPartInWishList.Tables(0).Rows(0).ItemArray(dsGetPartInWishList.Tables(0).Columns("WHLCODE").Ordinal).ToString()

                                Dim dsGetDataByVendorAndPartNoProdDesc = gnr.GetDataByVendorAndPartNoProdDesc(txtvendorno.Text, partno)
                                'Dim dsGetDataByVendorAndPartNoProdDesc = gnr.GetDataByVendorAndPartNoProdDesc(tetsVendorNo, testPartNo)
                                If Not dsGetDataByVendorAndPartNoProdDesc Is Nothing Then
                                    If dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows.Count > 0 Then
                                        'Dim dsGetDataByCodAndPartProdAndComm =
                                        'gnr.GetDataByCodAndPartProdAndComm(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Columns("PRHCOD").Ordinal).ToString(), partno)
                                        'test purposes
                                        Dim dsGetDataByCodAndPartProdAndComm =
                                            gnr.GetDataByCodAndPartProdAndComm(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNoProdDesc.Tables(0).Columns("PRHCOD").Ordinal).ToString(), txtpartno.Text)
                                        If Not dsGetDataByCodAndPartProdAndComm Is Nothing Then
                                            If dsGetDataByCodAndPartProdAndComm.Tables(0).Rows.Count > 0 Then
                                                Dim result4 As DialogResult = MessageBox.Show("This part# : " & Trim(UCase(partno)) & " has been quoted with this vendor# : " & Trim(txtvendorno.Text) & " before. Do you want to continue?", "CTP System", MessageBoxButtons.YesNo)
                                                If result4 = DialogResult.No Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                Dim dsGetDataFromProdHeaderAndDetail = gnr.GetDataFromProdHeaderAndDetail(partno)
                                Dim dtpDate = New DateTimePicker()
                                Dim dtpDate1 = New DateTimePicker()
                                Dim dt = DateTime.Now

                                Dim iDate As String = "1900-01-01"
                                Dim oDate As DateTime = DateTime.Parse(iDate)
                                dtpDate.Value = dt
                                dtpDate1.Value = oDate
                                Dim code As String
                                Dim name As String

                                If Not dsGetDataFromProdHeaderAndDetail Is Nothing Then
                                    If dsGetDataFromProdHeaderAndDetail.Tables(0).Rows.Count > 0 Then
                                        If Trim(txtCode.Text) = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString() Then
                                            code = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                            name = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                            Dim result5 As DialogResult = MessageBox.Show("This part no. already exists in this project. : " & code & "-" & name & " ", "CTP System", MessageBoxButtons.OK)
                                        Else
                                            code = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRHCOD").Ordinal).ToString()
                                            name = dsGetDataFromProdHeaderAndDetail.Tables(0).Rows(0).ItemArray(dsGetDataFromProdHeaderAndDetail.Tables(0).Columns("PRNAME").Ordinal).ToString()
                                            Dim result6 As DialogResult = MessageBox.Show("This part no. already exists in project no. : " & code & "-" & name & ". Do you want to create it?.", "CTP System", MessageBoxButtons.YesNo)
                                            If result6 = DialogResult.Yes Then

                                                InsertProductDetails(txtCode.Text, partstoshow)
                                                'Dim rsInsertProductDetailv2 = gnr.InsertProductDetailv2(txtCode.Text, txtpartno.Text, dtpDate, userid, dtpDate, userid, dtpDate, txtctpno.Text,
                                                '                                                        0, "", "", txtunitcost.Text, 0, "", dtpDate1, "E", "", "", userid, chknew, dtpDate1, 0, 0, txtvendorno.Text,
                                                '                                                        "", cmbminorcode.SelectedValue, 0, dtpDate1, dtpDate1, DTPicker2, 1)
                                                'If rsInsertProductDetailv2 <> 0 Then
                                                '    'error message
                                                'End If

                                                Dim statusquote = "D-Entered"
                                                Dim mpnopo1 As String
                                                Dim spacepoqota1 As String = String.Empty
                                                Dim strQueryAdd1 As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                                                Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text) 'aqui llegue full

                                                If dsPoQota IsNot Nothing Then
                                                    If dsPoQota.Tables(0).Rows.Count > 0 Then
                                                        mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                        Dim maxValue1 = 0
                                                        Dim dsUpdatedData1 As Integer

                                                        Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                        If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                            dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                                txtvendorno.Text, txtpartno.Text)
                                                            If dsUpdatedData1 <> 0 Then
                                                                'show message error
                                                            End If
                                                        Else
                                                            Dim arrayCheck As New List(Of String)
                                                            arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                            For Each item As String In arrayCheck
                                                                If item = "Sequencial" Then
                                                                    'show error message
                                                                    Exit For
                                                                ElseIf item = "Vendor Number" Then
                                                                    txtvendorno.Text = "0" 'ask for vendor??
                                                                ElseIf item = "Unit Cost New" Then
                                                                    txtunitcostnew.Text = "0"
                                                                End If
                                                            Next
                                                            dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                                txtvendorno.Text, txtpartno.Text)

                                                            If dsUpdatedData1 <> 0 Then
                                                                'show message error
                                                            End If
                                                        End If
                                                    Else
                                                        'warning message
                                                    End If
                                                Else
                                                    Dim maxValue1 = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd1)
                                                    If Not String.IsNullOrEmpty(maxValue1) Then
                                                        maxValue1 += 1
                                                    Else
                                                        maxValue1 = 1 'preguntar duda
                                                    End If
                                                    spacepoqota1 = "                               DEV"
                                                    mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                    Dim ResultQuery As String = String.Empty

                                                    Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                            DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                    If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                        ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                           DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                        If ResultQuery <> 0 Then
                                                            'show message error
                                                        End If
                                                    Else
                                                        Dim arrayCheck As New List(Of String)
                                                        arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                        For Each item As String In arrayCheck
                                                            If item = "Sequencial" Then
                                                                'show error message
                                                                Exit For
                                                            ElseIf item = "Vendor Number" Then
                                                                txtvendorno.Text = "0"
                                                            End If
                                                        Next

                                                        ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                           DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                        If ResultQuery <> 0 Then
                                                            'show message error
                                                        End If
                                                    End If

                                                End If

                                                Dim rsDeletion = gnr.DeleteDataByWSCod(wlcode)
                                                If rsDeletion = 1 Then
                                                    'deletion ok
                                                End If

                                                'Call gotonew()
                                                'Call cmdall_Click()
                                                Dim result7 As DialogResult = MessageBox.Show("Part # added to Project : " & Trim(txtCode.Text), "CTP System", MessageBoxButtons.OK)

                                            End If
                                        End If
                                    Else
                                        'Dim rsInsertProductDetailv2 = gnr.InsertProductDetailv2(txtCode.Text, txtpartno.Text, dtpDate, userid, dtpDate, userid, dtpDate, txtctpno.Text,
                                        '                                                                0, "", "", txtunitcost.Text, 0, "", dtpDate1, "E", "", "", userid, chknew, dtpDate1, 0, 0, txtvendorno.Text,
                                        '                                                                "", cmbminorcode.SelectedValue, 0, dtpDate1, dtpDate1, DTPicker2, 1)
                                        'If rsInsertProductDetailv2 <> 0 Then
                                        '    'error message
                                        'End If

                                        InsertProductDetails(txtCode.Text, partstoshow)

                                        Dim statusquote = "D-Entered"
                                        Dim mpnopo1 As String
                                        Dim spacepoqota1 As String = String.Empty
                                        Dim strQueryAdd1 As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                                        Dim dsPoQota = gnr.GetPOQotaData(txtvendorno.Text, txtpartno.Text)

                                        If dsPoQota IsNot Nothing Then
                                            If dsPoQota.Tables(0).Rows.Count > 0 Then
                                                mpnopo1 = Trim(UCase(txtmfrno.Text))
                                                Dim maxValue1 = 0
                                                Dim dsUpdatedData1 As Integer

                                                Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                    DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                                If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                    dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                        txtvendorno.Text, txtpartno.Text)
                                                    If dsUpdatedData1 <> 0 Then
                                                        'show message error
                                                    End If
                                                Else
                                                    Dim arrayCheck As New List(Of String)
                                                    arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                    For Each item As String In arrayCheck
                                                        If item = "Sequencial" Then
                                                            'show error message
                                                            Exit For
                                                        ElseIf item = "Vendor Number" Then
                                                            txtvendorno.Text = "0" 'ask for vendor??
                                                        End If
                                                    Next
                                                    dsUpdatedData1 = gnr.UpdatePoQoraRow1(mpnopo1, statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(),
                                                                        txtvendorno.Text, txtpartno.Text)

                                                    If dsUpdatedData1 <> 0 Then
                                                        'show message error
                                                    End If
                                                End If
                                            Else
                                                'warning message
                                            End If
                                        Else
                                            Dim maxValue1 = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd1)
                                            If Not String.IsNullOrEmpty(maxValue1) Then
                                                maxValue1 += 1
                                            Else
                                                maxValue1 = 1 'preguntar duda
                                            End If
                                            spacepoqota1 = "                               DEV"
                                            mpnopo1 = Trim(UCase(txtmfrno.Text))
                                            Dim ResultQuery As String = String.Empty

                                            Dim strCheckPoQoteIns1 = gnr.checkfieldsPoQote(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                    DateTime.Now.Day.ToString(), statusquote, spacepoqota1, txtunitcostnew.Text, txtminqty.Text)
                                            If String.IsNullOrEmpty(strCheckPoQoteIns1) Then
                                                ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                   DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                If ResultQuery <> 0 Then
                                                    'show message error
                                                End If
                                            Else
                                                Dim arrayCheck As New List(Of String)
                                                arrayCheck = strCheckPoQoteIns1.Split(",").ToList()
                                                For Each item As String In arrayCheck
                                                    If item = "Sequencial" Then
                                                        'show error message
                                                        Exit For
                                                    ElseIf item = "Vendor Number" Then
                                                        txtvendorno.Text = "0"
                                                    End If
                                                Next

                                                ResultQuery = gnr.InsertNewPOQota1(txtpartno.Text, txtvendorno.Text, maxValue1, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), mpnopo1,
                                                                                   DateTime.Now.Day.ToString(), statusquote, spacepoqota1)
                                                If ResultQuery <> 0 Then
                                                    'show message error
                                                End If
                                            End If
                                        End If

                                        Dim rsDeletion = gnr.DeleteDataByWSCod(wlcode)
                                        If rsDeletion = 1 Then
                                            'deletion ok
                                        End If
                                        'Call gotonew()
                                        'Call cmdall_Click()
                                        Dim result7 As DialogResult = MessageBox.Show("Part # added to Project : " & Trim(txtCode.Text), "CTP System", MessageBoxButtons.OK)

                                    End If
                                End If
                            End If
                        End If
                        txtsample.Text = "0"
                        txttcost.Text = "0"
                        txttoocost.Text = "0"

                        txtsampleqty.Text = "0"
                        txtBenefits.Text = " "
                        'txtainfo.Text = " "
                        txtqty.Text = "0"
                        txtunitcostnew.Text = "0"

                        'new item or new supplier
                        chknew.Checked = False
                        chkSupplier.Checked = False
                        chknew.Checked = If(itemCategory(txtpartno.Text, txtvendorno.Text) = 2, True, False)
                        chkSupplier.Checked = If(chknew.Checked, False, True)
                        'If chknew.Checked Then
                        '    chkSupplier.Checked = Not chknew.Checked
                        'End If
                    Else
                        cmdfiles.Visible = False
                        cmdcomments.Visible = False
                        cmdseecomments.Visible = False
                        cmdseefiles.Visible = False
                    End If
                Else
                    Dim result As DialogResult = MessageBox.Show("Enter Vendor.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                Dim result1 As DialogResult = MessageBox.Show("Part No. cannot be changed when is already created.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdvendor_Click(sender As Object, e As EventArgs) Handles cmdvendor.Click
        Dim exMessage As String = " "
        Dim partstoshow As String = displayPart()
        Try
            Dim oldvendorno = Trim(txtvendorno.Text)
            Dim vendorno = InputBox("Enter Vendor No. :", "Change Vendor")
            '889912

            If Not IsNumeric(vendorno) Then
                Dim result As DialogResult = MessageBox.Show("Enter just numbers.", "CTP System", MessageBoxButtons.OK)
            Else
                Dim dsGetVendorByVendorNo = gnr.GetVendorByVendorNo(vendorno)
                If Not dsGetVendorByVendorNo Is Nothing Then
                    If (dsGetVendorByVendorNo.Tables(0).Rows.Count > 0) Then
                        If gnr.customIsVendorAccepted(vendorno) Then
                            txtvendorno.Text = vendorno
                            txtvendorname.Text = dsGetVendorByVendorNo.Tables(0).Rows(0).ItemArray(dsGetVendorByVendorNo.Tables(0).Columns("VMNAME").Ordinal).ToString()
                            partstoshow = ""

                            optCTP.Checked = True
                            optVENDOR.Checked = False
                            optboth.Checked = False
                            partstoshow = "1"
                            Dim strQueryAdd As String = "WHERE PQVND = " & Trim(vendorno) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
                            If flagnewpart = 0 And Trim(txtpartno.Text) <> "" Then
                                Dim dsGetDataByVendorAndPartNo = gnr.GetDataByVendorAndPartNoDst(oldvendorno, txtpartno.Text)
                                If Not dsGetDataByVendorAndPartNo Is Nothing Then
                                    If dsGetDataByVendorAndPartNo.Tables(0).Rows.Count > 0 Then
                                        Dim rsUpdatePoQotaByVendorAndPart = gnr.UpdatePoQotaByVendorAndPart(vendorno, oldvendorno, txtpartno.Text,
                                                                            dsGetDataByVendorAndPartNo.Tables(0).Rows(0).ItemArray(dsGetDataByVendorAndPartNo.Tables(0).Columns("PQSEQ").Ordinal).ToString())
                                        If rsUpdatePoQotaByVendorAndPart <> 0 Then
                                            MessageBox.Show("Ann error ocurred updating POQOTA datatable.", "CTP System", MessageBoxButtons.OK)
                                        End If
                                    Else
                                        Dim maxValue = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                                        If Not String.IsNullOrEmpty(maxValue) Then
                                            maxValue += 1
                                        Else
                                            Dim spacepoqota = "                               DEV"
                                            Dim rsInsertNewPOQota = gnr.InsertNewPOQotaLess(txtpartno.Text, vendorno, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), "", DateTime.Now.Day.ToString(), "", spacepoqota, 0)
                                            If rsInsertNewPOQota <> 0 Then
                                                MessageBox.Show("Ann error ocurred inserting data in POQOTA.", "CTP System", MessageBoxButtons.OK)
                                            End If
                                            maxValue = 1 'preguntar duda
                                        End If

                                    End If
                                    Dim rsUpdProdDetVend = gnr.UpdateProdDetailVendor(partstoshow, vendorno, txtCode.Text, txtpartno.Text)
                                    If rsUpdProdDetVend <> 0 Then
                                        MessageBox.Show("Ann error ocurred updating data in Product Detail Datatable.", "CTP System", MessageBoxButtons.OK)
                                    End If
                                    fillcell2(txtCode.Text)
                                End If
                                Dim result2 As DialogResult = MessageBox.Show("Vendor Changed.", "CTP System", MessageBoxButtons.OK)
                            End If
                        Else
                            txtvendorno.Text = ""
                            txtvendorname.Text = ""
                            MessageBox.Show("Not valid vendor.", "CTP System", MessageBoxButtons.OK)
                        End If
                    Else
                        Dim result3 As DialogResult = MessageBox.Show("Vendor not found.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    Dim result4 As DialogResult = MessageBox.Show("Vendor not found.", "CTP System", MessageBoxButtons.OK)
                End If
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdgenerate_Click(sender As Object, e As EventArgs) Handles cmdgenerate.Click
        Dim exMessage As String = " "
        Try
            If Trim(txtpartno.Text) <> "" Then
                If Trim(txtctpno.Text) <> "" Then
                    Dim result As DialogResult = MessageBox.Show("CTP # has been already generated.", "CTP System", MessageBoxButtons.OK)
                Else
                    Dim strCTPExist = gnr.GetCTPPartRef(txtpartno.Text)
                    If Not String.IsNullOrEmpty(strCTPExist) Then
                        txtctpno.Text = strCTPExist
                        txtmfrno.Text = strCTPExist
                    Else
                        'Dim PartNo = Trim(UCase(txtpartno.Text)).Substring(0, 19) & "                   "
                        Dim PartNo = Trim(UCase(txtpartno.Text)) & "                   "
                        PartNo = PartNo.Substring(0, Math.Min(PartNo.Length, 19))
                        Dim ctppartno = "                   "
                        Dim flagctp = "9"
                        Dim dsctpValue = gnr.CallForCtpNumber(PartNo, ctppartno, flagctp)
                        If Not dsctpValue Is Nothing Then
                            If dsctpValue.Tables(0).Rows.Count > 0 Then
                                txtctpno.Text = Trim(UCase(dsctpValue.Tables(0).Rows(0).ItemArray(1).ToString()))
                                txtmfrno.Text = Trim(UCase(dsctpValue.Tables(0).Rows(0).ItemArray(1).ToString()))
                            Else
                                txtctpno.Text = ""
                                txtmfrno.Text = ""
                            End If
                        End If
                    End If
                End If
            Else
                Dim result1 As DialogResult = MessageBox.Show("Select Part No.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        cmdSearch_Click()
    End Sub

    Private Sub txts__KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) _
        Handles Me.KeyPress, txtsearchpart.KeyPress, txtsearchctp.KeyPress, txtsearchcode.KeyPress, txtsearch1.KeyPress, txtsearch.KeyPress, txtMfrNoSearch.KeyPress, txtJiratasksearch.KeyPress, cmbstatus1.KeyPress, cmbPrpech.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            cmdhidden_Click(sender, Nothing)
        End If
    End Sub

    Private Sub txts1__KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) _
        Handles Me.KeyPress, txtPartNoMore.KeyPress, txtCtpNoMore.KeyPress, txtMfrNoMore.KeyPress, cmbuser2.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            cmdhidden1_Click(sender, Nothing)
        End If
    End Sub

    Private Sub cmdhidden_Click(sender As Object, e As EventArgs)
        Dim exMessage As String = " "
        Dim controlSender As Object = Nothing
        Dim isText As Boolean = True
        Try
            Dim cnt = DirectCast(sender, System.Windows.Forms.Control)
            Dim sender_type = cnt.GetType().ToString()
            If sender_type.Equals("System.Windows.Forms.TextBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.TextBox)
            ElseIf sender_type.Equals("System.Windows.Forms.ComboBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.ComboBox)
                isText = False
            Else
                controlSender = Nothing
            End If
            Dim ctrl_name = If(controlSender IsNot Nothing, controlSender.Name, "")
            If Not String.IsNullOrEmpty(ctrl_name) Then

                'Dim button_name = If(isText, ctrl_name.Replace("txt", "cmd"), ctrl_name.Replace("cmb", "cmd"))
                'Dim button_method = button_name & "_click"
                Dim button_method = "cmdall_Click"
                Dim selection(2) As Object
                selection(0) = ctrl_name
                selection(1) = isText
                DataGridView1.Visible = True
                dgvProjectDetails.Visible = True
                CallByName(Me, button_method, CallType.Method, selection(0), selection(1))
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdall_Click(Optional control_name As Object = Nothing, Optional is_text As Object = Nothing)
        Dim exMessage As String = " "
        Dim button_name As String = Nothing
        Dim button_method As String = Nothing
        Try
            If is_text IsNot Nothing Then
                button_name = If(is_text, control_name.Replace("txt", "cmd"), control_name.Replace("cmb", "cmd"))
            Else
                button_name = control_name
            End If
            button_method = button_name & "_click"
            CallByName(Me, button_method, CallType.Method, Nothing)

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdall_Click_1(sender As Object, e As EventArgs) Handles cmdall.Click
        Dim b As System.Windows.Forms.Button = DirectCast(sender, System.Windows.Forms.Button)
        cmdall_Click(b.Name & 2, Nothing)
    End Sub


#Region "First Tab Searching Methods"

    Private Sub cmdClearFilters_Click(sender As Object, e As EventArgs) Handles cmdClearFilters.Click
        Dim exMessage As String = " "
        Try
            onlyClearSearchesComplex()
            cleanDataSources()
            LikeSession.focussedControl = Nothing
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdClearFilters1_Click(sender As Object, e As EventArgs) Handles cmdClearFilters.Click
        Dim exMessage As String = " "
        Try
            onlyClearSearchesComplexTab2()
            'cleanDataSources()
            fillcell2(txtCode.Text)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        cmdSearch_Click()
    End Sub

    Public Sub cmdall2_Click(Optional ByVal flag As Integer = 0)
        Dim exMessage As String = " "
        Dim genericObj As Object = Nothing
        'Dim tt As Windows.Forms.TextBox
        'Dim cm As Windows.Forms.ComboBox
        Dim controlSender As Object = Nothing
        Dim isText As Boolean = True
        'tt = txtsearch
        Dim lstQueries = New List(Of String)()
        Try

            'Dim cnt = FindFocussedControl(Me)
            Dim cnt = LikeSession.focussedControl
            If cnt IsNot Nothing Then
                Dim sender_type = cnt.GetType().ToString()
                If sender_type.Equals("System.Windows.Forms.TextBox") Then
                    controlSender = DirectCast(cnt, System.Windows.Forms.TextBox)
                ElseIf sender_type.Equals("System.Windows.Forms.ComboBox") Then
                    controlSender = DirectCast(cnt, System.Windows.Forms.ComboBox)
                    isText = False
                Else
                    controlSender = Nothing
                End If
                Dim ctrl_name = If(controlSender IsNot Nothing, controlSender.Name, "")

                If Not String.IsNullOrEmpty(ctrl_name) Then

                    'Dim button_name = If(isText, ctrl_name.Replace("txt", "cmd"), ctrl_name.Replace("cmb", "cmd"))
                    'Dim button_method = button_name & "_click"
                    Dim button_method = "cmdall_Click"
                    Dim selection(2) As Object
                    selection(0) = ctrl_name
                    selection(1) = isText
                    DataGridView1.Visible = True
                    dgvProjectDetails.Visible = True
                    CallByName(Me, button_method, CallType.Method, selection(0), selection(1))
                End If
            Else
                cmdSearchAll_Click()
            End If




            'If isText Then
            '    genericObj = DirectCast(cnt, System.Windows.Forms.TextBox)
            'Else
            '    genericObj = DirectCast(cnt, System.Windows.Forms.ComboBox)
            'End If

            'If Trim(tt.Text) <> "" Then
            '            If flagallow = 1 Then
            '                strwhere = buildMixedQuery(lstQueries, genericObj.Name, 0, True, True)
            '            Else
            '                If gnr.checkPurcByUser(userid) <> -1 Then
            '                    Dim purcValue = gnr.checkPurcByUser(userid)
            '                    strwhere = "WHERE (PRPECH = '" & userid & "' OR A1.PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "'))"
            '                    strToUnion = "UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue
            '                    strToUnionTab2 = "UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
            'Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue
            '                Else
            '                    strwhere = "WHERE (PRPECH = '" & userid & "' OR A1.PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "'))"
            '                End If
            '            End If
            'Else
            '    MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            'End If

            'lstQueries.Add(strwhere)
            'lstQueries.Add(strToUnion)
            'lstQueries.Add(strToUnionTab2)
            'buildMixedQuery(lstQueries, genericObj.Name, 0, True)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdSearchAll_Click(Optional ByVal flag As Integer = 0)
        Dim exMessage As String = " "
        Dim lstQueries = New List(Of String)()
        Dim sql As String = Nothing
        Try

            If flagallow = 1 Then
                strwhere = ""
            Else
                If gnr.checkPurcByUser(userid) <> -1 Then
                    Dim purcValue = gnr.checkPurcByUser(userid)
                    strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') "
                    strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & ""
                    strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " "
                Else
                    strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') "
                End If
                'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
            End If

            lstQueries.Add(strwhere)
            lstQueries.Add(strToUnion)
            lstQueries.Add(strToUnionTab2)

            sql = lstQueries(0)
            Dim IQ1 = lstQueries(1)
            Dim IQ2 = lstQueries(2)

            'sql += outputQuery
            'IQ1 += outputQuery
            'IQ2 += outputQuery
            'lstQueries(1) = IQ1

            'Dim txtTemp = initialQuery(2)
            lstQueries(2) = sql + IQ2

            sql += lstQueries(1)
            If flag = 1 Then
                fillcell1(sql, 0)
            Else
                fillcelldetail(sql, 0, lstQueries(2))
            End If

            'buildMixedQuery(lstQueries, Nothing, 0)
            'fillcell1(strwhere, flag)
            'cleanSearchTextBoxes(tt.Name)

            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'Call gotoerror("frmproductsdevelopment", "cmdsearch_click", Err.Number, Err.Description, Err.Source)
    End Sub

    Public Sub cmdSearch_Click(Optional ByVal flag As Integer = 0)
        Dim exMessage As String = " "
        Dim tt As Windows.Forms.TextBox
        tt = txtsearch
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                    End If
                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(txtsearch.Text)), "'", "") & "%'"
                End If

                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
                'fillcell1(strwhere, flag)
                'cleanSearchTextBoxes(tt.Name)
            Else
                MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'Call gotoerror("frmproductsdevelopment", "cmdsearch_click", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Sub cmdsearch1_Click(sender As Object, e As EventArgs) Handles cmdsearch1.Click
        cmdsearch1_Click()
    End Sub

    Public Sub cmdsearch1_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.TextBox
        Dim tt1 As Windows.Forms.ComboBox
        tt = txtsearch1
        tt1 = cmbstatus1
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        Dim lstQueries = New List(Of String)()
        Try

            If Not String.IsNullOrEmpty(tt.Text) Then
                If flagallow = 1 Then
                    strwhere = " WHERE A2.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        'strwhere = "WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND PRDVLD.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND A2.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND A2.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND A2.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND A2.VMVNUM = " & Trim(UCase(tt.Text)) & ""
                    End If
                End If
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
            Else
                MessageBox.Show("You must type a vendor number to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdsearchpart_Click(sender As Object, e As EventArgs) Handles cmdsearchpart.Click
        cmdSearchPart_Click()
    End Sub

    Public Sub cmdSearchPart_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.TextBox
        Dim tt1 As Windows.Forms.TextBox
        tt = txtsearchpart
        tt1 = txtsearchcode
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        Dim lstQueries = New List(Of String)()
        Try
            If Not String.IsNullOrEmpty(tt.Text) Then

                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(tt.Text)) & "' "
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(tt.Text)) & "' "
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(tt.Text)) & "' "
                    End If

                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDPTN)) = '" & Trim(UCase(txtsearchpart.Text)) & "' "
                End If

#Region "Previous"

                'If Not String.IsNullOrEmpty(tt1.Text) Then
                '    'project has value and part has value
                '    ds = fillcelldetailOther(strwhere)
                '    If ds IsNot Nothing Then
                '        If ds.Tables(0).Rows.Count > 0 Then
                '            Dim code As String = tt1.Text
                '            ds1 = gnr.GetDataByPRHCOD(code)

                '            Dim partNo As String = tt.Text

                '            txtCode.Text = Trim(ds1.Tables(0).Rows(0).ItemArray(0).ToString())
                '            txtname.Text = Trim(ds1.Tables(0).Rows(0).ItemArray(3).ToString()) ' format date
                '            TabPage2.Text = "Project: " + txtname.Text

                '            Dim CleanDateString As String = Regex.Replace(ds1.Tables(0).Rows(0).ItemArray(1).ToString(), "/[^0-9a-zA-Z:]/g", "")
                '            'Dim dtChange As DateTime = DateTime.ParseExact(CleanDateString, "MM/dd/yyyy HH:mm:ss tt", CultureInfo.InvariantCulture)
                '            Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                '            DTPicker1.Value = dtChange.ToShortDateString()

                '            If cmbuser1.FindStringExact(Trim(ds1.Tables(0).Rows(0).ItemArray(9).ToString())) Then
                '                cmbuser1.SelectedIndex = cmbuser1.FindString(Trim(ds1.Tables(0).Rows(0).ItemArray(9).ToString()))
                '            End If
                '            If cmbuser1.SelectedIndex = -1 Then
                '                cmbuser1.SelectedIndex = cmbuser1.Items.Count - 1
                '            End If
                '            If Trim(ds1.Tables(0).Rows(0).ItemArray(4).ToString()) = "I" Then
                '                cmbprstatus.SelectedIndex = 1
                '            ElseIf Trim(ds1.Tables(0).Rows(0).ItemArray(4).ToString()) = "F" Then
                '                cmbprstatus.SelectedIndex = 2
                '            Else
                '                cmbprstatus.SelectedIndex = 2
                '            End If

                '            fillcell2(code)

                '            fillTab3(code, partNo)

                '            SSTab1.SelectedIndex = 2
                '        Else
                '            MessageBox.Show("There is not search results with this criteria.", "CTP System", MessageBoxButtons.OK)
                '        End If
                '    Else
                '        MessageBox.Show("There is not search results with this criteria.", "CTP System", MessageBoxButtons.OK)
                '    End If
                '    'cleanSearchTextBoxes(tt.Name)
                'Else
                'fillcelldetail(strwhere)
                'cleanSearchTextBoxes(tt.Name)
                'End If

#End Region

                'only the part has value
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)

            Else
                'the part has no value
                MessageBox.Show("You must type a part number to find.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdsearchctp_Click(sender As Object, e As EventArgs) Handles cmdsearchctp.Click
        cmdsearchctp_Click()
    End Sub

    Public Sub cmdsearchctp_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.TextBox
        tt = txtsearchctp
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(tt.Text)) & "' "
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(tt.Text)) & "' "
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(tt.Text)) & "' "
                    End If

                    'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDCTP)) = '" & Trim(UCase(txtsearchctp.Text)) & "' "
                End If
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
            Else
                MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdJiratasksearch_Click(sender As Object, e As EventArgs) Handles cmdJiratasksearch.Click
        cmdJiratasksearch_Click()
    End Sub

    Public Sub cmdJiratasksearch_Click()
        Dim exMessage As String = " "
        Dim tt As Windows.Forms.TextBox
        tt = txtJiratasksearch
        Dim ds As New DataSet
        Dim ds1 As New DataSet
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRDJIRA)) = '" & Trim(UCase(tt.Text)) & "' "
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDJIRA)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRDJIRA)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRDJIRA)) = '" & Trim(UCase(tt.Text)) & "' "
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDJIRA)) = '" & Trim(UCase(tt.Text)) & "' "
                    End If
                End If

                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)

#Region "Previous"

                'ds = fillcelldetailOther(strwhere)
                'If ds IsNot Nothing Then
                '    If ds.Tables(0).Rows.Count > 0 Then
                '        Dim code As String = ds.Tables(0).Rows(0).ItemArray(0).ToString()
                '        ds1 = gnr.GetDataByPRHCOD(code)

                '        Dim partNo As String = ds.Tables(0).Rows(0).ItemArray(1).ToString()

                '        txtCode.Text = Trim(ds1.Tables(0).Rows(0).ItemArray(0).ToString())
                '        txtname.Text = Trim(ds1.Tables(0).Rows(0).ItemArray(3).ToString()) ' format date
                '        TabPage2.Text = "Project: " + txtname.Text

                '        Dim CleanDateString As String = Regex.Replace(ds1.Tables(0).Rows(0).ItemArray(1).ToString(), "/[^0-9a-zA-Z:]/g", "")
                '        'Dim dtChange As DateTime = DateTime.ParseExact(CleanDateString, "MM/dd/yyyy HH:mm:ss tt", CultureInfo.InvariantCulture)
                '        Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                '        DTPicker1.Value = dtChange.ToShortDateString()

                '        If cmbuser1.FindStringExact(Trim(ds1.Tables(0).Rows(0).ItemArray(9).ToString())) Then
                '            cmbuser1.SelectedIndex = cmbuser1.FindString(Trim(ds1.Tables(0).Rows(0).ItemArray(9).ToString()))
                '        End If
                '        If cmbuser1.SelectedIndex = -1 Then
                '            cmbuser1.SelectedIndex = cmbuser1.Items.Count - 1
                '        End If
                '        If Trim(ds1.Tables(0).Rows(0).ItemArray(4).ToString()) = "I" Then
                '            cmbprstatus.SelectedIndex = 1
                '        ElseIf Trim(ds1.Tables(0).Rows(0).ItemArray(4).ToString()) = "F" Then
                '            cmbprstatus.SelectedIndex = 2
                '        Else
                '            cmbprstatus.SelectedIndex = 2
                '        End If

                '        fillcell2(code)

                '        fillTab3(code, partNo)

                '        SSTab1.SelectedIndex = 2

                '        cleanSearchTextBoxes(tt.Name)
                '    Else
                '        MessageBox.Show("There is no matches to your searching criteria.", "CTP System", MessageBoxButtons.OK)
                '    End If
                'Else
                '    MessageBox.Show("There is no matches to your searching criteria.", "CTP System", MessageBoxButtons.OK)
                'End If

#End Region

            Else
                MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdMfrNoSearch_Click(sender As Object, e As EventArgs) Handles cmdMfrNoSearch.Click
        cmdMfrNoSearch_Click()
    End Sub

    Public Sub cmdMfrNoSearch_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"esta buen
        Dim tt As Windows.Forms.TextBox
        tt = txtMfrNoSearch
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then

                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRDMFR#)) = '" & Trim(UCase(tt.Text)) & "' "
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDMFR#)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND  TRIM(UCASE(PRDMFR#)) = '" & Trim(UCase(tt.Text)) & "' "
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRDMFR#)) = '" & Trim(UCase(tt.Text)) & "' "
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDMFR#)) = '" & Trim(UCase(tt.Text)) & "' "
                    End If
                End If
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
            Else
                MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdsearchcode_Click(sender As Object, e As EventArgs) Handles cmdsearchcode.Click
        Dim exMessage As String = " "
        Try
            'BackgroundWorker1.RunWorkerAsync()
            'Loading.Show()
            'Loading.BringToFront()

            cmdsearchcode_Click()

            'Dim bgWorker = CType(sender, BackgroundWorker)
            'For index = 0 To 10
            '    bgWorker.ReportProgress(index)
            '    Thread.Sleep(1000)
            'Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

    End Sub

    Public Sub cmdsearchcode_Click(Optional ByVal flag As Integer = 0, Optional ByVal flag1 As Boolean = Nothing)

        LikeSession.currentAction = "cmdsearchcode_Click"

        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.TextBox
        tt = txtsearchcode
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = " WHERE A1.PRHCOD = " & Trim(UCase(tt.Text))
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A1.PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND A1.PRHCOD = " & Trim(UCase(tt.Text))
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND A1.PRHCOD = " & Trim(UCase(tt.Text)) & ""
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND A1.PRHCOD = " & Trim(UCase(tt.Text)) & ""
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A1.PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "')) AND A1.PRHCOD = " & Trim(UCase(tt.Text))
                    End If
                End If

                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)

                If flag1 Then
                    buildMixedQuery(lstQueries, tt.Name, flag, False, False, True, True)
                Else
                    buildMixedQuery(lstQueries, tt.Name, flag, False, False, False, True)
                End If
                'buildMixedQuery(lstQueries, tt.Name, flag)
                'fillcell1(strwhere, flag)
                'cleanSearchTextBoxes(tt.Name)
            Else
                MessageBox.Show("You must type a search criteria to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdstatus1_Click(sender As Object, e As EventArgs) Handles cmdstatus1.Click
        Dim b As System.Windows.Forms.Button = DirectCast(sender, System.Windows.Forms.Button)
        cmdstatus1_Click()
    End Sub

    Public Sub cmdstatus1_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.ComboBox
        Dim tt1 As Windows.Forms.TextBox
        tt = cmbstatus1
        tt1 = txtsearch1
        Dim lstQueries = New List(Of String)()
        Try

            If Not String.IsNullOrEmpty(tt.SelectedValue) Then
                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRDSTS)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDSTS)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRDSTS)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRDSTS)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRDSTS)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                    End If
                End If
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
            Else
                MessageBox.Show("You must select a status value to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdPrpech_Click(sender As Object, e As EventArgs)
        cmdPrpech_Click()
    End Sub

    Public Sub cmdPrpech_Click()
        Dim exMessage As String = " "
        'userid = "LREDONDO"
        Dim tt As Windows.Forms.ComboBox
        tt = cmbPrpech
        Dim lstQueries = New List(Of String)()
        Try
            If Trim(tt.Text) <> "" Then
                If flagallow = 1 Then
                    strwhere = " WHERE TRIM(UCASE(PRPECH)) = '" & Trim(UCase(tt.SelectedValue)) & "' "
                Else
                    If gnr.checkPurcByUser(userid) <> -1 Then
                        Dim purcValue = gnr.checkPurcByUser(userid)
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR A2.PRDUSR = '" & userid & "') AND TRIM(UCASE(PRPECH)) = '" & Trim(UCase(tt.SelectedValue)) & "' "
                        strToUnion = " UNION SELECT DISTINCT (A1.prhcod),prname,prdate,prpech,prstat FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & " AND TRIM(UCASE(PRPECH)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                        strToUnionTab2 = " UNION SELECT DISTINCT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(A2.VMVNUM) as VMVNUM,
Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLH A1 INNER JOIN PRDVLD A2 ON A1.PRHCOD = A2.PRHCOD INNER JOIN VNMAS A3 ON A2.VMVNUM = A3.VMVNUM WHERE A3.VMABB# = " & purcValue & "  AND TRIM(UCASE(PRPECH)) = '" & Trim(UCase(tt.SelectedValue)) & "'"
                    Else
                        strwhere = " WHERE (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRPECH)) = '" & Trim(UCase(tt.SelectedValue)) & "' "
                    End If
                End If
                lstQueries.Add(strwhere)
                lstQueries.Add(strToUnion)
                lstQueries.Add(strToUnionTab2)
                buildMixedQuery(lstQueries, tt.Name, 0, False, False, False, True)
            Else
                MessageBox.Show("You must select a person in charge value to get results.", "CTP System", MessageBoxButtons.OK)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub


#End Region

#Region "Second Tab Search Filters"

    Public Sub cmdMfrNoMore_Click(sender As Object, e As EventArgs) Handles cmdMfrNoMore.Click
        cmdMfrNoMore_Click()
    End Sub

    Public Sub cmdMfrNoMore_Click()
        Dim strAddMfrSentence As String = ""
        strAddMfrSentence = " AND PRDMFR# = " ' & txtMfrNoMore.Text & '" 
        Dim exMessage As String = " "
        Dim Qry As New DataTable
        'Dim myButton As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
        'Dim myName As String = myButton.Name
        'cleanMoreBtns(myName)
        Try
            If Not String.IsNullOrEmpty(txtMfrNoMore.Text) Then

                'dgvProjectDetails.DataSource = Nothing
                'dgvProjectDetails.Refresh()

                'dgvProjectDetails.DataSource = LikeSession.dsDgvProjectDetails.Tables(0)

                Dim dt As New DataTable
                Dim ds As New DataSet

                dt = (DirectCast(dgvProjectDetails.DataSource, DataTable))
                'dt = If(LikeSession.dsDgvProjectDetails IsNot Nothing, LikeSession.dsDgvProjectDetails.Tables(0), Nothing)

                If dt IsNot Nothing Then
                    Dim Qry1 = dt.AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("PRDMFR#"))) = Trim(UCase(txtMfrNoMore.Text)))

                    If Qry1.Count > 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()
                    ElseIf Qry1.Count > 0 And Qry1.Count = 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()

                        fillTab3(txtCode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                        SSTab1.SelectedIndex = 2
                    Else
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        MessageBox.Show("There is not search matches for this criteria.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is an error loading data.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                fillcell2(txtCode.Text)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdPartNoMore_Click(sender As Object, e As EventArgs) Handles cmdPartNoMore.Click
        cmdPartNoMore_Click()
    End Sub

    Public Sub cmdPartNoMore_Click()
        Dim exMessage As String = " "
        Dim Qry As New DataTable
        Dim dsProjectDetails = LikeSession.dsDgvProjectDetails

        'Dim myButton As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
        'Dim myName As String = myButton.Name
        'cleanMoreBtns(myName)

        Try
            If Not String.IsNullOrEmpty(txtPartNoMore.Text) Then

                'dgvProjectDetails.DataSource = Nothing
                'dgvProjectDetails.Refresh()

                'dgvProjectDetails.DataSource = LikeSession.dsDgvProjectDetails.Tables(0)
                'fillcell2(txtCode.Text)

                Dim dt As New DataTable
                Dim ds As New DataSet
                'dt = (DirectCast(dgvProjectDetails.DataSource, DataTable))
                dt = If(LikeSession.dsDgvProjectDetails IsNot Nothing, LikeSession.dsDgvProjectDetails.Tables(0), Nothing)

                If dt IsNot Nothing Then
                    Dim Qry1 = dt.AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("PRDPTN"))) = Trim(UCase(txtPartNoMore.Text)))

                    If Qry1.Count > 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()
                    ElseIf Qry1.Count > 0 And Qry1.Count = 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()

                        fillTab3(txtCode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                        SSTab1.SelectedIndex = 2
                    Else
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        MessageBox.Show("There is not search matches for this criteria.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is an error loading data.", "CTP System", MessageBoxButtons.OK)
                    fillcell2(txtCode.Text)
                End If
            Else
                fillcell2(txtCode.Text)
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdCtpNoMore_Click(sender As Object, e As EventArgs) Handles cmdCtpNoMore.Click
        cmdCtpNoMore_Click()
    End Sub

    Public Sub cmdCtpNoMore_Click()
        'Dim strAddCtpSentence As String = ""
        'strAddCtpSentence = " AND PRDCTP = " ' & txtCtpNoMore.Text & '" 
        Dim Qry As New DataTable
        Dim exMessage As String = " "

        'Dim myButton As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
        'Dim myName As String = myButton.Name
        'cleanMoreBtns(myName)
        Try
            If Not String.IsNullOrEmpty(txtCtpNoMore.Text) Then

                'dgvProjectDetails.DataSource = Nothing
                'dgvProjectDetails.Refresh()

                'dgvProjectDetails.DataSource = LikeSession.dsDgvProjectDetails.Tables(0)

                Dim dt As New DataTable
                Dim ds As New DataSet

                dt = (DirectCast(dgvProjectDetails.DataSource, DataTable))
                'dt = If(LikeSession.dsDgvProjectDetails IsNot Nothing, LikeSession.dsDgvProjectDetails.Tables(0), Nothing)

                If dt IsNot Nothing Then
                    Dim Qry1 = dt.AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("PRDCTP"))) = Trim(UCase(txtCtpNoMore.Text)))

                    If Qry1.Count > 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()
                    ElseIf Qry1.Count > 0 And Qry1.Count = 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()

                        fillTab3(txtCode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                        SSTab1.SelectedIndex = 2
                    Else
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        MessageBox.Show("There is not search matches for this criteria.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is an error loading data.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                fillcell2(txtCode.Text)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdPePechMore_Click(sender As Object, e As EventArgs) Handles cmdUser2.Click
        cmduser2_click()
        'cmdPePechMore_Click()
    End Sub

    Public Sub cmduser2_click()
        'Dim strAddCtpSentence As String = ""
        'strAddCtpSentence = " AND PRDCTP = " ' & txtCtpNoMore.Text & '" 

        Dim exMessage As String = " "
        Dim Qry As New DataTable

        'Dim myButton As System.Windows.Forms.Button = CType(sender, System.Windows.Forms.Button)
        'Dim myName As String = myButton.Name
        'cleanMoreBtns(myName)
        Try
            If Not String.IsNullOrEmpty(cmbuser2.Text) And cmbuser2.SelectedIndex <> 0 Then

                'dgvProjectDetails.DataSource = Nothing
                'dgvProjectDetails.Refresh()

                'toPaginateDs(dgvProjectDetails, ds)
                'dgvProjectDetails.DataSource = LikeSession.dsDgvProjectDetails.Tables(0)

                Dim dt As New DataTable
                Dim ds As New DataSet
                dt = (DirectCast(dgvProjectDetails.DataSource, DataTable))


                'dt = If(LikeSession.dsDgvProjectDetails IsNot Nothing, LikeSession.dsDgvProjectDetails.Tables(0), Nothing)

                If dt IsNot Nothing Then
                    Dim Qry1 = dt.AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)("PRDUSR"))) = Trim(UCase(cmbuser2.SelectedValue)))

                    If Qry1.Count > 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()
                    ElseIf Qry1.Count > 0 And Qry1.Count = 1 Then
                        Qry = Qry1.CopyToDataTable
                        ds.Tables.Add(Qry)
                        toPaginateDs(dgvProjectDetails, ds)
                        'dgvProjectDetails.DataSource = Qry
                        'dgvProjectDetails.Refresh()

                        fillTab3(txtCode.Text, dgvProjectDetails.Rows(0).Cells(1).Value.ToString())
                        SSTab1.SelectedIndex = 2
                    Else
                        dgvProjectDetails.DataSource = Nothing
                        dgvProjectDetails.Refresh()
                        MessageBox.Show("There is not search matches for this criteria.", "CTP System", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("There is an error loading data.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                fillcell2(txtCode.Text)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdhidden1_Click(sender As Object, e As EventArgs)
        Dim exMessage As String = " "
        Dim controlSender As Object = Nothing
        Dim isText As Boolean = True
        Try
            Dim cnt = DirectCast(sender, System.Windows.Forms.Control)
            Dim sender_type = cnt.GetType().ToString()
            If sender_type.Equals("System.Windows.Forms.TextBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.TextBox)
            ElseIf sender_type.Equals("System.Windows.Forms.ComboBox") Then
                controlSender = DirectCast(sender, System.Windows.Forms.ComboBox)
                isText = False
            Else
                controlSender = Nothing
            End If
            Dim ctrl_name = If(controlSender IsNot Nothing, controlSender.Name, "")
            If Not String.IsNullOrEmpty(ctrl_name) Then

                'Dim button_name = If(isText, ctrl_name.Replace("txt", "cmd"), ctrl_name.Replace("cmb", "cmd"))
                'Dim button_method = button_name & "_click"
                Dim button_method = "cmdall1_Click_1"
                Dim selection(2) As Object
                selection(0) = ctrl_name
                selection(1) = isText
                CallByName(Me, button_method, CallType.Method, selection(0), selection(1))
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdall12_Click(Optional ByVal flag As Integer = 0)
        Dim Qry As New DataTable
        Dim exMessage As String = " "
        Try
            buildMixedQueryTab2()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdall1_Click_1(Optional control_name As Object = Nothing, Optional is_text As Object = Nothing)
        Dim exMessage As String = " "
        Try
            Dim button_name = If(is_text, control_name.Replace("txt", "cmd"), control_name.Replace("cmb", "cmd"))
            Dim button_method = button_name & "_click"
            CallByName(Me, button_method, CallType.Method, Nothing)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdAll1_Click(sender As Object, e As EventArgs) Handles cmdAll1.Click
        Dim b As System.Windows.Forms.Button = DirectCast(sender, System.Windows.Forms.Button)
        cmdall1_Click_1(b.Name & 2, Nothing)
    End Sub


#End Region

    Private Function buildMixedQueryTab2()
        Dim myTableLayout As TableLayoutPanel
        myTableLayout = Me.TableLayoutPanel15
        Dim sql As String = ""
        Dim hasVal As New List(Of Object)
        Dim selectedObj As Object = Nothing
        Dim exMessage As String = " "
        Try
            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    If tt.Text <> Nothing Then
                        hasVal.Add(tt)
                    End If
                    'If tt.Name <> valueSelectd Then
                    '    tt.Text = ""
                    'End If
                End If
            Next

            bs1.DataSource = Nothing
            sql += buildSearchQuerySintax(hasVal, 2)
            fillcell2(txtCode.Text, sql)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Function

    Private Function buildMixedQuery(initialQuery As List(Of String), selectedField As String, flag As Integer, Optional ByVal btnSelect As Boolean = Nothing,
                                     Optional ByVal flagBtn As Boolean = Nothing, Optional ByVal flagUpdate As Boolean = Nothing, Optional ByVal sessionFlag As Boolean = False) As String
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel1
            Dim sql As String = initialQuery(0)
            Dim hasVal As New List(Of Object)
            Dim selectedObj As Object = Nothing
            Dim outputQuery As String = Nothing

            If flagUpdate = False Then

                If Not btnSelect Then
                    For Each tt In myTableLayout.Controls
                        If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                            If tt.Text <> Nothing And tt.Name <> selectedField Then
                                hasVal.Add(tt)
                            ElseIf tt.Name = selectedField Then
                                selectedObj = tt
                            End If
                            'If tt.Name <> valueSelectd Then
                            '    tt.Text = ""
                            'End If
                        End If
                    Next
                Else
                    If Not flagBtn Then
                        For Each tt In myTableLayout.Controls
                            If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                                If tt.Text <> Nothing Then
                                    hasVal.Add(tt)
                                ElseIf tt.Name = selectedField Then
                                    selectedObj = tt
                                End If
                                'If tt.Name <> valueSelectd Then
                                '    tt.Text = ""
                                'End If
                            End If
                        Next
                    Else
                        For Each tt In myTableLayout.Controls
                            If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                                If tt.Name = selectedField Then
                                    hasVal.Add(tt)
                                    selectedObj = tt
                                End If
                                'If tt.Name <> valueSelectd Then
                                '    tt.Text = ""
                                'End If
                            End If
                        Next
                    End If
                End If

                LikeSession.searchControls = hasVal
                bs.DataSource = Nothing
                bs1.DataSource = Nothing
                outputQuery = buildSearchQuerySintax(hasVal, 1)

            End If

            If btnSelect And flagBtn Then
                Return outputQuery
            Else
                Dim IQ1 = initialQuery(1)
                Dim IQ2 = initialQuery(2)

                sql += outputQuery
                IQ1 += outputQuery
                IQ2 += outputQuery
                initialQuery(1) = IQ1

                'Dim txtTemp = initialQuery(2)
                initialQuery(2) = sql + IQ2

                sql += initialQuery(1)
                If flag = 1 Then
                    fillcell1(sql, 0)
                ElseIf flag = 2 Then
                    Dim code = If(txtCode.Text IsNot Nothing, txtCode.Text.Trim(), txtsearchcode.Text.Trim())
                    'forceDbClick_Action(code, 2)
                    fillcell2(code, initialQuery(1), Nothing, sessionFlag)
                Else
                    fillcelldetail(sql, 0, initialQuery(2), sessionFlag)
                End If
                hasVal.Add(selectedObj)
                cleanSearchTextBoxesComplex(hasVal, False)
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Sub forceDbClick_Action(code As String, Optional ByVal customFunct As Integer = 0, Optional ByVal isFromUpdate As Boolean = Nothing)
        Dim exMessage As String = " "
        Try
            Dim ctrl_name = txtsearchcode.Name
            Dim isText = True

            Dim button_name = If(isText, ctrl_name.Replace("txt", "cmd"), ctrl_name.Replace("cmb", "cmd"))
            Dim button_method = button_name & "_click"

            txtsearchcode.Text = code

            If customFunct <> 0 Then
                CallByName(Me, button_method, CallType.Method, customFunct, isFromUpdate)
            Else
                CallByName(Me, button_method, CallType.Method)
            End If


        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub onlyClearSearchesComplex()
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel1
            Dim hasVal As New List(Of Object)
            Dim selectedObj As Object = Nothing

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    If tt.Text <> Nothing Then
                        hasVal.Add(tt)
                    End If
                End If
            Next
            'hasVal.Add(selectedObj)
            cleanSearchTextBoxesComplex(hasVal, True)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub onlyClearSearchesComplexTab2()
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel15
            Dim hasVal As New List(Of Object)
            Dim selectedObj As Object = Nothing

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    If tt.Text <> Nothing Then
                        hasVal.Add(tt)
                    End If
                End If
            Next
            'hasVal.Add(selectedObj)
            cleanSearchTextBoxesComplexTab2(hasVal, True)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function getReferenceDocuments(code As String, partNo As String) As Dictionary(Of String, String)
        Dim exMessage As String = " "
        Try
            Dim dictionary As New Dictionary(Of String, String)
            gnr.FolderPath = gnr.UrlPDevelopmentMethod
            gnr.folderpathproject = gnr.FolderPath & Trim(code) & "\"
            Dim hasDocs As Boolean = False

            'If code IsNot Nothing Then
            If partNo IsNot Nothing Then
                'For Each item As DataRow In ds.Tables(0).Rows
                gnr.folderpathvendor = gnr.folderpathproject & Trim(partNo.ToString()) & "\"
                If Directory.Exists(gnr.folderpathvendor) Then
                    If Not IsDirectoryEmpty(gnr.folderpathvendor) Then
                        Dim lstPdfInside = Directory.GetFiles(gnr.folderpathvendor, "*.pdf", SearchOption.AllDirectories)
                        If lstPdfInside IsNot Nothing Then
                            If lstPdfInside.Count > 0 Then
                                hasDocs = True
                            End If
                        End If
                        dictionary.Add(partNo.ToString(), hasDocs.ToString())
                    Else
                        dictionary.Add(partNo.ToString(), hasDocs.ToString())
                    End If
                Else
                    dictionary.Add(partNo.ToString(), hasDocs.ToString())
                End If

                'Next
                Return dictionary
            End If
            'End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Function

    'Private Function buildSearchQuerySintax(tt As Object) As String
    Private Function buildSearchQuerySintax(lstSelected As List(Of Object), flag As Integer) As String
        Dim exMessage As String = " "
        Try
            Dim dictionary As New Dictionary(Of String, String)
            'Dim dictionary1 As New Dictionary(Of String, String)

            If flag = 1 Then
                dictionary.Add("txtMfrNoSearch", "PRDMFR#")
                dictionary.Add("txtsearchctp", "PRDCTP")
                dictionary.Add("cmbstatus1", "PRDSTS")
                dictionary.Add("cmbPrpech", "PRPECH")
                dictionary.Add("txtsearchpart", "PRDPTN")
                dictionary.Add("txtsearch1", "A2.VMVNUM")
                dictionary.Add("txtJiratasksearch", "PRDJIRA")
                dictionary.Add("txtsearchcode", "A1.PRHCOD")
                dictionary.Add("txtsearch", "PRNAME")
            Else
                dictionary.Add("cmbUser2", "PRDUSR")
                dictionary.Add("txtPartNoMore", "PRDPTN")
                dictionary.Add("txtCtpNoMore", "PRDCTP")
                dictionary.Add("txtMfrNoMore", "PRDMFR#")
            End If

            Dim strwhere As String = Nothing
            For Each tt As Object In lstSelected
                For Each pair As KeyValuePair(Of String, String) In dictionary
                    If tt.Name = pair.Key Then
                        If flagallow = 1 Then
                            If TypeOf tt Is Windows.Forms.TextBox Then
                                If pair.Key = "txtsearch" Then
                                    strwhere += " AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                                Else
                                    strwhere += " AND TRIM(UCASE(" & pair.Value & ")) = '" & Trim(UCase(tt.Text)) & "' "
                                End If

                            Else
                                strwhere += " AND TRIM(UCASE(" & pair.Value & ")) = '" & Trim(UCase(tt.SelectedValue)) & "' "
                            End If
                        Else
                            If TypeOf tt Is Windows.Forms.TextBox Then
                                If pair.Key = "txtsearch" Then
                                    'strwhere += " AND (PRPECH = '" & userid & "' OR PRDUSR = '" & userid & "') AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                                    strwhere += " AND TRIM(UCASE(PRNAME)) LIKE '%" & Replace(Trim(UCase(tt.Text)), "'", "") & "%'"
                                Else
                                    strwhere += " AND TRIM(UCASE(" & pair.Value & ")) = '" & Trim(UCase(tt.Text)) & "' "
                                End If
                            Else
                                strwhere += " AND TRIM(UCASE(" & pair.Value & ")) = '" & Trim(UCase(tt.SelectedValue)) & "' "
                            End If
                            'strwhere = "WHERE PRPECH = '" & UserID & "' AND TRIM(UCASE(PRDSTS)) = '" & Trim(Left(cmbstatus1.Text, 2)) & "' "
                        End If
                    End If
                Next
            Next

            Return strwhere
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Sub cmdcomments_Click(sender As Object, e As EventArgs) Handles cmdcomments.Click
        gnr.seeaddprocomments = 5
        frmproductsdevelopmentcomments.lblNotVisible.Text = gnr.seeaddprocomments
        frmproductsdevelopmentcomments.ShowDialog()
    End Sub

    Private Sub cmdseecomments_Click(sender As Object, e As EventArgs) Handles cmdseecomments.Click
        gnr.seeaddprocomments = 5
        frmPDevelopmentseecomments.lblNotVisible.Text = gnr.seeaddprocomments
        frmPDevelopmentseecomments.Show()
    End Sub

    Private Sub cmdseefiles_Click(sender As Object, e As EventArgs) Handles cmdseefiles.Click
        cmdseefiles_Click()
    End Sub

    Private Sub cmdfiles_Click(sender As Object, e As EventArgs) Handles cmdfiles.Click
        cmdfiles_Click()
    End Sub

    Private Sub cmdfiles_Click()
        Dim exMessage As String = " "
        Dim fullFilename As String
        Try
            If Trim(txtCode.Text) <> "" And Trim(txtpartno.Text) <> "" Then
                gnr.FolderPath = gnr.Path & "PDevelopment"
                gnr.folderpathvendor = gnr.FolderPath & "\" & Trim(txtCode.Text)

                If Not Directory.Exists(gnr.folderpathvendor) Then
                    System.IO.Directory.CreateDirectory(gnr.folderpathvendor)
                End If
                gnr.folderpathproject = gnr.folderpathvendor & "\" & Trim(UCase(txtpartno.Text)) & "\"
                If Not Directory.Exists(gnr.folderpathproject) Then
                    System.IO.Directory.CreateDirectory(gnr.folderpathproject)
                End If

                Using ofd As New OpenFileDialog
                    ' Give the user some info:
                    ofd.Title = "Select file to copy"
                    ofd.InitialDirectory = "C:\"
                    ' Set the file filter, it looks bad if it is empty.
                    'ofd.Filter = "All files (*.*)/*.*"
                    If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                        fullFilename = ofd.FileName
                    Else
                        ' error message
                        Exit Sub
                    End If
                End Using

                Dim destinationFilename As String = IO.Path.GetFileName(fullFilename)

                If System.IO.File.Exists(gnr.folderpathproject & destinationFilename) = True Then
                    Dim result As DialogResult = MessageBox.Show("Sorry,the file had existed in the folder! Do you want to replace it?", "CTP System", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        My.Computer.FileSystem.CopyFile(fullFilename, gnr.folderpathproject & destinationFilename, True)
                        MessageBox.Show("File had been copy successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    My.Computer.FileSystem.CopyFile(fullFilename, gnr.folderpathproject & destinationFilename, True)
                    MessageBox.Show("File had been copy successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                'Dim saveFileDialog1 As New SaveFileDialog()
                'saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif"
                ''saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif"
                'saveFileDialog1.Title = ""
                'saveFileDialog1.InitialDirectory = "C: \"
                'saveFileDialog1.ShowDialog()
                'If saveFileDialog1.FileName <> "" Then
                '    Dim fs As FileStream = DirectCast(saveFileDialog1.OpenFile(), FileStream)
                '    fs.Close()
                'End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub cmdseefiles_Click()
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" And Trim(txtpartno.Text) <> "" Then
                gnr.FolderPath = gnr.UrlPDevelopmentMethod
                gnr.folderpathvendor = gnr.FolderPath & Trim(txtCode.Text)
                gnr.folderpathproject = gnr.folderpathvendor & "\" & Trim(txtpartno.Text) & "\"
                If Directory.Exists(gnr.folderpathproject) Then
                    Using fbd As OpenFileDialog = New OpenFileDialog()
                        fbd.Title = "Open"
                        fbd.InitialDirectory = gnr.folderpathproject
                        Dim result As DialogResult = fbd.ShowDialog()

                        If result = DialogResult.OK And Not String.IsNullOrWhiteSpace(fbd.SafeFileName) Then
                            gnr.startProcessOF(gnr.folderpathproject & Trim(fbd.SafeFileName))
                        Else
                            'error message
                        End If
                    End Using
                Else
                    MessageBox.Show("There is not a directory in the seleted path.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("The Project Number and Part Number are mandatory fields.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        Exit Sub
    End Sub

    Private Sub cmdfilespart_Click(sender As Object, e As EventArgs) Handles cmdfilespart.Click
        cmdfilespart_Click()
    End Sub

    Private Sub cmdseeqcontrol_Click(sender As Object, e As EventArgs) Handles cmdseeqcontrol.Click
        cmdseeqcontrol_Click()
    End Sub

    Private Sub cmdfilespart_Click()
        Dim PartNo As String
        Dim fullFilename As String
        Dim exMessage As String = " "
        Try
            If Trim(txtpartno.Text) <> "" Then
                'fieldpart = Trim(txtpartno.Text)
                PartNo = Trim(txtpartno.Text)

                gnr.FolderPath = gnr.Path & "PartsFiles"
                gnr.folderpathpart = gnr.FolderPath & "\" & Trim(UCase(PartNo)) & "\"

                If Not Directory.Exists(gnr.folderpathpart) Then
                    System.IO.Directory.CreateDirectory(gnr.folderpathpart)
                End If

                Using ofd As New OpenFileDialog
                    ' Give the user some info:
                    ofd.Title = "Open"
                    ofd.InitialDirectory = "C:\"
                    ' Set the file filter, it looks bad if it is empty.
                    'ofd.Filter = "All files (*.*)/*.*"
                    If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                        fullFilename = ofd.FileName
                    Else
                        ' error message
                        Exit Sub
                    End If
                End Using

                Dim destinationFilename As String = IO.Path.GetFileName(fullFilename)

                If System.IO.File.Exists(gnr.folderpathpart & destinationFilename) = True Then
                    Dim result As DialogResult = MessageBox.Show("Sorry,the file had existed in the folder! Do you want to replace it?", "CTP System", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        My.Computer.FileSystem.CopyFile(fullFilename, gnr.folderpathpart & destinationFilename, True)
                        MessageBox.Show("File added to Part No." & PartNo, "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    My.Computer.FileSystem.CopyFile(fullFilename, gnr.folderpathpart & destinationFilename, True)
                    MessageBox.Show("File added to Part No." & PartNo, "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Dim rsUpdate = gnr.UpdateInvByPhotoAddition(PartNo)
                    If rsUpdate = "" Then
                        'check functionality
                    End If
                End If
            Else
                MessageBox.Show("The part number is a mandatory field.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'MsgBox "You didn't select any file.", vbOKOnly + vbInformation, "CTP System"
        'MsgBox "Select Part to add files.", vbOKOnly + vbInformation, "CTP System"
    End Sub

    Private Sub cmdseeqcontrol_Click()
        Dim PartNo As String
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" And Trim(txtpartno.Text) <> "" Then
                PartNo = Trim(txtpartno.Text)
                gnr.FolderPath = gnr.Path & "PartsFiles"
                gnr.folderpathpart = gnr.FolderPath & "\" & Trim(UCase(PartNo)) & "\"
                If Directory.Exists(gnr.folderpathpart) Then
                    Using fbd As OpenFileDialog = New OpenFileDialog()
                        fbd.Title = "Open"
                        fbd.InitialDirectory = gnr.folderpathpart
                        Dim result As DialogResult = fbd.ShowDialog()

                        If result = DialogResult.OK And Not String.IsNullOrWhiteSpace(fbd.SafeFileName) Then
                            gnr.startProcessOF(gnr.folderpathproject & Trim(fbd.SafeFileName))
                        Else
                            'error message
                        End If
                    End Using
                Else
                    MessageBox.Show("There is not a dairectory in the selected path.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'MsgBox "No files for this Part #.", vbOKOnly + vbInformation, "CTP System"
        'MsgBox "Select Project and Part # to see files.", vbOKOnly + vbInformation, "CTP System"
    End Sub

    Private Sub cmdcvendor_Click(sender As Object, e As EventArgs) Handles cmdcvendor.Click
        cmdcvendor_Click()
    End Sub

    Private Sub cmdcvendor_Click()
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" Then
                If dgvProjectDetails.Rows.Count > 0 Then
                    frmproductsdevelopmentvendor.Show()
                End If
                'actualizo el detalle
                If SSTab1.SelectedIndex = 2 Then
                    If Trim(txtpartno.Text) <> "" Then
                    End If
                End If
                'fillcell2(txtCode.Text)
            Else
                If Trim(txtpartno.Text) <> "" Then
                    Dim result As DialogResult = MessageBox.Show("Select Project.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdmpartno_Click(sender As Object, e As EventArgs) Handles cmdmpartno.Click
        cmdmpartno_Click()
    End Sub

    Private Sub cmdmpartno_Click()
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" Then
                If dgvProjectDetails.Rows.Count > 0 Then
                    frmproductsdevelopmentmanu.ShowDialog()
                End If
                'actualizo el detalle
                If SSTab1.SelectedIndex = 2 Then
                    If Trim(txtpartno.Text) <> "" Then
                    End If
                End If
                'fillcell2(txtCode.Text)
            Else
                If Trim(txtpartno.Text) <> "" Then
                    Dim result As DialogResult = MessageBox.Show("Select Project.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdunitcost_Click(sender As Object, e As EventArgs) Handles cmdunitcost.Click
        cmdunitcost_Click()
    End Sub

    Private Sub cmdunitcost_Click()
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" Then
                If dgvProjectDetails.Rows.Count > 0 Then
                    frmproductsdevelopmentunitcost.ShowDialog()
                End If
                'actualizo el detalle
                If SSTab1.SelectedIndex = 2 Then
                    If Trim(txtpartno.Text) <> "" Then
                    End If
                End If
                'fillcell2(txtCode.Text)
            Else
                If Trim(txtpartno.Text) <> "" Then
                    Dim result As DialogResult = MessageBox.Show("Select Project.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdchange_Click(sender As Object, e As EventArgs) Handles cmdchange.Click
        cmdchange_Click()
    End Sub

    Private Sub cmdchange_Click()
        Dim exMessage As String = " "
        Try
            If Trim(txtCode.Text) <> "" Then
                If dgvProjectDetails.Rows.Count > 0 Then
                    frmproductsdevelopmentstatus.ShowDialog()
                End If
                'actualizo el detalle
                If SSTab1.SelectedIndex = 2 Then
                    If Trim(txtpartno.Text) <> "" Then
                    End If
                End If
                'fillcell2(txtCode.Text)
            Else
                If Trim(txtpartno.Text) <> "" Then
                    Dim result As DialogResult = MessageBox.Show("Select Project.", "CTP System", MessageBoxButtons.OK)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub btnAll_Click(sender As Object, e As EventArgs)
        Dim exMessage As String = " "
        Dim strQueryPartNo = " AND PRDPTN = '" & txtPartNoMore.Text & "'"
        Dim strQueryMfrNo = " AND PRDMFR# = '" & txtMfrNoMore.Text & "'"
        Dim strQueryCtpNo = " AND PRDCTP = '" & txtCtpNoMore.Text & "'"
        Try
            'Dim foo as String = If(bar = buz, cat, dog)
            'ternary
            Dim strPartNo As String = If(Not String.IsNullOrEmpty(txtPartNoMore.Text), strQueryPartNo, "")
            Dim strMfrNo As String = If(Not String.IsNullOrEmpty(txtMfrNoMore.Text), strQueryMfrNo, "")
            Dim strCtpNo As String = If(Not String.IsNullOrEmpty(txtCtpNoMore.Text), strQueryCtpNo, "")

            If String.IsNullOrEmpty(txtPartNoMore.Text) And String.IsNullOrEmpty(txtMfrNoMore.Text) And String.IsNullOrEmpty(txtCtpNoMore.Text) Then
                fillcell2(txtCode.Text)
            Else
                sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM 
                    WHERE PRHCOD = " & txtCode.Text & " " + strPartNo + strMfrNo + strCtpNo 'DELETE BURNED REFERENCE

                'fillcell22(sql)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cleanMoreBtns(controlName As String)
        Dim exMessage As String = " "
        Try
            Dim listData As New List(Of String)(New String() {"txtMfrNoMore", "txtPartNoMore", "txtCtpNoMore", "cmdcmbUser2"})
            For Each tt As String In listData
                'Dim testVal = controlName.Substring(3, controlName.Length - 3)
                If Not tt.Contains(controlName.Substring(3, controlName.Length - 3)) Then
                    If TypeOf (Me.Controls.Find(tt, True).FirstOrDefault()) Is Windows.Forms.TextBox Then
                        Dim tbx As System.Windows.Forms.TextBox = TryCast((Me.Controls.Find(tt, True).FirstOrDefault()), System.Windows.Forms.TextBox)
                        tbx.Text = Nothing
                    ElseIf TypeOf (Me.Controls.Find(tt, True).FirstOrDefault()) Is Windows.Forms.ComboBox Then
                        Dim cmb As System.Windows.Forms.ComboBox = TryCast((Me.Controls.Find(tt, True).FirstOrDefault()), System.Windows.Forms.ComboBox)
                        cmb.SelectedIndex = 0
                    End If
                End If
            Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub Cmdjira_Click_1(sender As Object, e As EventArgs) Handles Cmdjira.Click
        Dim exMessage As String = " "
        Dim jiraUrl As String
        Try
            If Trim(txtCode.Text <> "" And txtpartno.Text <> "") Then
                Dim dsJiraPath = Trim(gnr.GetJiraPath())
                If dsJiraPath IsNot Nothing Then
                    Dim dsDataToDislpay = gnr.GetDataByCodeAndPartNo(txtCode.Text, txtpartno.Text)
                    If dsDataToDislpay IsNot Nothing Then
                        If dsDataToDislpay.Tables(0).Rows.Count > 0 Then
                            Dim jiraNumber = Trim(dsDataToDislpay.Tables(0).Rows(0).ItemArray((dsDataToDislpay.Tables(0).Columns("PRDJIRA").Ordinal).ToString()))
                            If Not String.IsNullOrEmpty(jiraNumber) Then
                                jiraUrl = dsJiraPath + jiraNumber
                                If System.Uri.IsWellFormedUriString(jiraUrl, UriKind.Absolute) Then
                                    Process.Start(jiraUrl)
                                Else
                                    MessageBox.Show("The jira url is not well formed.", "CTP System", MessageBoxButtons.OK)
                                End If
                            Else
                                MessageBox.Show("The jira field must have value.", "CTP System", MessageBoxButtons.OK)
                            End If
                        End If
                    End If
                Else
                    MessageBox.Show("There is a not base jira path value.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("The project number and part number are mandatory fields.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        Dim exMessage As String = Nothing
        Try
            ' code here
            flagdeve = 0
            flagnewpart = 0
            save()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

#End Region

#Region "Delegate"

    Private Delegate Sub closeMDIFormDelegate()

#Region "Delegate Methods"

    Private Sub execute_delegate_MDIClose()
        dspCall.BeginInvoke(New closeMDIFormDelegate(AddressOf closeMDIForm))
    End Sub

    Public Shared Sub closeMDIForm()
        MDIMain.Hide()
    End Sub

#End Region

#End Region

#Region "Utils"

    Public Shared Function GetComputerName() As String
        Dim exMessage As String = Nothing
        Try
            Dim ComputerName As String
            ComputerName = Environment.MachineName
            Return ComputerName
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Public Sub writeComputerEventLog(Optional strMessage As String = Nothing)
        Dim exMessage As String = Nothing
        Try

            If Not EventLog.SourceExists("CTPSystem-Net") Then
                EventLog.CreateEventSource("CTPSystem-Net", "CTPSystem-Log")
            End If
            'EventLog.CreateEventSource("CTPSystem-Net", "CTPSystem-Log")

            Dim lgSource = If(Not String.IsNullOrEmpty(gnr.Source), gnr.Source, "CTPSystem-Net")
            Dim lgName = If(Not String.IsNullOrEmpty(gnr.LogName), gnr.LogName, "CTPSystem-Log")
            Dim msg = If(Not String.IsNullOrEmpty(strMessage), strMessage, "Info: Session started for: " & Environment.UserName)

            eventLog1 = New EventLog(lgName, Environment.MachineName, lgSource)
            eventLog1.WriteEntry(msg, EventLogEntryType.Information)

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub writeLog(strLogCadenaCabecera As String, strLevel As VBLog.ErrorTypeEnum, strMessage As String, strDetails As String)
        strLogCadena = strLogCadenaCabecera + " " + System.Reflection.MethodBase.GetCurrentMethod().ToString()

        vblog.WriteLog(strLevel, "CTPSystem" & strLevel, strLogCadena, userid, strMessage, strDetails)
    End Sub

    Public Function checkPendingReferences(code As String) As String
        Dim exMessage As String = Nothing
        Dim notContainsStatus As Boolean
        Dim strResult As String = Nothing
        Try
            Dim dsStatuses = gnr.getReferencesStatusesByCode(Trim(txtCode.Text))
            Dim optStatuses As String() = gnr.GetCloseStatus.Split(",")
            Dim lstStatuses As List(Of String) = New List(Of String)()
            For Each item As String In optStatuses
                lstStatuses.Add(item)
            Next
            If dsStatuses IsNot Nothing Then

                If dsStatuses.Tables(0).Rows.Count > 1 Then

                    notContainsStatus = dsStatuses.Tables(0).AsEnumerable().Any(Function(x) Not lstStatuses.Contains(LCase(x.ItemArray(0).ToString())))

                    'Dim statusSel = LCase(dw.Item("prdsts").ToString())
                    'notContainsStatus = optStatuses.AsEnumerable().Any(Function(x) x <> statusSel)

                    'For Each dw As DataRow In dsStatuses.Tables(0).Rows
                    '    Dim statusSel = LCase(dw.Item("prdsts").ToString())
                    '    notContainsStatus = optStatuses.AsEnumerable().Any(Function(x) x <> statusSel)

                    '    'For Each item As String In optStatuses
                    '    '    If item = dw.Item("prdsts").ToString() Then
                    '    '        notContainsStatus = True
                    '    '        Exit For
                    '    '    End If
                    '    'Next

                    'Next

                Else
                    For Each item As String In optStatuses
                        If item = dsStatuses.Tables(0).Rows(0).ItemArray(0).ToString() Then
                            notContainsStatus = True
                            Exit For
                        End If
                    Next

                End If

                'For Each dw As DataRow In dsStatuses.Tables(0).Rows
                '    Dim statusSel = LCase(dw.ItemArray(0).ToString())
                '    'Dim notContainsStatus = optStatuses.AsEnumerable().Any(Function(x) x = statusSel)


                '    For Each item As String In optStatuses
                '        If item = statusSel Then
                '            notContainsStatus = True
                '            Exit For
                '        End If
                '    Next

                If Not notContainsStatus Then
                    'If cmbprstatus.SelectedIndex <> 1 Then
                    'cmbprstatus.SelectedIndex = 1
                    Dim rs = gnr.UpdateGeneralStatus(code, cmbprstatus.SelectedItem)
                        If rs < 1 Then
                        Else
                        strResult = "F"
                        'DataGridView1.Refresh()
                    End If
                    'End If
                    Return strResult
                Else
                    'If cmbprstatus.SelectedIndex <> 2 Then
                    'cmbprstatus.SelectedIndex = 2
                    Dim rs = gnr.UpdateGeneralStatus(code, cmbprstatus.SelectedItem)
                        If rs < 1 Then
                        Else
                        strResult = "I"
                        'DataGridView1.Refresh()
                    End If
                    'End If
                    Return strResult
                End If
                'Next
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function IsFileinUse(file As FileInfo) As Boolean
        Dim exMessage As String = Nothing
        Dim opened As Boolean = False
        Dim myStream As FileStream = Nothing
        Try
            myStream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            opened = True
            LikeSession.excelOpened = opened
        Finally
            If myStream IsNot Nothing Then
                LikeSession.excelOpened = False
                myStream.Close()
            End If
        End Try
        Return opened
    End Function

    Public Sub saveExcelReport(dt As DataTable, folderPath As String, fileName As String)
        Dim exMessage As String = " "
        Try
            Dim fullPath = folderPath & Convert.ToString(fileName)
            Using wb As New XLWorkbook()

                wb.Worksheets.Add(dt, "Project")
                wb.SaveAs(fullPath)

            End Using

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Sub launchExcelReport(fullPath As String, folderPath As String)
        Dim exMessage As String = " "
        Try
            Dim rsConfirm As DialogResult = MessageBox.Show("The file was created successfully .Do you want to see it (Yes) or to open the created document location (No)?", "CTP System", MessageBoxButtons.YesNo)
            If rsConfirm = DialogResult.Yes Then
                Dim newFile As FileInfo = New FileInfo(fullPath)
                If newFile.Exists Then
                    System.Diagnostics.Process.Start(fullPath)
                End If
            Else
                Try
                    Process.Start("explorer.exe", folderPath)
                Catch Win32Exception As Win32Exception
                    Shell("explorer " & folderPath, AppWinStyle.NormalFocus)
                End Try
            End If

        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    'dim p2Path as String = "\Excel-Template\"
    Public Sub prepareFilePath(folderPath As String)
        Dim exMessage As String = " "
        Try
            'Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            'Dim folderPath As String = userPath & p2Path
            'Dim fixedFolderPath = folderPath.Replace("\", "\\")
            'Dim sourcePath As String = gnr.getPdExcelTemplate

            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            Else
                Dim files = Directory.GetFiles(folderPath)
                Dim fi = Nothing
                'If files.Length = 1 Then
                For Each item In files
                    fi = item
                    Dim isOpened = IsFileinUse(New FileInfo(fi))
                    If Not isOpened Then
                        File.Delete(item)
                    Else
                        Dim rsError As DialogResult = MessageBox.Show("Please close the file " & fi & " in order to proceed!", "CTP System", MessageBoxButtons.OK)
                        Exit Sub
                    End If
                Next
                'Else
                '    Dim rsError As DialogResult = MessageBox.Show("Please close the file located in " & folderPath & " in order to proceed!", "CTP System", MessageBoxButtons.OK)
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub InactiveQotaAlertExcelGeneration(ds As DataSet, userid As String, ByRef created As Boolean, title As String)
        Dim exMessage As String = " "
        Try
            Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim folderPath As String = userPath & "\CTP-NEW-DOCS\"
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            'delete if previous documents
            prepareFilePath(folderPath)

            Dim dt As New DataTable
            dt = ds.Tables(0)
            'dt = (DirectCast(DataGridView2.DataSource, DataTable))
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim fileExtension As String = Determine_OfficeVersion()
                    If String.IsNullOrEmpty(fileExtension) Then
                        Exit Sub
                    End If

                    'Dim title As String
                    'title = "Inactivity report for " & userid & " running at "
                    Dim fileName = gnr.adjustDatetimeFormat(title, fileExtension)

                    'If Not String.IsNullOrEmpty(txtProjectNo.Text) Then
                    '    fileName = "Excel Custon Report for vendor " & vendorNo & " And Status " & status & " running in - " & DateTime.Now.ToString("d") & "." & fileExtension
                    'Else
                    '    fileName = "Project Name " & txtProjectName.Text & " - Errors. The project does not have a number yet." & fileExtension
                    'End If

                    Dim fullPath = folderPath & Convert.ToString(fileName)
                    saveExcelReport(dt, folderPath, fileName)

                    If File.Exists(fullPath) Then
                        launchExcelReport(fullPath, folderPath)
                        created = True
                    End If
                Else
                    MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub prodDevExcelGeneration(ds As DataSet, vendorNo As String, status As String, title As String)
        Dim exMessage As String = " "
        Try
            Dim strFolder = "\CTP-PD-Reports\"
            Dim userPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim folderPath As String = userPath & strFolder
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            'delete if previous documents
            prepareFilePath(folderPath)

            Dim dt As New DataTable
            dt = ds.Tables(0)
            'dt = (DirectCast(DataGridView2.DataSource, DataTable))
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim fileExtension As String = Determine_OfficeVersion()
                    If String.IsNullOrEmpty(fileExtension) Then
                        Exit Sub
                    End If

                    'Dim title As String
                    'title = "Status Report for vendor " & vendorNo & " And Status " & status & " requested by " & userid & " running at "
                    Dim fileName = gnr.adjustDatetimeFormat(title, fileExtension)


                    'If Not String.IsNullOrEmpty(txtProjectNo.Text) Then
                    '    fileName = "Excel Custon Report for vendor " & vendorNo & " And Status " & status & " running in - " & DateTime.Now.ToString("d") & "." & fileExtension
                    'Else
                    '    fileName = "Project Name " & txtProjectName.Text & " - Errors. The project does not have a number yet." & fileExtension
                    'End If

                    Dim fullPath = folderPath & Convert.ToString(fileName)
                    saveExcelReport(dt, folderPath, fileName)

                    If File.Exists(fullPath) Then
                        launchExcelReport(fullPath, folderPath)
                        'Dim rsConfirm As DialogResult = MessageBox.Show("The file was created successfully in this path " & folderPath & " .Do you want to open the created document location?", "CTP System", MessageBoxButtons.YesNo)
                    End If
                Else
                    MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                MessageBox.Show("There is not results to print to an excel document.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function Determine_OfficeVersion() As String
        Dim exMessage As String = " "
        Dim strExt As String = Nothing
        Try
            Dim strEVersionSubKey As String = "\Excel.Application\CurVer" '/HKEY_CLASSES_ROOT/Excel.Application/Curver

            Dim strValue As String 'Value Present In Above Key
            Dim strVersion As String 'Determines Excel Version
            Dim strExtension() As String = {"xls", "xlsx"}

            Dim rkVersion As RegistryKey = Nothing 'Registry Key To Determine Excel Version
            rkVersion = Registry.ClassesRoot.OpenSubKey(name:=strEVersionSubKey, writable:=False) 'Open Registry Key

            If Not rkVersion Is Nothing Then 'If Key Exists
                strValue = rkVersion.GetValue(String.Empty) 'get Value
                strValue = strValue.Substring(strValue.LastIndexOf(".") + 1) 'Store Value

                Select Case strValue 'Determine Version
                    Case "7"
                        strVersion = "95"
                        strExt = strExtension(0)
                    Case "8"
                        strVersion = "97"
                        strExt = strExtension(0)
                    Case "9"
                        strVersion = "2000"
                        strExt = strExtension(0)
                    Case "10"
                        strVersion = "2002"
                        strExt = strExtension(0)
                    Case "11"
                        strVersion = "2003"
                        strExt = strExtension(0)
                    Case "12"
                        strVersion = "2007"
                        strExt = strExtension(1)
                    Case "14"
                        strVersion = "2010"
                        strExt = strExtension(1)
                    Case "15"
                        strVersion = "2013"
                        strExt = strExtension(1)
                    Case "16"
                        strVersion = "2016"
                        strExt = strExtension(1)
                    Case Else
                        strExt = strExtension(1)
                End Select

                Return strExt
            Else
                MessageBox.Show("Microsoft Excel is not installed or corrupt in this computer.", "CTP System", MessageBoxButtons.OK)
                Return strExt
            End If
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return strExt
        End Try
    End Function

    Public Sub showTab2FilterPanel(dgv As DataGridView)
        Dim flag = If(dgv.DataSource Is Nothing, False, True)
        'Dim flag = False
        'TableLayoutPanel15.Visible = False
        txtCtpNoMore.Visible = flag
        txtMfrNoMore.Visible = flag
        txtPartNoMore.Visible = flag
        cmdAll1.Visible = flag
        LinkLabel2.Visible = flag
        dgvProjectDetails.Visible = flag
        cmdcvendor.Visible = flag
        cmdunitcost.Visible = flag
        cmdchange.Visible = flag
        cmdmpartno.Visible = flag
        cmbuser2.Visible = flag
    End Sub

    Public Function FindFocussedControl(ByVal ctr As Control) As Control
        Dim exMessage As String = Nothing
        Try
            Dim container As ContainerControl = TryCast(ctr, ContainerControl)
            Do While (container IsNot Nothing)
                ctr = container.ActiveControl
                container = TryCast(ctr, ContainerControl)
            Loop
            Return ctr
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Sub cleanDataSources()
        bs.DataSource = Nothing
        bs1.DataSource = Nothing
        DataGridView1.DataSource = Nothing
        'DataGridView1.DataSource.Clear()
        DataGridView1.Refresh()
        dgvProjectDetails.DataSource = Nothing
        'dgvProjectDetails.DataSource.Clear()
        dgvProjectDetails.Refresh()
        DataGridView1.Visible = False
        dgvProjectDetails.Visible = False
    End Sub

    Private Sub setVendorValues()
        Dim exMessage As String = " "
        Try
            Dim vendors = If(Not String.IsNullOrEmpty(txtCode.Text), gnr.GetVendorInProject(CInt(txtCode.Text)), Nothing)
            If vendors.Count > 1 Then
                MessageBox.Show("There is more than one vendor in this project.", "CTP System", MessageBoxButtons.OK)
                txtvendornoa.Text = vendors(0)
                txtvendorno.Text = txtvendornoa.Text
                Dim dsVnd = gnr.GetVendorByVendorNo(txtvendornoa.Text)
                txtvendornamea.Text = dsVnd.Tables(0).Rows(0).ItemArray(2).ToString()
                txtvendorname.Text = txtvendornamea.Text
            Else
                txtvendornoa.Text = vendors(0)
                txtvendorno.Text = txtvendornoa.Text
                Dim dsVnd = gnr.GetVendorByVendorNo(txtvendornoa.Text)
                txtvendornamea.Text = dsVnd.Tables(0).Rows(0).ItemArray(2).ToString()
                txtvendorname.Text = txtvendornamea.Text
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function itemCategory(partNo As String, vendorNo As String) As Integer
        Dim exMessage As String = " "
        Dim result As Integer = -1
        Try
            If String.IsNullOrEmpty(partNo) Then
                Return 2
            Else
                Dim listItemCat = gnr.VendorWhiteFlagMethod.Split(",")

                Dim dsResult1 = gnr.getItemCategoryByVendorAndPart(vendorNo, partNo)
                If dsResult1 IsNot Nothing Then
                    If dsResult1.Tables(0).Rows.Count > 0 Then
                        For Each item As String In listItemCat
                            If Trim(item).Equals(Trim(vendorNo)) Then
                                Return 2
                            End If
                        Next
                        Return -1
                    Else
                        Return 2
                    End If
                Else
                    Return 2
                End If

                Return result
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return result
        End Try

    End Function

    Private Sub LoadCombos(Optional ByVal sender As Object = Nothing, Optional ByVal e As EventArgs = Nothing)

        BackgroundWorker1.RunWorkerAsync()
        Loading.ShowDialog()
        Loading.BringToFront()

    End Sub

    Private Sub ExecuteCombos(Optional ByVal sender As Object = Nothing, Optional ByVal e As EventArgs = Nothing)

        'FillDDlUser() 'Fill user cmb
        FillDDlUser1()
        FillDDlUser2()
        FillDDLStatus()
        FillDDlMinorCode()
        FillDDlMajorCode()

        cmbprstatus.Items.Add("-- Select Status --")
        cmbprstatus.Items.Add("I - In Process")
        cmbprstatus.Items.Add("F - Finished")
        cmbprstatus.SelectedIndex = 0

        Dim posValue As Integer = 0
        For Each obj As DataRowView In cmbstatus.Items
            Dim VarQuery = "E"
            Dim VarCombo = Trim(obj.Item(2).ToString())
            If VarQuery = VarCombo Then
                cmbstatus.SelectedIndex = posValue
                Exit For
            Else
                posValue += 1
            End If
        Next

        Dim bgWorker = CType(sender, BackgroundWorker)
        For index = 0 To 2
            bgWorker.ReportProgress(index)
            Threading.Thread.Sleep(2000)
        Next

    End Sub

    Private Sub SetValues()
        Dim exMessage As String = Nothing
        Try
            SSTab1.ItemSize = (New Size((SSTab1.Width - 50) / SSTab1.TabCount, 0))
            SSTab1.Padding = New System.Drawing.Point(300, 10)
            SSTab1.Appearance = TabAppearance.FlatButtons
            SSTab1.SizeMode = TabSizeMode.Fixed

            TabPage3.Text = "Reference: "
            TabPage3.AutoScroll = True
            TabPage3.AutoScrollPosition = New Point(0, TabPage3.VerticalScroll.Maximum)
            TabPage3.Padding = New Padding(0, 0, SystemInformation.VerticalScrollBarWidth, 0)
            TabPage3.AutoScrollMinSize = New Drawing.Size(800, 0)

            TabPage2.Text = "Project: "

            'AddHandler vScrollBar1.Scroll, AddressOf vScrollBar1_Scroll
            ''vScrollBar1.Scroll += (sender, e) >= {Panel1.VerticalScroll.Value = vScrollBar1.Value; };
            'TabPage2.Controls.Add(vScrollBar1)
            'AddHandler frmProductsDevelopment.dgvProjectDetails_CellContentClick, AddressOf dgvProjectDetails_CellContentClick

            Me.WindowState = FormWindowState.Maximized

            cmdSave1.Enabled = False

            Button8.Enabled = False
            Button9.Enabled = False
            Button10.Enabled = False
            Button11.Enabled = False
            Button15.Enabled = False
            Button16.Enabled = False
            Button17.Enabled = False
            Button18.Enabled = False

            cmdsearch.FlatStyle = FlatStyle.Flat
            cmdsearchcode.FlatStyle = FlatStyle.Flat
            cmdsearch1.FlatStyle = FlatStyle.Flat
            cmdsearchpart.FlatStyle = FlatStyle.Flat
            cmdsearchctp.FlatStyle = FlatStyle.Flat
            cmdstatus1.FlatStyle = FlatStyle.Flat
            cmdall.FlatStyle = FlatStyle.Flat
            cmdJiratasksearch.FlatStyle = FlatStyle.Flat
            cmdPrpech.FlatStyle = FlatStyle.Flat
            cmdMfrNoSearch.FlatStyle = FlatStyle.Flat
            chknew.Enabled = False
            chkSupplier.Enabled = False

            DataGridView1.RowHeadersVisible = False
            dgvProjectDetails.RowHeadersVisible = False

            'Button12.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\doc.PNG")
            cmdnew1.ImageAlign = ContentAlignment.MiddleRight
            cmdnew1.TextAlign = ContentAlignment.MiddleLeft

            ' Button13.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\save.PNG")
            cmdSave1.ImageAlign = ContentAlignment.MiddleRight
            cmdSave1.TextAlign = ContentAlignment.MiddleLeft

            'Button14.Image = Image.FromFile("C:\\Users\\aavila\\Documents\\exit.PNG")
            cmdexit1.ImageAlign = ContentAlignment.MiddleRight
            cmdexit1.TextAlign = ContentAlignment.MiddleLeft

            'Datepickers customization
            DTPicker1.Format = DateTimePickerFormat.Custom
            DTPicker1.CustomFormat = "MM/dd/yyyy"

            DTPicker2.Format = DateTimePickerFormat.Custom
            DTPicker2.CustomFormat = "MM/dd/yyyy"

            DTPicker3.Format = DateTimePickerFormat.Custom
            DTPicker3.CustomFormat = "MM/dd/yyyy"

            DTPicker4.Format = DateTimePickerFormat.Custom
            DTPicker4.CustomFormat = "MM/dd/yyyy"

            'extra method
            Panel1.Enabled = True
            cmdSave1.Enabled = False
            cmdexit1.Enabled = True
            cmdnew1.Enabled = True

            Panel4.Enabled = False
            txtCode.Enabled = False
            txtvendorno.ReadOnly = True
            txtvendorname.ReadOnly = True
            txtvendornamea.ReadOnly = True
            txtvendornoa.ReadOnly = True
            txtminor.ReadOnly = True
            txtMajor.ReadOnly = True
            txtpartno.ReadOnly = True
            txtpartdescription.ReadOnly = True
            cmbminorcode.Enabled = False
            cmbmajorcode.Enabled = False
            txtctpno.ReadOnly = True

            optCTP.Checked = True
            optVENDOR.Checked = False
            optboth.Checked = False

            flagdeve = 1
            flagnewpart = 1

            logUser.Text += userid

            'add image to textbox
            'txtsearchcode.SetBtnTexbox(ImageList1)

            'tab 1
            txtsearch1.SetWatermark("Vendor No.")
            txtsearch.SetWatermark("Project Name")
            txtJiratasksearch.SetWatermark("Jira Task No.")
            txtsearchpart.SetWatermark("Part No.")
            txtsearchctp.SetWatermark("CTP No.")
            txtMfrNoSearch.SetWatermark("Manufacturer No.")
            txtsearchcode.SetWatermark("Project No.")
            cmbPrpech.SetWatermark("Person In Charge")

            cmbstatus1.SetWatermark("Project Reference Status")
            'MyComboBox1.SetWatermark("test water")

            'tab 2
            txtPartNoMore.SetWatermark("Part No.")
            txtCtpNoMore.SetWatermark("CTP No.")
            txtMfrNoMore.SetWatermark("Manufacturer No.")
            txtCode.SetWatermark("Project No.")
            txtname.SetWatermark("Project Name")

            cmbuser1.SetWatermark("Person In Charge")
            cmbprstatus.SetWatermark("Project Status")
            cmbuser2.SetWatermark("Person In Charge")
            ContextMenuStrip1.Visible = False

            cmbuser1.SelectedIndex = If(cmbuser1.FindString(Trim(UCase(userid))) <> -1,
                                        cmbuser1.FindString(Trim(UCase(userid))), 0)

            cmbuser.SelectedIndex = If(cmbuser.FindString(Trim(UCase(userid))) <> -1,
                                        cmbuser.FindString(Trim(UCase(userid))), 0)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Function checkIfDocsPresent(code As String, flag As Integer) As Boolean
        Dim exMessage As String = " "
        Dim hasDocs As Boolean = False
        Try
            'Dim projectNo As String = "3221"
            If Trim(code) <> "" Then
                gnr.FolderPath = gnr.UrlPDevelopmentMethod
                gnr.folderpathproject = gnr.FolderPath & Trim(code) & "\"

                If flag = 0 Then
                    'datagrid 1
                    'gnr.folderpathproject = gnr.folderpathvendor & "\" & Trim(txtpartno.Text) & "\"
                    If Directory.Exists(gnr.folderpathproject) Then
                        If Not IsDirectoryEmpty(gnr.folderpathproject) Then
                            Dim lstPdfInside = Directory.GetFiles(gnr.folderpathproject, "*.pdf", SearchOption.AllDirectories)
                            'Dim files = Directory.EnumerateFiles(gnr.folderpathproject, "*.*", SearchOption.AllDirectories) _
                            '.Where(Function(s) s >= s.EndsWith(".pdf", StringComparison.InvariantCultureIgnoreCase))

                            'Dim files1 = Directory.GetFiles(gnr.folderpathproject, "*.*", SearchOption.AllDirectories) _
                            '.Where(Function(s) s >= s.EndsWith(".pdf", StringComparison.InvariantCultureIgnoreCase))

                            ' Dim files = Directory.EnumerateFiles(gnr.folderpathproject, "*.*", SearchOption.AllDirectories) _
                            '.Where(Function(s) s >= s.EndsWith(".pdf", StringComparison.InvariantCultureIgnoreCase) Or s.EndsWith(".doc", StringComparison.InvariantCultureIgnoreCase) _
                            'Or s.EndsWith(".docx", StringComparison.InvariantCultureIgnoreCase))
                            If lstPdfInside IsNot Nothing Then
                                If lstPdfInside.Count > 0 Then
                                    hasDocs = True
                                    Return hasDocs
                                End If
                            End If
                            Return hasDocs
                        End If
                        'Else
                        'test purpose
                        'MessageBox.Show("There is not a directory in the seleted path.", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
            Else
                MessageBox.Show("The Project Number must be filled.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return hasDocs
        End Try
    End Function

    Public Function IsDirectoryEmpty(path As String) As Boolean
        Return Not Directory.EnumerateFileSystemEntries(path).Any()
    End Function

    Private Function GetAmountOfProjectReferences(code As String) As Integer
        Dim ds As DataSet
        Dim exMessage As String = " "
        Try
            sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS,PRDJIRA,PRDUSR FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "  'DELETE BURNED REFERENCE
            'get the query results
            ds = gnr.FillGrid(sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds.Tables(0).Rows.Count
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return 0
        End Try
    End Function

    Private Sub fillSecondTabUpp(code As String)
        Dim ds As DataSet
        Dim exMessage As String = " "
        Try
            ds = gnr.GetDataByPRHCOD(code)
            If ds.Tables(0).Rows.Count = 1 Then

                'SSTab1.SelectedTab = TabPage2
                For Each RowDs In ds.Tables(0).Rows
                    txtCode.Text = Trim(RowDs.Item("PRHCOD").ToString())
                    txtname.Text = Trim(RowDs.Item("PRNAME").ToString()) ' format date
                    TabPage2.Text = "Project: " + txtname.Text

                    Dim CleanDateString As String = Regex.Replace(RowDs.Item("PRDATE").ToString(), "/[^0-9a-zA-Z:]/g", "")
                    'Dim dtChange As DateTime = DateTime.ParseExact(CleanDateString, "MM/dd/yyyy HH:mm:ss tt", CultureInfo.InvariantCulture)
                    Dim dtChange As DateTime = DateTime.Parse(CleanDateString)
                    DTPicker1.Value = dtChange.ToShortDateString()

                    If cmbuser1.FindStringExact(Trim(RowDs.Item("PRPECH").ToString())) Then
                        cmbuser1.SelectedIndex = cmbuser1.FindString(Trim(RowDs.Item("PRPECH").ToString()))
                    End If

                    If cmbuser1.SelectedIndex = -1 Then
                        cmbuser1.SelectedIndex = cmbuser1.Items.Count - 1
                    End If


                    If Trim(RowDs.Item("PRNAME").ToString()) = "I" Then
                        cmbprstatus.SelectedIndex = 1
                    ElseIf Trim(RowDs.Item("PRNAME").ToString()) = "F" Then
                        cmbprstatus.SelectedIndex = 2
                    Else
                        cmbprstatus.SelectedIndex = 2
                    End If
                    'Dim Test1 = RowDs.Item(1).ToString() get the value begans with 0 pos
                    'Dim test2 = ds.Tables(0).Columns.Item(1).ColumnName  get the grid header
                Next
            Else
                'message box warning
            End If

            'fill second grid process
            'clean all other fields
            flagdeve = 0
            flagnewpart = 0
            cmdnew2.Enabled = True
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub copyProjecFiles(strProjectNo As String)
        Dim exMessage As String = Nothing
        Try
            'save files
            gnr.FolderPath = gnr.Path & "PDevelopment"
            gnr.folderpathvendor = gnr.FolderPath & "\" & Trim(txtCode.Text)
            If Not Directory.Exists(gnr.folderpathvendor) Then
                System.IO.Directory.CreateDirectory(gnr.folderpathvendor)
            End If

            gnr.folderpathproject = gnr.folderpathvendor & "\" & Trim(UCase(txtpartno.Text)) & "\"
            If Not Directory.Exists(gnr.folderpathproject) Then
                System.IO.Directory.CreateDirectory(gnr.folderpathproject)
            End If

            gnr.pathfolderfrom = gnr.FolderPath & "\" & strProjectNo & "\" & Trim(UCase(txtpartno.Text)) & "\"
            My.Computer.FileSystem.CopyDirectory(gnr.pathfolderfrom, gnr.folderpathproject)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub searchpart()
        Dim exMessage As String = " "
        Try
            Dim pathPictures = gnr.UrlPathGeneralMethod & "CTPPictures\"
            If Not Directory.Exists(pathPictures) Then
                'System.IO.Directory.CreateDirectory(pathPictures)
                Dim note = "the picture path it is not recognized.Check it please."
                MessageBox.Show(note, "CTP System", MessageBoxButtons.OK)
            End If
            pathpictureparts = pathPictures & "pic_not_av.jpg"
            Dim existsFile As Boolean = File.Exists(pathpictureparts)
            If existsFile Then
                PictureBox1.Load(pathpictureparts)
            End If

            If Trim(txtpartno.Text) <> "" Then
                Dim PartNo = Trim(UCase(txtpartno.Text))
                Dim folderpathproject = gnr.UrlPartFilesMethod & Trim(UCase(txtpartno.Text)) & "\OEM_" & Trim(UCase(PartNo)) & ".jpg"
                If File.Exists(folderpathproject) Then
                    PictureBox1.Load(folderpathproject)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        Exit Sub
    End Sub

    Private Sub changeControlAccess(value As Boolean)
        txtvendorno.ReadOnly = value
        txtvendorname.ReadOnly = value
        txtpartno.ReadOnly = value
        txtpartdescription.ReadOnly = value
        txtvendornoa.ReadOnly = value
        txtvendornamea.ReadOnly = value
        txtminor.ReadOnly = value
        txtCode.ReadOnly = value
    End Sub

    Private Function prepareEmailsToSend(flag As Integer) As String
        Dim exMessage As String = " "
        Dim toemailss As String = ""
        Dim toemailsok As String = ""
        Try
            If flag = 1 Then
                toemailss = prepareEmailSalesDict()
                toemailsok = prepareEmailMktDict(toemailss)
            ElseIf flag = 2 Then
                toemailsok = prepareEmailSalesDict()
            Else
                toemailsok = prepareEmailMktDict()
            End If

            Return toemailsok
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function prepareEmailsToSendReport(flag As Integer) As String
        Dim exMessage As String = " "
        Dim toemailss As String = ""
        Dim toemailsok As String = ""
        Try
            toemailsok = "alexei.ansberto85@gmail.com;ansberto.avila85@gmail.com"

            Return toemailsok
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function prepareEmailSalesDict() As String
        Dim exMessage As String = " "
        Try
            Dim toemailss As String = ""
            Dim dsSls As DataSet
            dsSls = gnr.GetEmailData(1)
            If dsSls IsNot Nothing Then
                If dsSls.Tables(0).Rows.Count > 0 Then
                    For Each tt As DataRow In dsSls.Tables(0).Rows
                        toemailss += Trim(tt.ItemArray(0).ToString()) + ";"
                    Next
                End If
            End If
            Return toemailss
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function prepareEmailMktDict(Optional ByVal toemailss As String = Nothing) As String
        Dim exMessage As String = " "
        Try
            'Dim toemailss As String = ""
            Dim dsMkt As DataSet
            dsMkt = gnr.GetEmailData(2)
            If dsMkt IsNot Nothing Then
                If dsMkt.Tables(0).Rows.Count > 0 Then
                    For Each tt As DataRow In dsMkt.Tables(0).Rows
                        toemailss += Trim(tt.ItemArray(0).ToString()) + ";"
                    Next
                End If
            End If
            Return toemailss
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function mandatoryFields(flag As String, index As Integer, Optional ByVal requireValidation As Integer = 0, Optional ByVal strArrayCheck As String = Nothing) As Integer
        Dim exMessage As String = Nothing
        Dim methodResult As Integer = -1
        Try

            Dim myTableLayout As TableLayoutPanel

            If requireValidation = 1 Then
                If index = 0 Then
                    myTableLayout = Me.TableLayoutPanel1
                ElseIf index = 1 Then
                    myTableLayout = Me.TableLayoutPanel3
                Else
                    myTableLayout = Me.TableLayoutPanel4
                End If

                If flag = "new" Then
                    Dim TextboxQty As Integer
                    Dim TextboxQtyEmpty As Integer
                    For Each tt In myTableLayout.Controls
                        If TypeOf tt Is Windows.Forms.TextBox Then
                            TextboxQty += 1
                            If tt.Text = "" Then
                                If tt.Name <> "txtainfo" And tt.Name <> "txtCode" Then
                                    TextboxQtyEmpty += 1
                                End If
                            End If
                        ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                            If tt.Name = "cmbuser1" Then
                                TextboxQty += 1
                                If tt.Text = "N/A" Then
                                    TextboxQtyEmpty += 1
                                End If
                            End If
                        End If
                    Next

                    If TextboxQtyEmpty <> 0 And TextboxQty > 0 Then
                        methodResult = 1
                    Else
                        methodResult = 0
                    End If

                Else

                    'If index = 1 Or index = 2 Then 'here when define the mandatory fields
                    If index = 1 Then
                        'txtCode.Text = " "

                        Dim empty = myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)().Where(Function(txt) txt.Text.Length = 0 And txt.Name <> "txtCode" And txt.Name <> "txtainfo")
                        If empty.Any Then
                            methodResult = 1
                        Else
                            methodResult = 0
                            'MessageBox.Show(String.Format("Please fill following textboxes: {0}", String.Join(",", empty.Select(Function(txt) txt.Name))))
                        End If
                    Else
                        Dim arrayCheck As New List(Of String)
                        Dim arrayCheckOk As New List(Of String)
                        arrayCheck = strArrayCheck.Split(",").ToList()
                        For Each item As String In arrayCheck
                            If item = "Part Number" Then
                                arrayCheckOk.Add("txtpartno")
                            ElseIf item = "Part Name" Then
                                arrayCheckOk.Add("txtpartdescription")
                            ElseIf item = "Vendor Number" Then
                                arrayCheckOk.Add("txtvendorno")
                            ElseIf item = "Vendor Name" Then
                                arrayCheckOk.Add("txtvendorname")
                            ElseIf item = "CTP Number" Then
                                arrayCheckOk.Add("txtctpno")
                            ElseIf item = "Person in Charge" Then
                                arrayCheckOk.Add("cmbuser")
                            ElseIf item = "Unit Cost New" Then
                                arrayCheckOk.Add("txtunitcostnew")
                            End If
                        Next

                        Dim empty = myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)().Where(Function(txt) txt.Text.Length = 0 And arrayCheckOk.Contains(txt.Name))
                        If empty.Any Then
                            methodResult = 1
                        Else
                            methodResult = 0
                            'MessageBox.Show(String.Format("Please fill following textboxes: {0}", String.Join(",", empty.Select(Function(txt) txt.Name))))
                        End If
                    End If
                End If

                Return methodResult
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return methodResult
        End Try
    End Function

    Private Sub cleanFormValues(tab As String, flag As Integer)
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            Dim myTableLayout1 As TableLayoutPanel
            Dim myPanel As Windows.Forms.Panel
            Dim lstLayouts As New List(Of TableLayoutPanel)

            If tab = "TabPage1" Then
                myTableLayout = Me.TableLayoutPanel1
                SSTab1.TabPages(0).Text = ""
            ElseIf tab = "TabPage2" Then
                myTableLayout = Me.TableLayoutPanel3
                myTableLayout1 = Me.TableLayoutPanel15
                SSTab1.TabPages(1).Text = "Project: "
                lstLayouts.Add(myTableLayout)
                lstLayouts.Add(myTableLayout1)
            Else
                myTableLayout = Me.TableLayoutPanel4
                myTableLayout1 = Me.TableLayoutPanel5
                SSTab1.TabPages(2).Text = "Reference: "
                lstLayouts.Add(myTableLayout)
                lstLayouts.Add(myTableLayout1)
            End If

            For Each ttt In lstLayouts
                For Each tt In ttt.Controls
                    If TypeOf tt Is Windows.Forms.TextBox Then
                        If flag = 1 Then
                            If (tt.Name <> "txtvendorno") And (tt.Name <> "txtvendorname") Then
                                tt.Text = ""
                            End If
                        Else
                            tt.Text = ""
                        End If
                    ElseIf TypeOf tt Is Windows.Forms.ComboBox Then
                        If (tt.name <> "cmbuser") And (tt.name <> "cmbmajorcode") Then
                            tt.selectedIndex = 0
                        End If
                    ElseIf TypeOf tt Is Windows.Forms.DateTimePicker Then
                        tt.Value = DateTime.Now
                    ElseIf TypeOf tt Is Windows.Forms.Panel Then
                        If tt.Name = "Panel7" Then
                            For Each item As Control In tt.Controls
                                Dim item_type = item.GetType().ToString()
                                If item_type.Equals("System.Windows.Forms.PictureBox") Then
                                    Dim controlSender = DirectCast(item, System.Windows.Forms.PictureBox)
                                    controlSender.Image = Nothing
                                End If
                            Next
                        End If
                    Else
                        Dim pepe = tt.Name
                        Dim papa = pepe
                    End If
                Next
            Next

            'TabPage2.Text = Nothing

            If flag = 2 Then
                DataGridView1.DataSource = Nothing
                DataGridView1.Refresh()

                dgvProjectDetails.DataSource = Nothing
                dgvProjectDetails.Refresh()
            End If
            'myTableLayout.Controls.OfType(Of Windows.Forms.TextBox)().Select(Function(ctx) ctx.Text = "")
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cleanValues()

        txtCode.Text = ""
        txtname.Text = ""
        txtainfo.Text = ""
        txtpartno.Text = ""
        txtvendornoa.Text = ""
        txtvendornamea.Text = ""
        txtpo.Text = ""
        txtcomm.Text = ""
        txtBenefits.Text = ""
        txttoocost.Text = 0
        txtpartdescription.Text = ""
        txtvendorno.Text = ""
        txtvendorname.Text = ""
        txtctpno.Text = ""
        txtqty.Text = 0
        txtsampleqty.Text = 0

        txtmfrno.Text = ""
        txtunitcost.Text = 0
        txtminqty.Text = 0
        txtsample.Text = 0
        txttcost.Text = 0
        txtunitcostnew.Text = 0

        dgvProjectDetails.DataSource = Nothing
        DataGridView1.DataSource = Nothing

        optCTP.Checked = True
        optVENDOR.Checked = False
        optboth.Checked = False
        chknew.Checked = False

        DTPicker1.Value = Format(Now, "MM/dd/yyyy")
        'DTPicker5.Value = "01/01/1900"
        DTPicker2.Value = Format(Now, "MM/dd/yyyy")
        'DTPicker3.Value = "01/01/1900"
        'DTPicker4.Value = "01/01/1900"

        FillDDlUser()
        cmbuser1.SelectedIndex = 0

        FillDDlUser1()
        cmbuser.SelectedIndex = 0

        FillDDLStatus()
        cmbstatus.SelectedIndex = 0

        cmbminorcode.Items.Clear()

        'cmbminorcode.Clear
        'cmbprstatus.ListIndex = 0
        'cmbstatus.ListIndex = 0

        TabPage2.Text = ""

        flagdeve = 1
        flagnewpart = 1

    End Sub

    Private Function displayPart() As String
        Dim result As String = "-1"
        If optCTP.Checked = True Then
            result = "1"
        ElseIf optVENDOR.Checked = True Then
            result = "2"
        ElseIf optboth.Checked = True Then
            result = ""
        End If
        Return result
    End Function

    Private Function CustomStrWhereResult() As String
        'If flagallow = 1 Then
        strwhere = ""
        'Else
        'TEST QUERY
        'strwhere = "WHERE PRPECH = 'LREDONDO' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = 'LREDONDO') "
        'strwhere = "WHERE PRPECH = '" & userid & "' OR PRHCOD IN (SELECT PRHCOD FROM PRDVLD WHERE PRDUSR = '" & userid & "') "
        'strwhere = "WHERE PRPECH = '" & UserID & "'
        'End If
        Return strwhere
    End Function

    Private Sub cleanSearchTextBoxes(valueSelectd As String)
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel1

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    If tt.Name <> valueSelectd Then
                        tt.Text = ""
                    End If
                End If
            Next

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cleanSearchTextBoxesComplex(selValues As List(Of Object), includedControl As Boolean)
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel1

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    'If tt.Name <> valueSelectd Then
                    If Not includedControl Then
                        If Not isStringPresentInList(tt.Name, selValues) Then
                            tt.Text = ""
                        End If
                    Else
                        If isStringPresentInList(tt.Name, selValues) Then
                            tt.Text = ""
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cleanSearchTextBoxesComplexTab2(selValues As List(Of Object), includedControl As Boolean)
        Dim exMessage As String = " "
        Try
            Dim myTableLayout As TableLayoutPanel
            myTableLayout = Me.TableLayoutPanel15

            For Each tt In myTableLayout.Controls
                If TypeOf tt Is Windows.Forms.TextBox Or TypeOf tt Is Windows.Forms.ComboBox Then
                    'If tt.Name <> valueSelectd Then
                    If Not includedControl Then
                        If Not isStringPresentInList(tt.Name, selValues) Then
                            tt.Text = ""
                        End If
                    Else
                        If isStringPresentInList(tt.Name, selValues) Then
                            tt.Text = ""
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function isStringPresentInList(compareStr As String, selValues As List(Of Object)) As Boolean
        Dim exMessage As String = " "
        Try
            For Each item As Object In selValues
                If item IsNot Nothing Then
                    If item.Name = compareStr Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return False
        End Try
    End Function

    Sub ResizeTabs()
        Dim exMessage As String = Nothing
        Try
            Dim numTabs As Integer = SSTab1.TabCount

            Dim totLen As Decimal = 0
            Using g As Graphics = CreateGraphics()
                ' Get total length of the text of each Tab name
                For i As Integer = 0 To numTabs - 1
                    totLen += g.MeasureString(SSTab1.TabPages(i).Text, SSTab1.Font).Width
                Next
            End Using

            Dim newX As Integer = ((SSTab1.Width - totLen) / numTabs) / 2
            SSTab1.Padding = New Point(newX, SSTab1.Padding.Y)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub VScrollBar2_Scroll(sender As Object, e As ScrollEventArgs)

    End Sub

    Private Sub BindingNavigator1_RefreshItems(sender As Object, e As EventArgs)

    End Sub

    'Private Sub cmdSplit_Click(sender As Object, e As EventArgs) Handles cmdSplit.Click
    '    Dim screenPoint As Point = cmdSplit.PointToScreen(New Point(cmdSplit.Left, cmdSplit.Bottom))
    '    If screenPoint.Y & ContextMenuStrip1.Size.Height > Screen.PrimaryScreen.WorkingArea.Height Then
    '        ContextMenuStrip1.Show(cmdSplit, New Point(0, -ContextMenuStrip1.Size.Height))
    '    Else
    '        ContextMenuStrip1.Show(cmdSplit, New Point(0, cmdSplit.Height))
    '    End If
    '    'ContextMenuStrip1.Show(cmdSplit, New Point(0, cmdSplit.Height))
    'End Sub

#End Region

    'Protected Sub OnRowCommand(ByVal sender As Object, ByVal e As GridViewCommandEventArgs)
    'Dim index As Integer = Convert.ToInt32(e.CommandArgument)
    'Dim gvRow As DataGridViewRow = DataGridView1.Rows(index)
    'End Sub'

    'Private Sub dataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
    'For Each row As DataGridViewRow In DataGridView1.SelectedRows
    'Dim value11 As String = row.Cells(0).Value.ToString()
    'Next
    'End Sub

End Class

