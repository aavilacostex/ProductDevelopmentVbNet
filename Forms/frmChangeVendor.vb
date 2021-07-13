Imports System.Reflection

Public Class frmChangeVendor
    Dim gnr As Gn1 = New Gn1()
    Public userid As String
    Public flagchangevendor As Integer

    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

#Region "Action Methods"

    Private Sub frmChangeVendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmChangeVendor_Load()
    End Sub

    Private Sub frmChangeVendor_Load()
        Dim exMessage As String = " "
        Try
            txtCode.Text = ""
            txtsearch1.Text = ""
            flagchangevendor = 1
            cmbvendor.Items.Clear()

            userid = LikeSession.userid

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - Change Vendor Value", "")

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()

            MessageBox.Show(exMessage, "CTP System", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub cmdSearch_Click()
        Dim strfield As String
        Dim dfield
        Dim exMessage As String = " "
        Try
            txtsearch1.Text = ""
            cmbvendor.DataSource = Nothing
            dfield = Trim(txtCode.Text)
            If Len(Trim(txtCode.Text)) > 0 Then
                cmbvendor.Items.Clear()
                For i = 1 To Len(Trim(dfield))
                    strfield = Mid(dfield, i, 1)
                    If strfield Like "[!0-9]" Then
                        MsgBox("Just numbers", vbOKOnly + vbInformation, "CTP System")
                        i = Len(Trim(dfield)) + 1
                        Exit Sub
                    End If
                Next i
                If Trim(txtCode.Text) > 0 Then
                    Dim dsResult = gnr.GetVendorByVendorNo(txtCode.Text)

                    If dsResult IsNot Nothing Then
                        If dsResult.Tables(0).Rows.Count > 0 Then
                            cmbvendor.DataSource = dsResult.Tables(0)
                            cmbvendor.DisplayMember = "VMNAME"
                            cmbvendor.ValueMember = "VMVNUM"
                            If cmbvendor.Items.Count > 0 Then
                                cmbvendor.SelectedIndex = 0
                            End If
                        Else
                            cmbvendor.Items.Clear()
                            MsgBox("Vendor(s) not found.", vbOKOnly + vbInformation, "CTP System")
                            txtCode.Text = 0
                        End If
                    End If
                Else
                    If Trim(txtCode.Text) = 0 Then
                        cmbvendor.Items.Clear()
                    End If
                End If
            Else
                If Trim(txtCode.Text) = "" Then
                    txtCode.Text = 0
                    Exit Sub
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdsearch1_Click()
        Dim exMessage As String = " "
        Try
            cmbvendor.Items.Clear()
            txtCode.Text = ""
            If Trim(txtsearch1.Text) <> "" Then

                'Dim Sql = "SELECT * FROM VNMAS WHERE TRIM(UCASE(VMNAME)) LIKE '" & Trim(UCase(txtsearch1.Text)) & "%'"
                Dim dsResult = gnr.GetVendorByName(txtsearch1.Text)

                If dsResult IsNot Nothing Then
                    If dsResult.Tables(0).Rows.Count > 0 Then
                        cmbvendor.DataSource = dsResult.Tables(0)
                        cmbvendor.DisplayMember = "VMNAME"
                        cmbvendor.ValueMember = "VMVNUM"
                        If cmbvendor.Items.Count > 0 Then
                            cmbvendor.SelectedIndex = 0
                        End If
                    Else
                        cmbvendor.Items.Clear()
                        MsgBox("Vendor(s) not found.", vbOKOnly + vbInformation, "CTP System")
                        txtCode.Text = 0
                    End If
                End If
            Else
                If Trim(txtCode.Text) = 0 Then
                    cmbvendor.Items.Clear()
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub cmdchange_Click()

        If Trim(txtCode.Text) <> "" Then
            If flagchangevendor = 1 Then
                Call gnr.changeVendor(frmProductsDevelopment.txtpartno.Text, txtCode.Text, userid)
            End If
            If flagchangevendor = 2 Then
                'Call gnr.changeVendor(frmproductsdevelopmentTS.txtpartno.Text, txtCode.Text, userid)
            End If
            If flagchangevendor = 3 Then
                ' Call gnr.changeVendor(frmproductsdevelopmentpur.txtpartno.Text, txtCode.Text, userid)
            End If
            MsgBox("Vendor Changed.", vbOKOnly + vbInformation, "CTP System")
        End If
        Exit Sub
    End Sub

    Private Sub cmdexit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdchange_Click(sender As Object, e As EventArgs) Handles cmdchange.Click
        cmdchange_Click()
    End Sub

    Private Sub cmdsearch_Click(sender As Object, e As EventArgs) Handles cmdsearch.Click
        cmdSearch_Click()
    End Sub

    Private Sub cmdsearch1_Click(sender As Object, e As EventArgs) Handles cmdsearch1.Click
        cmdsearch1_Click()
    End Sub

#End Region

#Region "Utils"

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

    Public Shared Function GetComputerName() As String
        Dim exMessage As String = Nothing
        Try
            Dim ComputerName As String
            ComputerName = Environment.MachineName
            Return ComputerName
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Sub writeLog(strLogCadenaCabecera As String, strLevel As VBLog.ErrorTypeEnum, strMessage As String, strDetails As String)
        strLogCadena = strLogCadenaCabecera + " " + System.Reflection.MethodBase.GetCurrentMethod().ToString()

        vblog.WriteLog(strLevel, "CTPSystem" & strLevel, strLogCadena, userid, strMessage, strDetails)
    End Sub

#End Region


End Class