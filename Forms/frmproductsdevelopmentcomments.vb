Imports System.Globalization
Imports System.Reflection

Public Class frmproductsdevelopmentcomments

    Dim gnr As Gn1 = New Gn1()
    Dim sql As String
    Public userid As String
    Public flagallow As Integer
    Public cod_detcomment As Integer
    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

    Private Sub frmproductsdevelopmentcomments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim exMessage As String = " "
        Try

            'userid = frmLogin.txtUserName.Text
            userid = LikeSession.userid

            'AddHandler dgvAddComments.CellClick, AddressOf Me.dgvAddComments_CellClick

            'test purpose
            If UCase(userid) = "AALZATE" Then
                flagallow = 1
            End If

            'Dim rsDeletionSql = gnr.DeleteDataSqlByUser("tbtempproductcomment", userid)
            'If rsDeletionSql >= 0 Then
            '    Dim dsSelection = gnr.GetDataSqlByUser("tbtempproductcomment", userid)
            '    fillDgvAddComments(dsSelection)
            'Else
            '    'error message
            'End If


            gnr.seeaddprocomments = lblNotVisible.Text
            If gnr.seeaddprocomments = 5 Then
                txtCode.Text = frmProductsDevelopment.txtCode.Text
                txtpartno.Text = Trim(UCase(frmProductsDevelopment.txtpartno.Text))

                DTPicker1.Value = Format(Now, "MM/dd/yyyy")
                DTPicker2.Value = Format(Now, "MM/dd/yyyy")
                txtsubject.Text = ""
            End If

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - Add Comments Start", "")

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()
        End Try

    End Sub

#Region "Grid Views"

    Private Sub dgvAddComments_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAddComments.CellClick
        If e.ColumnIndex = -1 Then
            dgvAddComments.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            dgvAddComments.EndEdit()
        ElseIf dgvAddComments.EditMode <> DataGridViewEditMode.EditOnEnter Then
            dgvAddComments.EditMode = DataGridViewEditMode.EditOnEnter
            dgvAddComments.BeginEdit(False)
        End If
    End Sub

    Public Sub fillDgvAddComments(dsData As DataSet)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            If dsData.Tables(0).Rows.Count > 0 Then
                dgvAddComments.DataSource = Nothing
                dgvAddComments.Refresh()
                dgvAddComments.AutoGenerateColumns = False

                'Add Columns
                dgvAddComments.Columns(0).Name = "clComments"
                dgvAddComments.Columns(0).HeaderText = "Comments"
                dgvAddComments.Columns(0).DataPropertyName = "comment"

                dgvAddComments.DataSource = dsData.Tables(0)
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub fillCellComments()
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            'ds = gnr.FillGridSql(sql)
            ds = Nothing
            If ds IsNot Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then

                    'DataGridView1.DataSource = Nothing
                    'DataGridView1.Refresh()
                    'DataGridView1.AutoGenerateColumns = False
                    'DataGridView1.ColumnCount = 5

                    'Add Columns
                    'DataGridView1.Columns(0).Name = "ProjectNo"
                    'DataGridView1.Columns(0).HeaderText = "Project No."
                    'DataGridView1.Columns(0).DataPropertyName = "PRHCOD"

                    'DataGridView1.Columns(1).Name = "ProjectName"
                    'DataGridView1.Columns(1).HeaderText = "Project Name"
                    'DataGridView1.Columns(1).DataPropertyName = "PRNAME"

                    'DataGridView1.Columns(2).Name = "DateEnt"
                    'DataGridView1.Columns(2).HeaderText = "Date Entered"
                    'DataGridView1.Columns(2).DataPropertyName = "PRDATE"

                    'DataGridView1.Columns(3).Name = "PersonInCharge"
                    'DataGridView1.Columns(3).HeaderText = "Person In Charge"
                    'DataGridView1.Columns(3).DataPropertyName = "PRPECH"

                    'DataGridView1.Columns(4).Name = "Status"
                    'DataGridView1.Columns(4).HeaderText = "Status"
                    'DataGridView1.Columns(4).DataPropertyName = "PRSTAT"

                    'FILL GRID
                    'DataGridView1.DataSource = ds.Tables(0)

                Else
                    'DataGridView1.DataSource = Nothing
                    'DataGridView1.Refresh()
                    'If flag = 0 Then
                    '    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                    'End If
                    'Exit Sub
                End If
            Else
                'DataGridView1.DataSource = Nothing
                'DataGridView1.Refresh()
                'If flag = 0 Then
                '    Dim resultAlert As DialogResult = MessageBox.Show("There is not results for this search criteria. Please try again with other text!", "CTP System", MessageBoxButtons.OK)
                'End If
                'Exit Sub
            End If
        Catch ex As Exception
            'DataGridView1.DataSource = Nothing
            'DataGridView1.Refresh()
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

#End Region

#Region "Buttons"

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdnew_Click(sender As Object, e As EventArgs) Handles cmdnew.Click
        Dim exMessage As String = " "
        Try
            cmdSave.Enabled = True
            DTPicker1.Value = Format(Now, "mm/dd/yyyy")
            DTPicker2.Value = Format(Now, "hh:mm:ss")
            txtsubject.Text = ""

            Dim rsDeletion = gnr.DeleteDataSqlByUser("tbtempproductcomment", userid)
            If rsDeletion = -1 Then
                'error
            Else
                Dim dsSelection = gnr.GetDataSqlByUser("tbtempproductcomment", userid)
                If dsSelection IsNot Nothing Then
                    fillDgvAddComments(dsSelection)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Dim exMessage As String = " "
        Dim strMessage As String
        Dim rsCommentDetail As Integer
        Dim rsCommentHeader As Integer
        Dim maxValue As Object
        Dim lstSqlMessDet As New List(Of String)
        Dim sqlInsert As Object
        Try
            If Trim(txtCode.Text) <> "" Then
                If txtsubject.Text <> "" Then


                    maxValue = gnr.getmax("PRDCMH", "PRDCCO")
                    maxValue += 1
                    rsCommentHeader = gnr.InsertProductComment(txtCode.Text, txtpartno.Text, maxValue, userid)
                    If rsCommentHeader >= 0 Then
                        cod_detcomment = 1
                        For Each row As DataGridViewRow In dgvAddComments.Rows
                            strMessage = Trim(row.Cells("clComments").Value)
                            If Not String.IsNullOrEmpty(strMessage) Then
                                sqlInsert = saveSqlComments(strMessage)
                                If sqlInsert > 0 Then
                                    rsCommentDetail = gnr.InsertProductCommentDetail(txtCode.Text, txtpartno.Text, maxValue, cod_detcomment, strMessage)
                                    cod_detcomment = cod_detcomment + 1
                                Else
                                    'error insertando en sql server
                                End If
                            End If
                        Next
                    Else
                        'error insertando en cabecera de detalle de producto
                    End If
                    cmdSave.Enabled = False
                    Dim msgInsertOk As DialogResult = MessageBox.Show("Comment(s) Saved.", "CTP System", MessageBoxButtons.OK)
                Else
                    Dim msgNotSubject As DialogResult = MessageBox.Show("Enter Subject.", "CTP System", MessageBoxButtons.OK)
                End If
            Else
                Dim msgNotCode As DialogResult = MessageBox.Show("Enter Comments.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub


#End Region

#Region "Utis"

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

    Private Function saveSqlComments(strComment As String) As Integer
        Dim rsInsertion As Integer = -1
        Dim cod_detcomment As Integer
        Dim lstSqlMessDet As New List(Of String)
        Dim exMessage As String = " "
        Try
            Dim dsMaxValue = gnr.GetMaxCodeDetSql("tbtempproductcomment", "cod_comment")
            If dsMaxValue IsNot Nothing Then
                If dsMaxValue.Tables(0).Rows.Count > 0 Then
                    If Not String.IsNullOrEmpty(dsMaxValue.Tables(0).Rows(0).ItemArray(0).ToString()) Then
                        If CInt(dsMaxValue.Tables(0).Rows(0).ItemArray(0).ToString()) > 0 Then
                            Dim tempValue = CInt(dsMaxValue.Tables(0).Rows(0).ItemArray(0).ToString())
                            cod_detcomment = tempValue + 1
                        Else
                            cod_detcomment = 1
                        End If
                    End If
                    lstSqlMessDet.Add(cod_detcomment)
                    lstSqlMessDet.Add(strComment)
                    rsInsertion = gnr.InsertDataSqlByUser("tbtempproductcomment", userid, lstSqlMessDet)
                    If rsInsertion < 0 Then
                        'error message
                        Return rsInsertion
                    Else
                        Return rsInsertion
                    End If
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsInsertion
        End Try
    End Function

#End Region

End Class