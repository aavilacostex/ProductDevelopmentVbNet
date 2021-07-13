Imports System.Globalization
Imports System.Reflection

Public Class frmPDevelopmentseecomments

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

#Region "Action Methods"

    Private Sub frmPDevelopmentseecomments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Load()
    End Sub

    Private Sub Form_Load()
        Dim exMessage As String = " "
        cmddelete.Enabled = False
        Try

            'userid = frmLogin.txtUserName.Text
            userid = LikeSession.userid

            If UCase(userid) = "AALZATE" Then
                flagallow = 1
            End If

            Dim rsDeletionSql = gnr.DeleteDataSqlByUser("tbtemppdevelopmentseecomment", userid)
            Dim dsSelection = gnr.GetDataSqlByUser("tbtemppdevelopmentseecomment", userid)

            gnr.seeaddprocomments = lblNotVisible.Text
            If gnr.seeaddprocomments = 5 Then
                txtCode.Text = frmProductsDevelopment.txtCode.Text
                txtpartno.Text = Trim(UCase(frmProductsDevelopment.txtpartno.Text))
            End If

            Dim dsProdHeaderMess = gnr.GetDataByCodAndPartProdAndComm1(Trim(txtCode.Text), Trim(txtpartno.Text))
            fillDgvProjectMessages(dsProdHeaderMess)

            TabPage1.Text = ""

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - PD See Comments", "")

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()
        End Try
    End Sub

    Private Sub SSTab1_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) _
    Handles SSTab1.Selected
        If SSTab1.SelectedTab.Name = "TabPage1" Then
            cmddelete.Enabled = False
            dgvProjectMessage2.DataSource = Nothing
            dgvProjectMessage2.Refresh()
            dgvProjectMessage2.AutoGenerateColumns = False
            Form_Load()
            'dgvProjectMessage2.ColumnCount = 1
        Else
            cmddelete.Enabled = True
        End If
    End Sub

    Public Sub fillDgvProjectMessages(dsData As DataSet)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            If dsData.Tables(0).Rows.Count > 0 Then
                dgvProjectMessages.DataSource = Nothing
                dgvProjectMessages.Refresh()
                dgvProjectMessages.AutoGenerateColumns = False

                'Add Columns
                dgvProjectMessages.Columns(0).Name = "clSubject"
                dgvProjectMessages.Columns(0).HeaderText = "Subject"
                dgvProjectMessages.Columns(0).DataPropertyName = "PRDCSU"

                dgvProjectMessages.Columns(1).Name = "clDateEntered"
                dgvProjectMessages.Columns(1).HeaderText = "Date Entered"
                dgvProjectMessages.Columns(1).DataPropertyName = "PRDCDA"

                dgvProjectMessages.Columns(2).Name = "clTimeEntered"
                dgvProjectMessages.Columns(2).HeaderText = "Time Entered"
                dgvProjectMessages.Columns(2).DataPropertyName = "PRDCTI"

                dgvProjectMessages.Columns(3).Name = "clUser"
                dgvProjectMessages.Columns(3).HeaderText = "User"
                dgvProjectMessages.Columns(3).DataPropertyName = "USUSER"

                dgvProjectMessages.Columns(4).Name = "clTableCode"
                dgvProjectMessages.Columns(4).HeaderText = "Table Code"
                dgvProjectMessages.Columns(4).DataPropertyName = "PRDCCO"

                dgvProjectMessages.DataSource = dsData.Tables(0)
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Public Sub fillDgvProjectMessage2(dsData As DataSet)
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            If dsData.Tables(0).Rows.Count > 0 Then
                dgvProjectMessage2.Columns.Clear()
                dgvProjectMessage2.DataSource = Nothing
                dgvProjectMessage2.Refresh()
                dgvProjectMessage2.AutoGenerateColumns = False
                dgvProjectMessage2.ColumnCount = 2

                'Add Columns
                dgvProjectMessage2.Columns(1).Name = "clTableCode"
                dgvProjectMessage2.Columns(1).HeaderText = "Table Code"
                dgvProjectMessage2.Columns(1).DataPropertyName = "PRDCCO"

                'dgvProjectMessage2.Columns(2).Name = "clCommentNo"
                'dgvProjectMessage2.Columns(2).HeaderText = "Comment No"
                'dgvProjectMessage2.Columns(2).DataPropertyName = "PRDCDC"

                dgvProjectMessage2.Columns(0).Name = "clComments"
                dgvProjectMessage2.Columns(0).HeaderText = "Comments"
                dgvProjectMessage2.Columns(0).DataPropertyName = "PRDCTX"

                dgvProjectMessage2.DataSource = dsData.Tables(0)

                'Find the Location of Header Cell.
                Dim headerCellLocation As Point = Me.dgvProjectMessage2.GetCellDisplayRectangle(0, -1, True).Location

                'Place the Header CheckBox in the Location of the Header Cell.
                Dim headerCheckBox As New CheckBox
                headerCheckBox.Location = New Point(headerCellLocation.X + 8, headerCellLocation.Y + 2)
                headerCheckBox.BackColor = Color.White
                headerCheckBox.Size = New Size(18, 18)

                'Assign Click event to the Header CheckBox.
                AddHandler headerCheckBox.Click, AddressOf HeaderCheckBox_Clicked
                dgvProjectMessage2.Controls.Add(headerCheckBox)

                'Add a CheckBox Column to the DataGridView at the first position.
                Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                checkBoxColumn.HeaderText = "All"
                checkBoxColumn.Width = 30
                checkBoxColumn.Name = "checkBoxColumn"
                dgvProjectMessage2.Columns.Insert(0, checkBoxColumn)

                Me.dgvProjectMessage2.Columns("clTableCode").Visible = False

                'Assign Click event to the DataGridView Cell.
                'AddHandler dgvProjectMessage2.CellContentClick, AddressOf DataGridView_CellClick
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub dgvProjectMessages_DoubleClick(ByVal sender As Object, ByVal e As EventArgs) Handles dgvProjectMessages.DoubleClick
        Dim exMessage As String = " "
        Dim Index As Integer
        Dim dsResult As DataSet
        Try
            Dim rsDeletionSql = gnr.DeleteDataSqlByUser("tbtemppdevelopmentseecomment", userid)
            Dim dsSelection = gnr.GetDataSqlByUser("tbtemppdevelopmentseecomment", userid)

            For Each row As DataGridViewRow In dgvProjectMessages.SelectedRows
                Index = dgvProjectMessages.CurrentCell.RowIndex

                If dgvProjectMessages.Rows(Index).Selected = True Then
                    Dim tableCode As String = row.Cells(4).Value.ToString()
                    hdField.Text = tableCode
                    dsResult = gnr.GetDataByCodAndPartProdAndComm2(tableCode)
                    If dsResult IsNot Nothing Then
                        If dsResult.Tables(0).Rows.Count > 0 Then
                            fillDgvProjectMessage2(dsResult)
                            SSTab1.SelectedIndex = 1
                        Else
                            Dim rsMessage As DialogResult = MessageBox.Show("There is not a detail message for this header message!", "CTP System", MessageBoxButtons.OK)
                        End If
                    Else
                        Dim rsMessage As DialogResult = MessageBox.Show("There is not a detail message for this header message!", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
            Next

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub HeaderCheckBox_Clicked(ByVal sender As Object, ByVal e As EventArgs)
        'Necessary to end the edit mode of the Cell.
        dgvProjectMessage2.EndEdit()

        'Loop and check and uncheck all row CheckBoxes based on Header Cell CheckBox.
        For Each row As DataGridViewRow In dgvProjectMessage2.Rows
            Dim checkBox As DataGridViewCheckBoxCell = (TryCast(row.Cells("checkBoxColumn"), DataGridViewCheckBoxCell))

            Dim myItem As CheckBox = CType(sender, CheckBox)
            checkBox.Value = myItem.Checked
        Next
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    'Private Sub DataGridView_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
    '    'Check to ensure that the row CheckBox is clicked.
    '    If e.RowIndex >= 0 AndAlso e.ColumnIndex = 0 Then

    '        'Loop to verify whether all row CheckBoxes are checked or not.
    '        Dim isChecked As Boolean = True
    '        For Each row As DataGridViewRow In dgvProjectMessage2.Rows
    '            If Convert.ToBoolean(row.Cells("checkBoxColumn").EditedFormattedValue) = False Then
    '                isChecked = False
    '                Exit For
    '            End If
    '        Next
    '        Dim myItem() As Control = Me.Controls.Find("Checkbox", True)
    '        If TypeOf myItem(0) Is CheckBox Then
    '            Dim chk As CheckBox = DirectCast(myItem(0), CheckBox)
    '            If chk.Name = "headerCheckBox" Then
    '                chk.Checked = isChecked
    '            End If
    '        End If
    '        'headerCheckBox.Checked = isChecked
    '    End If
    'End Sub

    Private Sub cmddelete_Click(sender As Object, e As EventArgs) Handles cmddelete.Click
        Dim exMessage As String = " "
        Dim codeList As New List(Of String)
        Dim strValues As String = ""
        Try
            For Each oRow As DataGridViewRow In dgvProjectMessage2.Rows
                If oRow.Cells("checkBoxColumn").Value = True Then
                    If Not codeList.Contains(oRow.Cells("clTableCode").Value.ToString()) Then
                        codeList.Add(oRow.Cells("clTableCode").Value.ToString())
                    End If
                End If
            Next

            For Each tt As String In codeList
                If tt IsNot codeList.Last Then
                    strValues += tt + ","
                Else
                    strValues += tt
                End If
            Next

            Dim rsDeletion = gnr.DeleteGeneral("PRDCMD", "PRDCCO", strValues)
            'Dim rsDeletion = 1
            If rsDeletion = 1 Then
                Dim ds As New DataSet
                Dim dt As New DataTable
                dt = (DirectCast(dgvProjectMessage2.DataSource, DataTable))

                Dim rows As DataRow() = (From anyNamehere In dt.AsEnumerable().Cast(Of DataRow)() Where anyNamehere.ItemArray(0).ToString().Contains(strValues)).ToArray()
                If rows.Count > 0 Then
                    For Each row As DataRow In rows
                        dt.Rows.Remove(row)
                    Next

                    'ds.Tables.Add(dt)
                    If dt IsNot Nothing Then
                        If dt.Rows.Count > 0 Then
                            dgvProjectMessage2.DataSource = dt
                            dgvProjectMessage2.Refresh()
                        Else
                            dgvProjectMessage2.DataSource = Nothing
                            dgvProjectMessage2.Refresh()
                            SSTab1.SelectedIndex = 0
                        End If
                    Else
                        dgvProjectMessage2.DataSource = Nothing
                        dgvProjectMessage2.Refresh()
                        SSTab1.SelectedIndex = 0
                    End If
                Else
                    dgvProjectMessage2.DataSource = Nothing
                    dgvProjectMessage2.Refresh()
                    SSTab1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub cmdprint_Click(sender As Object, e As EventArgs) Handles cmdprint.Click
        Try
            Dim code = frmProductsDevelopment.txtCode.Text
            Dim partNo = Trim(UCase(frmProductsDevelopment.txtpartno.Text))

            Dim dsPrintData = gnr.GetDataByCodAndPartProdAndComm2(code, partNo)
            If dsPrintData IsNot Nothing Then
                If dsPrintData.Tables(0).Rows.Count > 0 Then
                    Dim sql = "SELECT PRDCMH.*,PRDCMD.* FROM PRDCMH INNER JOIN PRDCMD ON PRDCMH.PRDCCO = PRDCMD.PRDCCO WHERE PRDCMH.PRHCOD = " & Trim(code) & " AND PRDCMH.PRDPTN = '" & Trim(partNo) & "' ORDER BY  PRDCDA ASC,PRDCTI ASC"
                    AxCrystalReport1.SQLQuery = sql
                    AxCrystalReport1.ReportFileName = Trim(gnr.ReportsValue) & "\projectcomments.rpt"
                    AxCrystalReport1.Connect = gnr.CrystalCon
                    AxCrystalReport1.RetrieveDataFiles()
                    AxCrystalReport1.PrintReport()
                    AxCrystalReport1.WindowState = FormWindowState.Maximized
                End If
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
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