Imports System.Globalization
Imports System.Reflection

Public Class frmproductsdevelopmentmanu

    Dim gnr As Gn1 = New Gn1()
    Public userid As String

    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

#Region "Action Methods"

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
    End Sub

    Private Sub frmproductsdevelopmentmanu_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Load()
    End Sub

    Private Sub Form_Load()
        Dim exMessage As String = " "
        Try
            If gnr.ConnSql.State = 1 Then
            Else
                gnr.ConnSql.ConnectionString = gnr.SQLCon
                gnr.ConnSql.Open()
            End If

            Dim codeproject = frmProductsDevelopment.txtCode.Text
            lblproject.Text = frmProductsDevelopment.txtCode.Text & " - " & Trim(frmProductsDevelopment.txtname.Text)

            'check delete temp
            fillcell2(codeproject)

            'userid = frmLogin.txtUserName.Text
            userid = LikeSession.userid

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - Massive Manufacturer Update", "")

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()
        End Try
    End Sub

    Private Sub fillcell2(code As String)
        Dim exMessage As String = " "
        Try

            Dim ds = gnr.GetInvProdDetailByProject(code)

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 3

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPartNo"
                    DataGridView1.Columns(0).HeaderText = "Part No."
                    DataGridView1.Columns(0).DataPropertyName = "PRDPTN"

                    DataGridView1.Columns(1).Name = "clDescription"
                    DataGridView1.Columns(1).HeaderText = "Descripcion"
                    DataGridView1.Columns(1).DataPropertyName = "IMDSC"

                    DataGridView1.Columns(2).Name = "clMfrNo"
                    DataGridView1.Columns(2).HeaderText = "Manufacture Part No."
                    DataGridView1.Columns(2).DataPropertyName = "PRDMFR#"

                    'FILL GRID
                    DataGridView1.DataSource = ds.Tables(0)

                    Dim headerCellLocation As Point = Me.DataGridView1.GetCellDisplayRectangle(0, -1, True).Location

                    'Place the Header CheckBox in the Location of the Header Cell.
                    Dim headerCheckBox As New CheckBox
                    headerCheckBox.Location = New Point(headerCellLocation.X + 8, headerCellLocation.Y + 2)
                    headerCheckBox.BackColor = Color.White
                    headerCheckBox.Size = New Size(18, 18)

                    'Assign Click event to the Header CheckBox.
                    AddHandler headerCheckBox.Click, AddressOf HeaderCheckBox_Clicked
                    DataGridView1.Controls.Add(headerCheckBox)

                    'Add a CheckBox Column to the DataGridView at the first position.
                    Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                    checkBoxColumn.HeaderText = "All"
                    checkBoxColumn.Width = 30
                    checkBoxColumn.Name = "checkBoxColumn"
                    DataGridView1.Columns.Insert(0, checkBoxColumn)

                    DataGridView1.Columns("clPartNo").ReadOnly = True
                    DataGridView1.Columns("clDescription").ReadOnly = True

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
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub HeaderCheckBox_Clicked(ByVal sender As Object, ByVal e As EventArgs)
        'Necessary to end the edit mode of the Cell.
        DataGridView1.EndEdit()

        'Loop and check and uncheck all row CheckBoxes based on Header Cell CheckBox.
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim checkBox As DataGridViewCheckBoxCell = (TryCast(row.Cells("checkBoxColumn"), DataGridViewCheckBoxCell))

            Dim myItem As CheckBox = CType(sender, CheckBox)
            checkBox.Value = myItem.Checked
        Next
    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click
        Dim exMessage As String = " "
        Dim ds As DataSet
        Dim updatedRecords As Integer = 0
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("checkBoxColumn").Value = True Then
                    ds = gnr.GetDataByCodeAndPartNoProdDesc(frmProductsDevelopment.txtCode.Text, row.Cells("clPartNo").Value.ToString())
                    If ds IsNot Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            Dim oldVendorNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("VMVNUM").Ordinal)
                            Dim oldMFROld = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDMFR#").Ordinal)
                            If Trim(UCase(oldMFROld)) <> Trim(UCase(row.Cells("clMfrNo").Value.ToString())) Then
                                PoQotaFunction(oldVendorNo, row.Cells("clPartNo").Value.ToString(), row.Cells("clMfrNo").Value.ToString())
                                Dim rsResultUpdate = gnr.UpdateChangedMFR(userid, row.Cells("clMfrNo").Value.ToString(), row.Cells("clPartNo").Value.ToString(), frmProductsDevelopment.txtCode.Text)
                                'update validation
                                updatedRecords += 1
                            End If
                        Else
                            'no data message
                        End If
                    Else
                        'no data message
                    End If

                End If
            Next

            If updatedRecords > 0 Then
                MessageBox.Show("Records Updated.", "CTP System", MessageBoxButtons.OK)
                Form_Load()
                frmProductsDevelopment.fillcell2(frmProductsDevelopment.txtCode.Text)
            Else
                MessageBox.Show("No records to update.", "CTP System", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub PoQotaFunction(oldVendorNo As String, partNo As String, newMFRNo As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(oldVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            Dim dsPoQota = gnr.GetPOQotaData(oldVendorNo, partNo)
            Dim auxpart1 = newMFRNo
            Dim auxpart2 = ""
            Dim wrtmessage = 0
            Dim MaxValue As Integer
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    auxpart2 = dsPoQota.Tables(0).Rows(0).ItemArray(dsPoQota.Tables(0).Columns("PQMPTN").Ordinal)
                    If auxpart1 <> Trim(UCase(auxpart2)) Then
                        wrtmessage = 1
                    End If
                    Dim rsUpdQota = gnr.UpdatePoQoraMfr(newMFRNo, oldVendorNo, partNo)
                    'validate update message
                Else
                    MaxValue = 1
                    wrtmessage = 1
                    MaxValue = If(Not String.IsNullOrEmpty(gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)), CInt(gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)) + 1, "")
                    spacepoqota = "                               DEV"
                    Dim rsInsertionPoqota = gnr.InsertNewPOQotaLess(partNo, oldVendorNo, MaxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), newMFRNo, DateTime.Now.Day.ToString(), "", spacepoqota, 0)
                    'validate insertion message
                End If

                If wrtmessage = 1 Then
                    Dim cod_comment = CInt(gnr.getmax("PRDCMH", "PRDCCO")) + 1

                    Dim rsInsertCommentHeader = gnr.InsertProductCommentNew(frmProductsDevelopment.txtCode.Text, partNo, cod_comment, "Manufacturer Number Changed", userid)
                    'validation message result
                    Dim cod_detcomment = 1
                    Dim messcomm = "Part Number changed from " & auxpart2 & " to " & auxpart1
                    Dim rsInsertCommetnDetail = gnr.InsertProductCommentDetail(frmProductsDevelopment.txtCode.Text, partNo, cod_comment, cod_detcomment, messcomm)
                    'validation message result
                End If
            Else
                'error message
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function PoQotaFunctionDuplex(newMFRNo As String, partNo As String, seqNo As String) As Integer
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim strQueryAdd As String = " WHERE PRDMFR# = " & Trim(newMFRNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND PQSEQ = '" & Trim(seqNo) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            'Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaDataDuplex(strQueryAdd)
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    Return 0
                    'validation result
                Else
                    Return -1
                    'error message
                End If
            Else
                Return -1
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return -1
        End Try
    End Function

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

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