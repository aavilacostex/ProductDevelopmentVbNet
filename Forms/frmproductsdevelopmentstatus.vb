Imports System.Reflection

Public Class frmproductsdevelopmentstatus

    Dim gnr As Gn1 = New Gn1()
    Public userid As String
    Dim toemails As String = ""

    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

#Region "Action Methods"

    Private Sub frmproductsdevelopmentstatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

            FillDDLStatus()

            Dim codeproject = frmProductsDevelopment.txtCode.Text
            lblproject.Text = frmProductsDevelopment.txtCode.Text & " - " & Trim(frmProductsDevelopment.txtname.Text)

            'check delete temp


            Dim dsInvPrdoDet = gnr.GetInvProdDetailByProject(codeproject)
            fillcell2(dsInvPrdoDet)

            'userid = frmLogin.txtUserName.Text
            userid = LikeSession.userid

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - Massive status start", "")

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()
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

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) _
    Handles DataGridView1.CellFormatting
        Dim CurrentState As String = ""
        If e.ColumnIndex = 5 Then
            If e.Value IsNot Nothing Then
                CurrentState = e.Value.ToString
                If CurrentState = "A " Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Approved"
                ElseIf CurrentState = "R " Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Rejected"
                ElseIf CurrentState = "NS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Negotiation with Supplier"
                ElseIf CurrentState = "RP" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Receiving of First Production"
                ElseIf CurrentState = "CS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed Successfully"
                ElseIf CurrentState = "CN" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed-Approved w/o Negotiation"
                ElseIf CurrentState = "CD" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed-Rejected"
                ElseIf CurrentState = "CL" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Closed w/o negotiation"
                ElseIf CurrentState = "AA" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Approved with advice"
                ElseIf CurrentState = "Q " Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Quoting"
                ElseIf CurrentState = "TD" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Technical Documentation"
                ElseIf CurrentState = "DP" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Documentation in Process"
                ElseIf CurrentState = "DF" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Documentation Finalized"
                ElseIf CurrentState = "SS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Sample already Sent"
                ElseIf CurrentState = "PS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Pending from Supplier"
                ElseIf CurrentState = "AS" Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Analysis of Samples"
                ElseIf CurrentState = "E " Then
                    DataGridView1.Rows(e.RowIndex).Cells("clStatus").Value = "Entered"
                End If
            End If
        End If
    End Sub

    Private Sub fillcell2(ds As DataSet)
        Dim exMessage As String = " "
        Try

            If Not ds Is Nothing Then

                If ds.Tables(0).Rows.Count > 0 Then
                    DataGridView1.Columns.Clear()
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Refresh()
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.ColumnCount = 5

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPartNo"
                    DataGridView1.Columns(0).HeaderText = "Part No."
                    DataGridView1.Columns(0).DataPropertyName = "PRDPTN"

                    DataGridView1.Columns(1).Name = "clCtpNo"
                    DataGridView1.Columns(1).HeaderText = "CTP No."
                    DataGridView1.Columns(1).DataPropertyName = "PRDCTP"

                    DataGridView1.Columns(2).Name = "clMfrNo"
                    DataGridView1.Columns(2).HeaderText = "MFR No."
                    DataGridView1.Columns(2).DataPropertyName = "PRDMFR#"

                    DataGridView1.Columns(3).Name = "clDescription"
                    DataGridView1.Columns(3).HeaderText = "Descripcion"
                    DataGridView1.Columns(3).DataPropertyName = "IMDSC"

                    DataGridView1.Columns(4).Name = "clStatus"
                    DataGridView1.Columns(4).HeaderText = "Status"
                    DataGridView1.Columns(4).DataPropertyName = "PRDSTS"

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
                    DataGridView1.Columns("clMfrNo").ReadOnly = True
                    DataGridView1.Columns("clCtpNo").ReadOnly = True

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

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
    End Sub

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click
        Dim exMessage As String = " "
        Dim ds As DataSet
        Dim flagustatus As Integer
        Dim oldStatus As String
        Dim unitCostVendor As String
        Dim status1 As String = ""
        Dim status2 As String = ""
        Dim messcomm As String
        Dim updatedRecords As Integer = 0
        Dim intMassVendNoFlg As Integer = 0
        Dim allpartno As String = ""
        Dim partNo As String
        Try
            Dim rsMessage As DialogResult = MessageBox.Show("Do you want to change the assigned vendor for these part number?", "CTP System", MessageBoxButtons.YesNo)
            If rsMessage = DialogResult.Yes Then
                intMassVendNoFlg = 1
            End If

            Dim cmbStatusSelection = Trim(cmbstatus.Text.Substring(0, 2))

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("checkBoxColumn").Value = True Then
                    partNo = Trim(row.Cells("clPartNo").Value.ToString())
                    ds = gnr.GetDataByCodeAndPartNoProdDesc(frmProductsDevelopment.txtCode.Text, partNo)
                    If ds IsNot Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            flagustatus = 1
                            oldStatus = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDSTS").Ordinal)
                            unitCostVendor = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("PRDCON").Ordinal)
                            If Trim(oldStatus) = "AS" And cmbStatusSelection <> "AS" Then
                                If Trim(cmbStatusSelection) = "R" Or cmbStatusSelection = "A" Or cmbStatusSelection = "AA" Then
                                    flagustatus = 1
                                Else
                                    flagustatus = 0
                                End If
                            Else
                                flagustatus = 1
                            End If

                            If Trim(cmbStatusSelection) <> Trim(oldStatus) Then
                                If flagustatus = 1 Then
                                    Dim cod_comment = CInt(gnr.getmax("PRDCMH", "PRDCCO")) + 1
                                    Dim rsInsertCommentHeader = gnr.InsertProductCommentNew(frmProductsDevelopment.txtCode.Text, partNo, cod_comment, "Status changed", userid)
                                    'validation message result
                                    Dim cod_detcomment = 1
                                    status1 = ""
                                    status1 = If(Not String.IsNullOrEmpty(gnr.GetProjectStatusDescription(oldStatus)), Trim(gnr.GetProjectStatusDescription(oldStatus)), "")
                                    status2 = ""
                                    status2 = If(Not String.IsNullOrEmpty(gnr.GetProjectStatusDescription(Trim(cmbStatusSelection))), Trim(gnr.GetProjectStatusDescription(Trim(cmbStatusSelection))), "")

                                    messcomm = "Status changed from " & status1 & " to " & status2
                                    Dim rsResult = gnr.InsertProductCommentDetail(frmProductsDevelopment.txtCode.Text, partNo, cod_comment, cod_detcomment, messcomm)
                                    'validate error message update

                                End If
                            End If

                            Dim oldVendorNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("VMVNUM").Ordinal)
                            PoQotaFunction(oldVendorNo, partNo, Trim(cmbstatus.GetItemText(cmbstatus.SelectedItem(1))))

                            If (Trim(status2) = "Approved") Or (Trim(status2) = "Approved with advice") Then
                                If intMassVendNoFlg = 1 Then
                                    Dim dsDataPart = gnr.GetDataByPartVendor(partNo)
                                    If dsDataPart IsNot Nothing Then
                                        If dsDataPart.Tables(0).Rows.Count > 0 Then
                                            'call changeVendor
                                        End If
                                    Else
                                        Dim dsNoDataPart As DataSet
                                        dsNoDataPart = gnr.GetDataByPartNoVendor(partNo)
                                        If dsNoDataPart IsNot Nothing Then
                                            If dsNoDataPart.Tables(0).Rows.Count > 0 Then
                                                Dim impc1 = dsNoDataPart.Tables(0).Rows(0).ItemArray(dsNoDataPart.Tables(0).Columns("impc1").Ordinal)
                                                Dim impc2 = dsNoDataPart.Tables(0).Rows(0).ItemArray(dsNoDataPart.Tables(0).Columns("impc2").Ordinal)
                                                Dim decimalDate = getDataAsDecimal()
                                                Dim newChar = oldVendorNo.ToString().Insert(0, "000000")
                                                Dim lengthVendor = newChar.Length()
                                                Dim value1 = lengthVendor - 6
                                                Dim vendor = newChar.Substring(value1, lengthVendor - value1)
                                                'check dvinva insertion
                                                Dim rsInsertResult = gnr.InsertNewInv("01", partNo, impc1, impc2, decimalDate, unitCostVendor, "99999", "99999", vendor)
                                                'validation insert message
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            'no data message
                        End If
                    Else
                        'no data message
                    End If

                    If flagustatus = 1 Then
                        Dim rsResultUpdate = gnr.UpdateChangedStatus(userid, Trim(cmbStatusSelection), partNo, frmProductsDevelopment.txtCode.Text)
                        'validation update result message
                        updatedRecords += 1
                        allpartno = allpartno & " - " & UCase(partNo)
                    End If
                End If
            Next
            If (Trim(status2) = "Technical Documentation") Or (Trim(status2) = "Analysis of Samples") Or (Trim(status2) = "Pending from Supplier") Then
                'send email
            End If
            If Trim(status2) = "Closed Successfully" Then
                toemails = prepareEmailsToSend(1)
                Dim rsResult = gnr.sendEmail("", UCase(partNo))
                If rsResult < 0 Then
                    'mensaje de error
                End If
            End If

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

    Private Sub PoQotaFunction(oldVendorNo As String, partNo As String, status As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = status
        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(oldVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            Dim dsPoQota = gnr.GetPOQotaData(oldVendorNo, partNo)
            Dim MaxValue As Integer
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    Dim rsUpdResult = gnr.UpdatePoQoraRowNew(statusquote, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(), oldVendorNo, partNo)
                Else
                    MaxValue = 1
                    MaxValue = If(Not String.IsNullOrEmpty(gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)), CInt(gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)) + 1, "")
                    spacepoqota = "                               DEV"
                    Dim rsInsertionPoqota = gnr.InsertNewPOQota1(partNo, oldVendorNo, MaxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), "", DateTime.Now.Day.ToString(), Trim(statusquote), spacepoqota)
                    'validate insertion message
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

    Private Sub cmdSelectAll_Click(sender As Object, e As EventArgs) Handles Button1.Click

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
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function getDataAsDecimal() As Decimal
        Dim exMessage As String = " "
        Dim days As Decimal = -1
        Try
            Dim epoch As DateTime = New DateTime(1900, 1, 1)
            Dim difference As TimeSpan = DateTime.Now - epoch
            days = difference.TotalDays
            Return days
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return days
        End Try

    End Function

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