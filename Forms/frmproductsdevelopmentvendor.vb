Imports System.Globalization
Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class frmproductsdevelopmentvendor

    Dim gnr As Gn1 = New Gn1()
    Public userid As String

    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

#Region "Action Methods"

    Private Sub frmproductsdevelopmentvendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Load()
    End Sub

    Private Sub cmdexit1_Click(sender As Object, e As EventArgs) Handles cmdexit1.Click
        Me.Close()
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

            Dim dsInvPrdoDet = gnr.GetInvProdDetailByProject(codeproject)
            fillcell2(dsInvPrdoDet)

            'userid = frmLogin.txtUserName.Text
            userid = LikeSession.userid
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Information, "User Info - Massive Vendor Update Start", "")

            CType(Me.DataGridView1.Columns(3), DataGridViewTextBoxColumn).MaxInputLength = 6

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString

            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            writeComputerEventLog()
        End Try
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
                    DataGridView1.ColumnCount = 3

                    'Add Columns
                    DataGridView1.Columns(0).Name = "clPartNo"
                    DataGridView1.Columns(0).HeaderText = "Part No."
                    DataGridView1.Columns(0).DataPropertyName = "PRDPTN"

                    DataGridView1.Columns(1).Name = "clDescription"
                    DataGridView1.Columns(1).HeaderText = "Descripcion"
                    DataGridView1.Columns(1).DataPropertyName = "IMDSC"

                    DataGridView1.Columns(2).Name = "clVendorNo"
                    DataGridView1.Columns(2).HeaderText = "Vendor No."
                    DataGridView1.Columns(2).DataPropertyName = "VMVNUM"

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
                    DataGridView1.Columns("clVendorNo").ReadOnly = True


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

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) _
    Handles DataGridView1.DataError
        Dim exMessage As String = " "
        Try
            If e.ColumnIndex = 3 Then
                Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
                Dim inputText = DataGridView1.EditingControl.Text
                If Not Regex.IsMatch(inputText, "^[0-9]{1,6}$") Then
                    DataGridView1.CancelEdit()
                    DataGridView1.RefreshEdit()
                    MessageBox.Show("The Vendor Number must be changed for a numeric value!", "CTP System", MessageBoxButtons.OK)
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub Datagridview1_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) _
        Handles DataGridView1.CellBeginEdit
        Try
            If DataGridView1(e.ColumnIndex, e.RowIndex).Value IsNot Nothing Then
                Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
                LikeSession.flyingValue = value
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub Datagridview1_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
        Handles DataGridView1.CellContentClick
        Try
            If e.ColumnIndex = 0 Then
                If DataGridView1(e.ColumnIndex, e.RowIndex).Value IsNot Nothing Then
                    Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
                    Dim inputText = If(DataGridView1.EditingControl IsNot Nothing, DataGridView1.EditingControl.Text, Nothing)

                    DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    If CBool(DataGridView1.CurrentCell.Value) = True Then
                        Dim ppe = ""
                        Dim calros = "1"

                        Dim ok = ppe & " - " & calros
                    Else
                        Dim ppe = ""
                        Dim calros = "1"

                        Dim ok = ppe & " - " & calros
                    End If
                End If
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView1_CellMouseUp(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) _
        Handles DataGridView1.CellMouseUp
        Dim exMessage As String = " "
        Try
            If e.ColumnIndex = 0 Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                row.Cells("checkBoxColumn").Value = Convert.ToBoolean(row.Cells("checkBoxColumn").EditedFormattedValue)
                If Convert.ToBoolean(row.Cells("checkBoxColumn").Value) Then
                    Dim value = DataGridView1(3, e.RowIndex).Value.ToString()
                    LikeSession.flyingValue = value
                    DataGridView1(3, e.RowIndex).ReadOnly = False
                Else
                    DataGridView1(3, e.RowIndex).ReadOnly = True
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) _
    Handles DataGridView1.CellValueChanged
        Dim exMessage As String = " "
        Try
            If e.ColumnIndex = 3 Then
                Dim value = DataGridView1(e.ColumnIndex, e.RowIndex).Value.ToString()
                Dim inputText = If(DataGridView1.EditingControl IsNot Nothing, DataGridView1.EditingControl.Text, Nothing)
                If Not Regex.IsMatch(inputText, "^[0-9]{1,6}$") Then
                    ' DataGridView1(e.ColumnIndex, e.RowIndex).Value = LikeSession.flyingValue
                    DataGridView1.CancelEdit()
                    DataGridView1.RefreshEdit()
                    MessageBox.Show("The Vendor Number must be changed for a numeric value!", "CTP System", MessageBoxButtons.OK)
                ElseIf Not gnr.isVendorAccepted(inputText) Then
                    'DataGridView1(e.ColumnIndex, e.RowIndex).Value = LikeSession.flyingValue
                    DataGridView1.CancelEdit()
                    DataGridView1.RefreshEdit()
                    MessageBox.Show("Invalid Vendor Number.", "CTP System", MessageBoxButtons.OK)
                Else
                    Dim result = cmdSave_custom(inputText)
                    If result = -1 Then
                        DataGridView1.CancelEdit()
                        DataGridView1.RefreshEdit()
                        MessageBox.Show("The Vendor Number does not exist in our records!", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
            End If
        Catch ex As Exception
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

    Private Function cmdSave_custom(vendorNo As String) As Integer
        Dim result As String = -1
        Dim exMessage As String = " "
        Try
            Dim dsVendor = gnr.GetVendorByVendorNo(Trim(UCase(vendorNo)))
            If dsVendor Is Nothing Then
                Return result
            Else
                If dsVendor.Tables(0).Rows.Count <= 0 Then
                    Return result
                Else
                    Return result = 0
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return result
        End Try

    End Function

    Private Sub cmdSave1_Click(sender As Object, e As EventArgs) Handles cmdSave1.Click
        Dim exMessage As String = " "
        Dim ds As DataSet
        Dim updatedRecords As Integer = 0
        Dim strVendorErrors As String = ""
        Dim value As String
        Dim lstChar As Array
        Dim lngLstArray As Integer
        Dim strDupMessage As String
        Dim strEndMessage As String
        Try

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("checkBoxColumn").Value = True Then
                    If cmdSave_custom(row.Cells("clVendorNo").Value.ToString()) = 0 Then
                        ds = gnr.GetDataByCodeAndPartNoProdDesc(frmProductsDevelopment.txtCode.Text, row.Cells("clPartNo").Value.ToString())
                        Dim oldVendorNo = ds.Tables(0).Rows(0).ItemArray(ds.Tables(0).Columns("VMVNUM").Ordinal)
                        If Trim(UCase(oldVendorNo)) <> Trim(UCase(row.Cells("clVendorNo").Value.ToString())) Then
                            Dim dsVendor = gnr.GetVendorByVendorNo(Trim(UCase(row.Cells("clVendorNo").Value.ToString())))
                            If dsVendor IsNot Nothing Then
                                If dsVendor.Tables(0).Rows.Count > 0 Then
                                    PoQotaFunction(oldVendorNo, row.Cells("clPartNo").Value.ToString(), row.Cells("clVendorNo").Value.ToString())
                                    gnr.UpdateChangedVendor(userid, row.Cells("clVendorNo").Value.ToString(), row.Cells("clPartNo").Value.ToString(), frmProductsDevelopment.txtCode.Text)
                                    'update validation
                                    updatedRecords += 1
                                Else
                                    strVendorErrors += Trim(UCase(row.Cells("clVendorNo").Value.ToString())) & ","
                                End If
                            Else
                                strVendorErrors += Trim(UCase(row.Cells("clVendorNo").Value.ToString())) & ","
                            End If
                        End If
                    Else
                        MessageBox.Show("The Vendor Number does not exist in our records!", "CTP System", MessageBoxButtons.OK)
                    End If
                End If
            Next

            If updatedRecords > 0 Then
                If String.IsNullOrEmpty(strVendorErrors) Then
                    MessageBox.Show("Records Updated.", "CTP System", MessageBoxButtons.OK)
                Else
                    value = strVendorErrors(strVendorErrors.Length - 1)
                    strEndMessage = If(value.Equals(","), strVendorErrors.Substring(0, strVendorErrors.Length - 1), strVendorErrors)
                    lstChar = strVendorErrors.Split(",")
                    lngLstArray = lstChar.Length
                    strDupMessage = If(lngLstArray > 0, "There are errors with this vendor numbers.", "There is an error with this vendor number.")

                    MessageBox.Show("Check the data input. " & strDupMessage & " - " & strEndMessage, "CTP System", MessageBoxButtons.OK)
                End If
                Form_Load()
                frmProductsDevelopment.fillcell2(frmProductsDevelopment.txtCode.Text)
            Else
                If String.IsNullOrEmpty(strVendorErrors) Then
                    MessageBox.Show("No records to update.", "CTP System", MessageBoxButtons.OK)
                Else
                    value = strVendorErrors(strVendorErrors.Length - 1)
                    strEndMessage = If(value.Equals(","), strVendorErrors.Remove(strVendorErrors.Length - 1), strVendorErrors)
                    lstChar = strEndMessage.Split(",")
                    lngLstArray = lstChar.Length
                    strDupMessage = If(lngLstArray > 1, "There are errors with this vendor numbers.", "There is an error with this vendor number.")

                    MessageBox.Show("Check the data input. " & strDupMessage & " - " & strEndMessage, "CTP System", MessageBoxButtons.OK)
                    Form_Load()
                    frmProductsDevelopment.fillcell2(frmProductsDevelopment.txtCode.Text)
                End If
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Sub PoQotaFunction(oldVendorNo As String, partNo As String, newVendorNo As String)
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim dsUpdatedData As Integer
        Dim strQueryAdd As String = "WHERE PQVND = " & Trim(newVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'"
        Try
            statusquote = "D-" & Status2
            Dim spacepoqota As String = String.Empty
            'Dim strQueryAdd As String = "WHERE PQVND = " & Trim(txtvendorno.Text) & " AND PQPTN = '" & Trim(UCase(txtpartno.Text)) & "'"
            Dim dsPoQota = gnr.GetPOQotaData(oldVendorNo, partNo)
            If dsPoQota IsNot Nothing Then
                If dsPoQota.Tables(0).Rows.Count > 0 Then
                    Dim poqSeq = dsPoQota.Tables(0).Rows(0).ItemArray(dsPoQota.Tables(0).Columns("PQSEQ").Ordinal)
                    Dim rsResult = PoQotaFunctionDuplex(newVendorNo, partNo, poqSeq)
                    If rsResult = 0 Then
                        Dim updatedSeq = CInt(poqSeq) + 1
                        dsUpdatedData = gnr.UpdatePoQoraRowVendor(oldVendorNo, newVendorNo, partNo, updatedSeq)
                        'validation result
                    Else
                        dsUpdatedData = gnr.UpdatePoQoraRowVendor(oldVendorNo, newVendorNo, partNo, poqSeq)
                        'validation result
                    End If
                Else
                    'error message
                End If
            Else
                Dim maxValue As Integer = 1
                Dim maxQotaVal = gnr.getmaxComplex("POQOTA", "PQSEQ", strQueryAdd)
                If Not String.IsNullOrEmpty(maxQotaVal) Then
                    maxValue = CInt(Trim(maxQotaVal)) + 1
                End If
                spacepoqota = "                               DEV"
                Dim rsInsertion = gnr.InsertNewPOQotaLess(partNo, newVendorNo, maxValue, DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), "", DateTime.Now.Day.ToString(), "", spacepoqota, 0)
                'insertion validation
            End If
        Catch ex As Exception
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Private Function PoQotaFunctionDuplex(newVendorNo As String, partNo As String, seqNo As String) As Integer
        Dim exMessage As String = " "
        Dim statusquote As String
        Dim Status2 As String = ""
        Dim strQueryAdd As String = " WHERE PQVND = " & Trim(newVendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND PQSEQ = '" & Trim(seqNo) & "'"
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