Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data
Imports System.Data.OleDb

Public Class LoadExcel

    'Private Excel03ConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"
    'Private Excel07ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"

    'Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
    '    OpenFileDialog1.ShowDialog()
    'End Sub

    'Private Sub openFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
    '    Dim filePath As String = OpenFileDialog1.FileName
    '    Dim extension As String = Path.GetExtension(filePath)
    '    Dim header As String = If(rbHeaderYes.Checked, "YES", "NO")
    '    Dim conStr As String, sheetName As String

    '    conStr = String.Empty
    '    Select Case extension

    '        Case ".xls"
    '            'Excel 97-03
    '            conStr = String.Format(Excel03ConString, filePath, header)
    '            Exit Select

    '        Case ".xlsx"
    '            'Excel 07
    '            conStr = String.Format(Excel07ConString, filePath, header)
    '            Exit Select
    '    End Select

    '    'Get the name of the First Sheet.
    '    Using con As New OleDbConnection(conStr)
    '        Using cmd As New OleDbCommand()
    '            cmd.Connection = con
    '            con.Open()
    '            Dim dtExcelSchema As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '            sheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
    '            con.Close()
    '        End Using
    '    End Using

    '    'Read Data from the First Sheet.
    '    Using con As New OleDbConnection(conStr)
    '        Using cmd As New OleDbCommand()
    '            Using oda As New OleDbDataAdapter()
    '                Dim dt As New DataTable()
    '                cmd.CommandText = (Convert.ToString("SELECT * From [") & sheetName) + "]"
    '                cmd.Connection = con
    '                con.Open()
    '                oda.SelectCommand = cmd
    '                oda.Fill(dt)
    '                con.Close()

    '                'Populate DataGridView.
    '                dataGridView1.DataSource = dt
    '            End Using
    '        End Using
    '    End Using
    'End Sub

End Class
