Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Xml
Imports DocumentFormat.OpenXml.Office2010.ExcelAc

Public Class ConvertXml : Implements IDisposable

    Private disposedValue As Boolean

    Public Function CreateXltoXML(dt As DataTable, XmlFile As String, RowName As String, InnerName As String) As Boolean
        Dim exMessage As String = " "
        Dim IsCreated As Boolean = False
        Try

            'Dim dt As DataTable = GetTableDataXl(XlFile)
            'Dim writer As XmlTextWriter = New XmlTextWriter(XmlFile, System.Text.Encoding.UTF8)
            Dim i As Integer = 0
            Dim ColumnNames As List(Of String) = dt.Columns.Cast(Of DataColumn)().ToList().Select(Function(x) x.ColumnName).ToList() 'Column Names List  
            Dim RowList As List(Of DataRow) = dt.Rows.Cast(Of DataRow)().ToList()

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = ("    ")
            settings.CloseOutput = True
            settings.OmitXmlDeclaration = True

            Dim culture As CultureInfo = CultureInfo.CreateSpecificCulture("en-US")
            Dim dtfi As DateTimeFormatInfo = culture.DateTimeFormat
            dtfi.DateSeparator = "."

            Dim now As DateTime = DateTime.Now
            Dim halfName = now.ToString("G", dtfi)
            halfName = halfName.Replace(" ", ".")
            halfName = halfName.Replace(":", "")
            Dim fileName = "Input." & halfName & ".xml"
            XmlFile += fileName

            Using fs As New FileStream(XmlFile, FileMode.Create)
                Using writer1 As XmlWriter = XmlWriter.Create(fs, settings)
                    writer1.WriteStartElement(RowName)
                    'writer1.Formatting = Formatting.Indented
                    'writer1.Indentation = 2

                    For Each dr As DataRow In RowList
                        writer1.WriteStartElement(InnerName)
                        For Each str As String In ColumnNames
                            If str = "ErrorDesc" Then
                                Continue For
                            ElseIf str = "VMVNUM" Then
                                Continue For
                            Else
                                writer1.WriteStartElement(str)
                                writer1.WriteString(dr.ItemArray(i).ToString())
                                writer1.WriteEndElement()
                                i += 1
                            End If
                        Next
                        writer1.WriteEndElement()
                        i = 0
                    Next
                    writer1.WriteEndElement()
                    writer1.WriteEndDocument()
                    writer1.Flush()

                End Using

                If (File.Exists(XmlFile)) Then
                    IsCreated = True
                    LikeSession.fullFilePath = XmlFile
                End If

                fs.Dispose()

            End Using

            GC.Collect()
            GC.WaitForPendingFinalizers()

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

        Return IsCreated

    End Function

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    'Public Sub Dispose()
    '    'Me.Close
    'End Sub

End Class
