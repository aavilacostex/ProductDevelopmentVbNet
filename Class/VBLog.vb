Imports System.Configuration

Public Class VBLog

    'Private Shared eventLog1 As EventLog = New EventLog(Constantes.logName, Environment.MachineName, Constantes.source);

    Shared gnr As Gn1 = New Gn1()
    Private Shared eventLog1 As EventLog = New EventLog("CTPSystem-Log", GetComputerName(), "CTPSystem-Net")

    Private Shared strconnSQL As String
    Public Shared Property SQLCon() As String
        Get
            Return strconnSQL
        End Get
        Set(ByVal value As String)
            strconnSQL = value
        End Set
    End Property

    Private _Source As String
    Public Property Source() As String
        Get
            Return _Source
        End Get
        Set(ByVal value As String)
            _Source = value
        End Set
    End Property

    Private _LogName As String
    Public Property LogName() As String
        Get
            Return _LogName
        End Get
        Set(ByVal value As String)
            _LogName = value
        End Set
    End Property

    Public Sub New()
        SQLCon = ConfigurationManager.AppSettings("strconnSQL").ToString()
        Source = ConfigurationManager.AppSettings("Source").ToString()
        LogName = ConfigurationManager.AppSettings("LogName").ToString()
    End Sub


    Public Enum ErrorTypeEnum
        [Start]
        [Stop]
        [Information]
        [Error]
        [Trace]
        [Warning]
        [Exception]
    End Enum

    Public Sub WriteLog(ErrorType As ErrorTypeEnum, strTipo As String, strMethod As String, strUser As String, strMessage As String, strDetails As String)
        Dim exMessage As String = Nothing
        Try
            Dim logMapping As String = If(String.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings("LogMapping").ToLower()),
                                            "all", System.Configuration.ConfigurationManager.AppSettings("LogMapping").ToLower())

            If System.String.Compare(logMapping.ToLower(), "none", System.StringComparison.Ordinal) = 0 Then
            Else
                WriteToLogDB(ErrorType, strTipo, strMethod, strUser, strMessage, strDetails)
                If ErrorType.Equals(ErrorTypeEnum.Error) Then
                    'send email
                End If
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Shared Sub WriteToLogDB(ErrorType As ErrorTypeEnum, strTipo As String, strMethod As String, strUsuario As String, strMessage As String, strDetalle As String)
        Dim exMessage As String = Nothing
        Try

            'Dim sqlQuery As String = "DELETE FROM {0} WHERE USERID = '{1}'"

            Dim dt As DateTime = DateTime.Now
            Dim curDate As String = dt.ToString()

            Dim sqlQuery As String = "INSERT INTO dbCTPSystem.dbo.CtpSystemLog (LOGAPP,LOGLEVEL,LOGTYPE,LOGUSER,LOGORIGEN,LOGMESSAGE,LOGEXCEPTION,LOGDATE )
                                        VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, "CTPSystem-Net", ErrorType, strTipo, strUsuario, strMethod, strMessage, strDetalle, curDate)
            Dim rsResult = gnr.ExecuteNotQueryCommand(sqlFormattedQuery, SQLCon)

            Dim pepe = rsResult

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            If Not EventLog.SourceExists("CTPSystem-Log") Then
                EventLog.CreateEventSource("CTPSystem-Net", "CTPSystem-Log")
            End If
            eventLog1 = New EventLog("CTPSystem-Log", Environment.MachineName, "CTPSystem-Net")
            eventLog1.WriteEntry("Error: " + ex.Message, EventLogEntryType.Error)
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
            Return Nothing
        End Try
    End Function


#Region "IDisposable Members"

    'Public Sub Dispose()
    'this.Dispose();
    'End Sub

#End Region

End Class
