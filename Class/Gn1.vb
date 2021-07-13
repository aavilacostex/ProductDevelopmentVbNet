Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.Diagnostics
Imports System.Configuration
Imports System.Reflection

NotInheritable Class Gn1

    'Public Const Version = "V.02/20/20"
    'Public Const strCompany = "COSTEX"
    'Public Const strdatabase = "dbCTPSystem"

    'Public pathgeneral As String
    'Public Const strconnection = "DSN=COSTEX400;UID=INTRANET;PWD=CTP6100"
    'Public Const strcrystalconn = "DSN=COSTEX400;UID=INTRANET;PWD=CTP6100;"
    'Public Const strconnSQL = "Data Source=CTPSystem;Initial Catalog=dbCTPSystem;User Id=sa;Password=ctp6100;"
    'Public Const strconnSQL = "DSN=CTPSystem;UID=sa;PWD=ctp6100;"
    'Public Const strcrystalconnSQL = "DSN=CTPSystem;UID=sa;PWD=ctp6100;"
    'Public Const strmailhostctp = "mail.costex.com"
    'Public Const strmailhostctp = "mail.costex.com"
    'Public Const strconnSQL = "Driver={SQL Server};Server=dellserver;Database=dbCTPSystem;Uid=sa;Pwd=ctp6100;"
    'Public Const strcrystalconnSQL = "Driver={SQL Server};Server=dellserver;Database=dbCTPSystem;Uid=sa;Pwd=ctp6100;"
    'Public Const strconnSQLNOVA = "DSN=NOVATIME;UID=NTI_CS;PWD=csadmin;"
    'Public Const strcrystalconnSQLNOVA = "DSN=NOVATIME;UID=NTI_CS;PWD=csadmin;"
    Public pathpicture As String
    Public emailspath As String
    Public pictureSizeFlag As Integer
    Public filesQuantity As Integer
    Public filesToWrite As Integer
    Public userDepartment As String
    Public flagSalesman As Integer
    Public primaryservername As String
    Public Conn As New ADODB.Connection
    Public ConnSql As New ADODB.Connection
    Public ConnSqlNOVA As New ADODB.Connection
    Public CMD As New ADODB.Command
    Public rs As New ADODB.Recordset
    Public Rs1 As New ADODB.Recordset
    Public Rs2 As New ADODB.Recordset
    Public Rs3 As New ADODB.Recordset
    Public Rs4 As New ADODB.Recordset
    Public RsGeneral As New ADODB.Recordset
    Public userid As String
    Public flagexit As Integer
    Public flaguserrec As Integer
    Public getclaimflag As Integer
    Public claimsplit As Integer
    Public getclaimnosave As Integer
    Public seeaddcomments As Integer
    Public seeaddprocomments As Integer
    'Public printpath As String
    Public flagchangevendor As Integer
    Public getclaimno As Long
    Public getclaim As Long
    Public pass As String
    Public check As String
    Public encrpwd As String
    Public passcomm As String
    Public pototalcost As Double
    Public actupdatepo As Long
    Public prpagrid As Integer
    Public fso As New Scripting.FileSystemObject
    Public IP As String
    Public ipaddresslocal As String
    Public Provider As String
    Public projectnoadd As Long
    Public DataSource As String
    Public user As String
    Public password As String
    Public InitialCatalog As String
    'Public objMail As New MailSender
    Public strHost As String, strPort As String, strfrom As String
    Public strFromName As String, strto As String, strSubject As String
    Public strBody As String, stratt As String, stratt1 As String
    Public DirTrabajo, DirLog As String
    Public LoginSucceeded As Boolean
    Public codloginctp As Long
    Public strWarr As String, strNonw As String, strIntr As String
    Public strOpen As String, strClos As String, strwhere As String
    Public strInts As String, strPcus As String, strPoth As String, strFinl As String
    Public countgridrows As Integer
    'variables para convertir un amount en text - begin
    'Set up two arrays to hold string values we
    'will use to convert numbers to words
    Public BigOnes(9) As String
    Public SmallOnes(19) As String
    'Declare variables
    Public Dollars As String
    Public Cents As String
    Public Words As String
    Public Chunk As String
    Public Digits As Integer
    Public LeftDigit As Integer
    Public folderpathproject As String
    Public folderpathvendor As String
    Public FolderPath As String
    Public folderpathpart As String
    Public pathfolderfrom As String

    Dim vblog As VBLog = New VBLog()

    Private strLogCadenaCabecera As String = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString()
    Dim strLogCadena As String = Nothing


#Region "Defaul Variables"

    Private Pathgeneral As String
    Public Property Path() As String
        Get
            Return Pathgeneral
        End Get
        Set(ByVal value As String)
            Pathgeneral = value
        End Set
    End Property

    Private ConString As String
    Public Property Number() As String
        Get
            Return ConString
        End Get
        Set(ByVal value As String)
            ConString = value
        End Set
    End Property

    Private Version As String
    Public Property VersionNo() As String
        Get
            Return Version
        End Get
        Set(ByVal value As String)
            Version = value
        End Set
    End Property

    Private strCompany As String
    Public Property Company() As String
        Get
            Return strCompany
        End Get
        Set(ByVal value As String)
            strCompany = value
        End Set
    End Property

    Private strdatabase As String
    Public Property Database() As String
        Get
            Return strdatabase
        End Get
        Set(ByVal value As String)
            strdatabase = value
        End Set
    End Property

    Private strconnection As String
    Public Property Conexion() As String
        Get
            Return strconnection
        End Get
        Set(ByVal value As String)
            strconnection = value
        End Set
    End Property

    Private strcrystalconn As String
    Public Property CrystalCon() As String
        Get
            Return strcrystalconn
        End Get
        Set(ByVal value As String)
            strcrystalconn = value
        End Set
    End Property

    Private strconnSQL As String
    Public Property SQLCon() As String
        Get
            Return strconnSQL
        End Get
        Set(ByVal value As String)
            strconnSQL = value
        End Set
    End Property

    Private strcrystalconnSQL As String
    Public Property CrystalSQLCon() As String
        Get
            Return strcrystalconnSQL
        End Get
        Set(ByVal value As String)
            strcrystalconnSQL = value
        End Set
    End Property

    Private strmailhostctp As String
    Public Property MailHostCTP() As String
        Get
            Return strmailhostctp
        End Get
        Set(ByVal value As String)
            strmailhostctp = value
        End Set
    End Property

    Private strconnSQLNOVA As String
    Public Property NOVASQLCon() As String
        Get
            Return strconnSQLNOVA
        End Get
        Set(ByVal value As String)
            strconnSQLNOVA = value
        End Set
    End Property

    Private strcrystalconnSQLNOVA As String
    Public Property NOVASQLCRYSTALCon() As String
        Get
            Return strcrystalconnSQLNOVA
        End Get
        Set(ByVal value As String)
            strcrystalconnSQLNOVA = value
        End Set
    End Property

    Private printpath As String
    Public Property ReportsValue() As String
        Get
            Return printpath
        End Get
        Set(ByVal value As String)
            printpath = value
        End Set
    End Property

    Private jiraPathBase As String
    Public Property JiraPathBaseValue() As String
        Get
            Return jiraPathBase
        End Get
        Set(ByVal value As String)
            jiraPathBase = value
        End Set
    End Property

    Private UrlPartFiles As String
    Public Property UrlPartFilesMethod() As String
        Get
            Return UrlPartFiles
        End Get
        Set(ByVal value As String)
            UrlPartFiles = value
        End Set
    End Property

    Private UrlPDevelopment As String
    Public Property UrlPDevelopmentMethod() As String
        Get
            Return UrlPDevelopment
        End Get
        Set(ByVal value As String)
            UrlPDevelopment = value
        End Set
    End Property

    Private FlagProduction As String
    Public Property FlagProductionMethod() As String
        Get
            Return FlagProduction
        End Get
        Set(ByVal value As String)
            FlagProduction = value
        End Set
    End Property

    Private UrlPathGeneral As String
    Public Property UrlPathGeneralMethod() As String
        Get
            Return UrlPathGeneral
        End Get
        Set(ByVal value As String)
            UrlPathGeneral = value
        End Set
    End Property

    Private VendorOEMCodeDenied As String
    Public Property VendorOEMCodeDeniedMethod() As String
        Get
            Return VendorOEMCodeDenied
        End Get
        Set(ByVal value As String)
            VendorOEMCodeDenied = value
        End Set
    End Property

    Private VendorCodesDenied As String
    Public Property VendorCodesDeniedMethod() As String
        Get
            Return VendorCodesDenied
        End Get
        Set(ByVal value As String)
            VendorCodesDenied = value
        End Set
    End Property

    Private VendorWhiteFlag As String
    Public Property VendorWhiteFlagMethod() As String
        Get
            Return VendorWhiteFlag
        End Get
        Set(ByVal value As String)
            VendorWhiteFlag = value
        End Set
    End Property

    Private PathStartImage As String
    Public Property PathStartImageMethod() As String
        Get
            Return PathStartImage
        End Get
        Set(ByVal value As String)
            PathStartImage = value
        End Set
    End Property

    Private UrlPathImgNew As String
    Public Property UrlPathImgNewMethod() As String
        Get
            Return UrlPathImgNew
        End Get
        Set(ByVal value As String)
            UrlPathImgNew = value
        End Set
    End Property

    Private UrlPathXsdFile As String
    Public Property UrlPathXsdFileMethod() As String
        Get
            Return UrlPathXsdFile
        End Get
        Set(ByVal value As String)
            UrlPathXsdFile = value
        End Set
    End Property

    Private authorizeUser As String
    Public Property AuthorizatedUser() As String
        Get
            Return authorizeUser
        End Get
        Set(ByVal value As String)
            authorizeUser = value
        End Set
    End Property

    Private authorizeTestUser As String
    Public Property AuthorizatedTestUser() As String
        Get
            Return authorizeTestUser
        End Get
        Set(ByVal value As String)
            authorizeTestUser = value
        End Set
    End Property

    Private newMenuCodes As String
    Public Property NewUserMenuCodes() As String
        Get
            Return newMenuCodes
        End Get
        Set(ByVal value As String)
            newMenuCodes = value
        End Set
    End Property

    Private CloseMDIForm As String
    Public Property FlagCloseMDIForm() As String
        Get
            Return CloseMDIForm
        End Get
        Set(ByVal value As String)
            CloseMDIForm = value
        End Set
    End Property

    Private sendTestEmails As String
    Public Property FlagTestEmails() As String
        Get
            Return sendTestEmails
        End Get
        Set(ByVal value As String)
            sendTestEmails = value
        End Set
    End Property

    Private testEmails As String
    Public Property TestEmailAddresess() As String
        Get
            Return testEmails
        End Get
        Set(ByVal value As String)
            testEmails = value
        End Set
    End Property

    Private excelColumnNames As String
    Public Property GetColumnNames() As String
        Get
            Return excelColumnNames
        End Get
        Set(ByVal value As String)
            excelColumnNames = value
        End Set
    End Property

    Private closeStatus As String
    Public Property GetCloseStatus() As String
        Get
            Return closeStatus
        End Get
        Set(ByVal value As String)
            closeStatus = value
        End Set
    End Property

    Private _pathPdTemplate As String
    Public Property getPdExcelTemplate() As String
        Get
            Return _pathPdTemplate
        End Get
        Set(ByVal value As String)
            _pathPdTemplate = value
        End Set
    End Property

    Private _referenceUsersReports As String
    Public Property ReferenceUsersReport() As String
        Get
            Return _referenceUsersReports
        End Get
        Set(ByVal value As String)
            _referenceUsersReports = value
        End Set
    End Property

    Private _vendorOEMExclude As String
    Public Property VendorOEMExclude() As String
        Get
            Return _vendorOEMExclude
        End Get
        Set(ByVal value As String)
            _vendorOEMExclude = value
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

    Private _automaticExcel As String
    Public Property AutomaticExcel() As String
        Get
            Return _automaticExcel
        End Get
        Set(ByVal value As String)
            _automaticExcel = value
        End Set
    End Property

    Private _processName As String
    Public Property ProcessName() As String
        Get
            Return _processName
        End Get
        Set(ByVal value As String)
            _processName = value
        End Set
    End Property

    Private _excelUserTest As String
    Public Property ExcelUserTest() As String
        Get
            Return _excelUserTest
        End Get
        Set(ByVal value As String)
            _excelUserTest = value
        End Set
    End Property

#End Region

    Private Shared ReadOnly Log As log4net.ILog = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New()
        ConString = ConfigurationManager.AppSettings("ConnectionString").ToString()
        Pathgeneral = ConfigurationManager.AppSettings("urlPathGeneral").ToString()
        Version = ConfigurationManager.AppSettings("Version").ToString()
        Company = ConfigurationManager.AppSettings("strCompany").ToString()
        Database = ConfigurationManager.AppSettings("strdatabase").ToString()
        Conexion = ConfigurationManager.AppSettings("strconnection").ToString()
        CrystalCon = ConfigurationManager.AppSettings("strcrystalconn").ToString()
        SQLCon = ConfigurationManager.AppSettings("strconnSQL").ToString()
        CrystalSQLCon = ConfigurationManager.AppSettings("strcrystalconnSQL").ToString()
        MailHostCTP = ConfigurationManager.AppSettings("strmailhostctp").ToString()
        NOVASQLCon = ConfigurationManager.AppSettings("strconnSQLNOVA").ToString()
        NOVASQLCRYSTALCon = ConfigurationManager.AppSettings("strcrystalconnSQLNOVA").ToString()
        ReportsValue = ConfigurationManager.AppSettings("printpath").ToString()
        JiraPathBaseValue = ConfigurationManager.AppSettings("urlPathBase").ToString()
        UrlPartFiles = ConfigurationManager.AppSettings("urlPartFiles").ToString()
        UrlPDevelopment = ConfigurationManager.AppSettings("urlPDevelopment").ToString()
        FlagProduction = ConfigurationManager.AppSettings("flagProduction").ToString()
        UrlPathGeneral = ConfigurationManager.AppSettings("urlPathGeneral").ToString()
        VendorOEMCodeDenied = ConfigurationManager.AppSettings("vendorOEMCodeDenied").ToString()
        VendorCodesDenied = ConfigurationManager.AppSettings("vendorCodesDenied").ToString()
        VendorWhiteFlag = ConfigurationManager.AppSettings("itemCategories").ToString()
        PathStartImage = ConfigurationManager.AppSettings("urlPathStartImg").ToString()
        UrlPathImgNew = ConfigurationManager.AppSettings("urlPathImgNew").ToString()
        UrlPathXsdFile = ConfigurationManager.AppSettings("urlPathXsdFile").ToString()
        AuthorizatedUser = ConfigurationManager.AppSettings("authorizeUser").ToString()
        AuthorizatedTestUser = ConfigurationManager.AppSettings("authorizeTestUser").ToString()
        NewUserMenuCodes = ConfigurationManager.AppSettings("newMenuCodes").ToString()
        FlagCloseMDIForm = ConfigurationManager.AppSettings("hideMDIForm").ToString()
        FlagTestEmails = ConfigurationManager.AppSettings("sendToTestEmails").ToString()
        TestEmailAddresess = ConfigurationManager.AppSettings("testEmails").ToString()
        GetColumnNames = ConfigurationManager.AppSettings("checkColumns").ToString()
        GetCloseStatus = ConfigurationManager.AppSettings("closeStatus").ToString()
        getPdExcelTemplate = ConfigurationManager.AppSettings("urlPathPDTemplate").ToString()
        ReferenceUsersReport = ConfigurationManager.AppSettings("referenceUsersReports").ToString()
        VendorOEMExclude = ConfigurationManager.AppSettings("vendorOEMExclude").ToString()
        Source = ConfigurationManager.AppSettings("Source").ToString()
        LogName = ConfigurationManager.AppSettings("LogName").ToString()
        AutomaticExcel = ConfigurationManager.AppSettings("AutomaticExcel").ToString()
        ProcessName = ConfigurationManager.AppSettings("ProcessName").ToString()
        ExcelUserTest = ConfigurationManager.AppSettings("UserExcelTest").ToString()
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Boolean, ByVal Param As IntPtr) As Integer
    End Function

    Private Const WM_SETREDRAW As Integer = 11


    <DllImport("IpHlpApi.dll")>
    Private Shared Function GetIpAddrTable_API(pIPAddrTable As String, pdwSize As Long, ByVal bOrder As Long) As Long
    End Function

    <DllImport("IpHlpApi.dll")>
    Private Shared Function GetIpNetTable(pIpNetTable As IntPtr, <MarshalAs(UnmanagedType.U4)> ByRef pdwSize As Integer, bOrder As Boolean) As <MarshalAs(UnmanagedType.U4)> Integer
    End Function
    Public Const ERROR_SUCCESS As Integer = 0
    Public Const ERROR_INSUFFICIENT_BUFFER As Integer = 122

    Public Structure MIB_IPNETROW
        <MarshalAs(UnmanagedType.U4)>
        Public dwIndex As UInteger
        <MarshalAs(UnmanagedType.U4)>
        Public dwPhysAddrLen As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)>
        Public bPhysAddr() As Byte
        <MarshalAs(UnmanagedType.U4)>
        Public dwAddr As UInteger
        <MarshalAs(UnmanagedType.U4)>
        Public dwType As DWTYPES
    End Structure

    Public Enum DWTYPES As UInteger

        <MarshalAs(UnmanagedType.U4)>
        Other = 1
        <MarshalAs(UnmanagedType.U4)>
        Invalid = 2
        <MarshalAs(UnmanagedType.U4)>
        Dynamic = 3
        <MarshalAs(UnmanagedType.U4)>
        [Static] = 4
    End Enum

    Public RightDigit As Integer
    Public instanceOfModel_ID As Integer
    Public test As String
    Public Const VBObjectError As Integer = -2147221504
    Public Versionctp As String
    Public strDate As String = "1900,01,01"
    Public formats() As String = {"M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt", "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss", "M/d/yyyy hh:mm tt",
        "M/d/yyyy hh tt", "M/d/yyyy h:mm", "M/d/yyyy h:mm", "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm", "MM/d/yyyy HH:mm:ss.ffffff"}


#Region "Selects"

#Region "Optimized"

    Public Function getReferencesStatusesByCode(code As String) As DataSet
        Dim exMessage As String = Nothing
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select prdsts from qs36f.prdvld where prhcod =  " & code & " "
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds
                End If
            End If
            Return Nothing
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getVendorNoAndNameByNameDS() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT VMVNUM, VMNAME, VMVTYP FROM VNMAS WHERE VMVTYP NOT IN (" & VendorCodesDenied & ") 
                   AND VMVNUM NOT IN (SELECT CNTDE1 FROM CNTRLL WHERE CNT01 IN (" & VendorOEMCodeDenied & "))
                   ORDER BY VMNAME"
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds
                End If
            End If
            Return Nothing
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getVendorNoAndNameByName() As AutoCompleteStringCollection
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        'Dim rsValue As Integer = -1
        Dim lstCollection As AutoCompleteStringCollection = New AutoCompleteStringCollection()
        Try
            'Sql = "SELECT VMVNUM, VMNAME FROM VNMAS WHERE VMNAME LIKE '%" & Replace(Trim(UCase(vendorName)), "'", "") & "%' "
            Sql = "SELECT VMVNUM, VMNAME FROM VNMAS ORDER BY VMVNUM, VMNAME ASC"
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then

                    For Each row As DataRow In ds.Tables(0).Rows
                        lstCollection.Add(row(0).ToString())
                    Next
                    Return lstCollection
                End If
            End If
            Return Nothing
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getVendorNoAndNameByNameLike(vendorName As String) As AutoCompleteStringCollection
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        'Dim rsValue As Integer = -1
        Dim lstCollection As AutoCompleteStringCollection = New AutoCompleteStringCollection()
        Try
            Sql = "SELECT VMNAME FROM VNMAS WHERE VMNAME LIKE '%" & Replace(Trim(UCase(vendorName)), "'", "") & "%' "
            'Sql = "SELECT VMVNUM, VMNAME FROM VNMAS ORDER BY VMVNUM, VMNAME ASC"
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then

                    For Each row As DataRow In ds.Tables(0).Rows
                        lstCollection.Add(row(0).ToString())
                    Next
                    Return lstCollection
                End If
            End If
            Return Nothing
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetReferencesInProject(projectCode As Integer) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim amount As Integer = -1
        Try

            Sql = "Select COUNT(PRHCOD) FROM PRDVLD WHERE PRHCOD = " & projectCode
            ds = GetDataFromDatabase(Sql)

            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    amount = CInt(ds.Tables(0).Rows(0).ItemArray(0).ToString())
                End If
            End If
            Return amount
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return amount
        End Try
    End Function

    Public Function GetVendorInProject(projectCode As Integer) As List(Of Integer)
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim vendorNo As Integer = 0
        Dim lstVendors = New List(Of Integer)()
        Try
            Sql = "Select DISTINCT VMVNUM FROM PRDVLD WHERE PRHCOD = " & projectCode
            ds = GetDataFromDatabase(Sql)

            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 And ds.Tables(0).Rows.Count < 2 Then
                    vendorNo = CInt(ds.Tables(0).Rows(0).ItemArray(0).ToString())
                    lstVendors.Add(vendorNo)
                Else
                    For Each dr As DataRow In ds.Tables(0).Rows
                        lstVendors.Add(CInt(dr.ItemArray(0).ToString()))
                    Next
                End If
            End If
            Return lstVendors
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return lstVendors
        End Try
    End Function

    Public Function GetPartInProdDesc(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRDPTN FROM PRDVLD WHERE TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartInImnsta(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from inmsta where trim(ucase(imptn)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartInDvinva(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select dvpart from dvinva where TRIM(dvpart) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartInCater(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT CATPTN FROM CATER where TRIM(catptn) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartInKomat(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT KOPTNO FROM KOMAT where TRIM(koptno) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getItemCategoryByVendorAndPart(vendorNo As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Dim Sql = "SELECT PRHCOD FROM PRDVLD WHERE VMVNUM = " & Trim(vendorNo) & " And trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
        Try
            ds = FillGrid(Sql)
            If ds IsNot Nothing Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Function

    Public Function getVendorTypeByVendorNum(vendorNo As String) As String
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Dim Sql = "select vmvtyp, vmname from vnmas where vmvnum = " & vendorNo & " "
        Try
            ds = FillGrid(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds.Tables(0).Rows(0).ItemArray(0).ToString()
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getVendorTypeByVendorNum(vendorNo As String, Optional ByVal flag As Integer = 0) As Data.DataSet
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Dim Sql = "select vmvtyp, vmname from vnmas where vmvnum = " & vendorNo & " "
        Try
            ds = FillGrid(Sql)
            If ds IsNot Nothing Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getOEMVendorCodes(cntrCode As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Dim Sql = "select CNTDE1 from cntrll where cnt01 = " & cntrCode & " "
        Try
            ds = FillGrid(Sql)
            If ds IsNot Nothing Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPRHCODInDetails(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRDDAT,Trim(PRDPTN) as PRDPTN,Trim(PRDCTP) as PRDCTP,Trim(PRDMFR#) as PRDMFR#,Trim(PRDVLD.VMVNUM) as VMVNUM,
                    Trim(VMNAME) as VMNAME,Trim(PRDSTS) as PRDSTS,Trim(PRDJIRA) as PRDJIRA,Trim(PRDUSR) as PRDUSR FROM PRDVLD 
                    INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetExistByPRNAME(name As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRHCOD FROM PRDVLH WHERE PRNAME = '" & Trim(name) & "' ORDER BY 1 DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByVendorAndPartNoProdDesc(partNo As String, vendorNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select PRHCOD from prdvld where TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "' and vmvnum = " & Trim(vendorNo) & " order by 1 desc"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByVendorAndPartNoDevPoq(partNo As String, vendorNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select A1.prhcod from prdvld A1 inner join poqota A2 on A1.prdptn = A2.pqptn and A1.vmvnum = A2.pqvnd
                    where SUBSTR(UCASE(A2.SPACE),32,3) = 'DEV' and  A2.PQCOMM LIKE 'D%' and A2.pqcomm <> 'D-'
                    and TRIM(A1.PRDPTN) = '" & Trim(UCase(partNo)) & "' and A1.vmvnum = " & Trim(vendorNo) & "
                    order by 1 desc "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetCodeAndNameByPartNo(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRDVLH.PRHCOD,PRDVLH.PRNAME,PRDVLD.VMVNUM FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD WHERE TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "' ORDER BY PRDVLD.CRDATE DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNoVendor(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select IMDSC,impc2,impc1 from inmsta where trim(ucase(imptn)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'Log.Error(exMessage)
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetAllStatuses() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT CNT03, CNTDE1 FROM cntrll where cnt01 = 'DSI' order by cnt02"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetAllStatusesReturn(strValue As String, strColumn As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        Dim Qry As New DataTable
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT cnt03, cntde1 FROM cntrll where cnt01 = 'DSI' order by cnt02"
            ds = GetDataFromDatabase(Sql)

            Dim Qry1 = ds.Tables(0).AsEnumerable() _
                          .Where(Function(x) Trim(UCase(x.Field(Of String)(strColumn))) = Trim(UCase(strValue)))
            If Qry1.Count > 0 Then
                Qry = Qry1.CopyToDataTable
                Dim result = Qry(0).Item("CNT03").ToString()
                Return result
            Else
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPOQotaData(vendorNo As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture

        Try
            Sql = "SELECT PQMPTN,PQPRC,PQSEQ,PQPTN,PQVND,PQMIN,PQCOMM FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' 
                    AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataFromDualInventory(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT A2.DVPRMG,A1.IMPC1,A1.IMPC2,A1.IMDSC FROM INMSTA A1 INNER JOIN DVINVA A2 
                                        ON A1.IMPTN = A2.DVPART WHERE A2.DVLOCN = '01' AND UCASE(A1.IMPTN) = '" & UCase(partNo) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetEmailData(flag As Integer) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            If flag = 1 Then
                Sql = "select cntde1 from cntrll where cnt01 = 'SLS' and cnt03 = 'MGR'"
            Else
                Sql = "select cntde1 from cntrll where cnt01 = 'MKT' "
            End If

            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function CallForCtpNumber(partno As String, ctppartno As String, flagctp As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "CALL CTPINV.CATCTPR ('" & partno & "','" & ctppartno & "','" & flagctp & "')"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function FillDDLUser() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select distinct A1.ususer, A1.usname from csuser A1 inner join prdvld A2 on A1.ususer = A2.prdusr WHERE USPTY8 = 'X' AND USPTY9 <> 'R' ORDER BY USNAME"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function FillDDlMinorCode() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT CNT03, CNTDE1 FROM CNTRLL WHERE CNT01 = '120' ORDER BY TRIM(CNTDE1) "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function FillDDlMajorCode() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT CNT03, CNTDE1 FROM CNTRLL WHERE CNT01 = '110' ORDER BY TRIM(CNTDE1) "
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getFlagAllow(userid As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim rsValue As Integer = -1
        Try
            Sql = "SELECT CNT01 FROM CNTRLL WHERE TRIM(CNT01) = '988' AND TRIM(CNTDE1) = Trim (UCase('" & userid & "'))"
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    rsValue = 1
                End If
            End If
            Return rsValue
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsValue
        End Try
    End Function

    Public Function getVendorsAccepted(ds As DataSet) As DataSet
        Dim exMessage As String = " "
        Try
            Dim dsRsult = New DataSet()
            Dim dtRsult = New DataTable()
            dtRsult.Columns.Add("VMVNUM", GetType(Integer))
            dtRsult.Columns.Add("VMNAME", GetType(String))
            dtRsult.Columns.Add("vmvtyp", GetType(String))
            dsRsult.Tables.Add(dtRsult)

            If ds IsNot Nothing Then

                For Each dw As DataRow In ds.Tables(0).Rows

                    Dim vendorType = dw.ItemArray(2).ToString()
                    Dim vendorName = dw.ItemArray(1).ToString()
                    Dim vendorNo = dw.ItemArray(0).ToString()
                    Dim listDeniedCodes = VendorCodesDenied.Split(",")
                    Dim containsDenied = listDeniedCodes.AsEnumerable().Any(Function(x) x = vendorType)
                    If Not containsDenied Then
                        Dim OEMContain = getOEMVendorCodes(VendorOEMCodeDenied)
                        Dim containsOEM = OEMContain.Tables(0).AsEnumerable().Any(Function(x) Trim(x.ItemArray(0).ToString()) = Trim(vendorNo))
                        If Not containsOEM Then

                            Dim newRow As DataRow = dsRsult.Tables(0).NewRow
                            newRow("VMVNUM") = vendorNo
                            newRow("VMNAME") = vendorName
                            newRow("vmvtyp") = vendorType
                            dsRsult.Tables(0).Rows.Add(newRow)

                            'mustDelete = True
                            'ds.Tables(0).Rows.RemoveAt(i)
                            'frmLoadExcel.lblVendorDesc.Text = vendorName
                            'MessageBox.Show("The vendor " & RTrim(vendorName) & " is an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                            'Return True
                        End If
                    End If
                Next
                Return dsRsult
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function isVendorAccepted(vendorNo As String) As Boolean
        Dim exMessage As String = " "
        Try
            'Dim vendorType = getVendorTypeByVendorNum(vendorNo)
            Dim ds As DataSet = getVendorTypeByVendorNum(vendorNo, 0)
            If ds IsNot Nothing Then
                Dim vendorType = ds.Tables(0).Rows(0).ItemArray(0).ToString()
                Dim vendorName = ds.Tables(0).Rows(0).ItemArray(1).ToString()
                Dim listDeniedCodes = VendorCodesDenied.Split(",")
                Dim ExcludedVendors = VendorOEMExclude.ToString()
                Dim containsDenied = listDeniedCodes.AsEnumerable().Any(Function(x As String) x = "'" & vendorType & "'")
                If Not containsDenied Then
                    Dim OEMContain = getOEMVendorCodes(VendorOEMCodeDenied)
                    Dim containsOEM = OEMContain.Tables(0).AsEnumerable().Any(Function(x) (Trim(x.ItemArray(0).ToString()) = Trim(vendorNo)) And x.ItemArray(0).ToString() <> ExcludedVendors)
                    If Not containsOEM Then
                        frmLoadExcel.lblVendorDesc.Text = vendorName
                        'MessageBox.Show("The vendor " & RTrim(vendorName) & " is an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                        Return True
                    Else
                        MessageBox.Show("The vendor " & RTrim(vendorName) & " is not an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                        Return False
                    End If
                Else
                    MessageBox.Show("The vendor " & RTrim(vendorName) & " is not an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                    Return False
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return False
        End Try
    End Function

    Public Function customIsVendorAccepted(vendorNo As String) As Boolean
        Dim exMessage As String = " "
        Try
            'Dim vendorType = getVendorTypeByVendorNum(vendorNo)
            Dim ds As DataSet = getVendorTypeByVendorNum(vendorNo, 0)
            If ds IsNot Nothing Then
                Dim vendorType = ds.Tables(0).Rows(0).ItemArray(0).ToString()
                Dim vendorName = ds.Tables(0).Rows(0).ItemArray(1).ToString()
                Dim listDeniedCodes = VendorCodesDenied.Split(",")
                Dim ExcludedVendors = VendorOEMExclude.Split(",")
                Dim containsDenied = listDeniedCodes.AsEnumerable().Any(Function(x As String) x = "'" & vendorType & "'")
                If Not containsDenied Then
                    Dim OEMContain = getOEMVendorCodes(VendorOEMCodeDenied)
                    'Dim firstFilter = OEMContain.Tables(0).AsEnumerable().Any(Function(x) Not ExcludedVendors.Contains(x.ItemArray(0).ToString()))
                    'If Not firstFilter Then
                    Dim containsOEM = OEMContain.Tables(0).AsEnumerable().Any(Function(x) (Trim(x.ItemArray(0).ToString()) = Trim(vendorNo)) And (Not ExcludedVendors.Contains(vendorNo)))
                    If Not containsOEM Then
                        'frmLoadExcel.lblVendorDesc.Text = vendorName
                        'MessageBox.Show("The vendor " & RTrim(vendorName) & " is an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                        Return True
                    Else
                        'MessageBox.Show("The vendor " & RTrim(vendorName) & " is not an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                        Return False
                    End If
                    'End If
                Else
                    'MessageBox.Show("The vendor " & RTrim(vendorName) & " is not an accepted vendor for the operation.", "CTP System", MessageBoxButtons.OK)
                    Return False
                End If
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return False
        End Try
    End Function

    Public Function isPartInExistence(partNo As String) As Boolean
        'check for part number inm imnsta, cater y komat
        Dim ds1 = New DataSet()
        Dim ds2 = New DataSet()
        Dim ds3 = New DataSet()
        Dim ds4 = New DataSet()
        Dim ds5 = New DataSet()
        Dim exMessage As String = " "

        Try
            ds5 = GetPartInImnsta(partNo)
            If ds5 Is Nothing Then
                ds1 = GetPartInProdDesc(partNo)
                If ds1 Is Nothing Then
                    ds2 = GetPartInDvinva(partNo)
                    If ds2 Is Nothing Then
                        ds3 = GetPartInCater(partNo)
                        If ds3 Is Nothing Then
                            ds4 = GetPartInKomat(partNo)
                            If ds4 Is Nothing Then
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return False
        End Try

    End Function


#End Region

    Public Function checkPurcByUser(userid As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Dim rsValue As Integer = -1
        Try
            Sql = "SELECT * FROM CSUSER WHERE USUSER = '" & userid & "' AND USPURC <> 0 "
            ds = GetDataFromDatabase(Sql)
            If ds IsNot Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    rsValue = If(ds.Tables(0).Rows(0).Item("USPURC").ToString() IsNot Nothing, ds.Tables(0).Rows(0).Item("USPURC").ToString(), -1)
                End If
            End If
            Return rsValue
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsValue
        End Try
    End Function

    Public Function GetDataByPRHCOD(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLH WHERE PRHCOD = " & Trim(code)
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndPartNo(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & Trim(code) & " AND trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndPartNoProdDesc(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD WHERE PRHCOD = " & Trim(code) & " AND trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndPartNoProdDesc1(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD WHERE PRHCOD = " & Trim(code) & " AND trim(ucase(PRDPTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetProdDetByCodeAndExc(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from prdvld where prhcod = " & Trim(code) & " and prdsts <> 'CS' and prdsts <> 'CN' and prdsts <> 'CD' and prdsts <> 'CL'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodeAndVendorAndPart(code As String, vendorNo As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD WHERE PRHCOD = " & Trim(code) & " and VMVNUM = " & Trim(vendorNo) & " AND TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodAndPartProdAndComm(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDCMH INNER JOIN PRDCMD ON PRDCMH.PRDCCO = PRDCMD.PRDCCO WHERE PRDCMH.PRHCOD = " & code & " AND prdcmh.PRDPTN = '" & Trim(UCase(partNo)) & "' and trim(ucase(prdctx)) like '%QUOTING%'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodAndPartProdAndComm1(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDCMH WHERE PRDCMH.PRHCOD = " & code & " AND prdcmh.PRDPTN = '" & Trim(UCase(partNo)) & "' ORDER BY  PRDCDA DESC,PRDCTI DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodAndPartProdAndComm2(code As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT PRDCMH.*,PRDCMD.* FROM PRDCMH INNER JOIN PRDCMD ON PRDCMH.PRDCCO = PRDCMD.PRDCCO WHERE PRDCMH.PRHCOD = " & Trim(code) & " AND PRDCMH.PRDPTN = '" & Trim(partNo) & "' ORDER BY  PRDCDA ASC,PRDCTI ASC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByCodAndPartProdAndComm2(tableCode As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDCMD WHERE PRDCCO = " & tableCode
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNo2(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM DVINVA INNER JOIN VNMAS ON DVINVA.DVPRMG = digits(VNMAS.VMVNUM) WHERE DVPART = '" & Trim(UCase(partNo)) & "' and dvlocn = '01'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartVendor(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM DVINVA WHERE DVPART = '" & Trim(UCase(partNo)) & "' and dvlocn = '01'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartMix() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM CNTRLL WHERE CNT01 = '120' ORDER BY TRIM(CNTDE1)"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartMix1() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from cntrll where cnt01 = '102'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByVendorAndPartNo(vendorNo As String, partNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim strDescrption As String
        Dim columnToChange = "PQMIN"
        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            strDescrption = GetSingleDataFromDatabase(Sql, columnToChange)
            Return strDescrption
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByVendorAndPartNoDst(vendorNo As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataByPartNo(partNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim strDescrption As String
        Dim columnToChange = "IMDSC"
        Try
            Sql = "SELECT * FROM INMSTA INNER JOIN DVINVA ON INMSTA.IMPTN = DVINVA.DVPART WHERE UCASE(IMPTN) = '" & Trim(UCase(partNo)) & "'"
            strDescrption = GetSingleDataFromDatabase(Sql, columnToChange)
            Return strDescrption
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetJiraPath() As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim JiraPath As String = " "
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM cntrll where cnt01 = 'JIR' and trim(ucase(cnt03)) = 'PRO'"
            ds = GetDataFromDatabase(Sql)
            If ds.Tables(0).Rows.Count = 1 Then
                JiraPath = ds.Tables(0).Rows(0).Item(3).ToString()
            End If
            Return JiraPath
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPOQotaDataDuplex(strWhereAdd As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM POQOTA " & strWhereAdd & "  AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%' ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetProjectStatusDescription(code As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ProjectDescStatus As String = " "
        Dim columnToChange = "CNTDE1"
        Try
            Dim CodeOk As String = Trim(UCase(code))
            Sql = "SELECT * FROM cntrll where cnt01 = 'DSI' and cnt03 = '" & CodeOk & "'"
            ProjectDescStatus = GetSingleDataFromDatabase(Sql, columnToChange)
            Return Trim(ProjectDescStatus)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetWLDataByPartNo(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDWL WHERE TRIM(UCASE(WHLPARTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetVendorByVendorNo(vendorNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM VNMAS WHERE VMVNUM = " & Trim(UCase(vendorNo))
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetVendorQuey(variable As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM VNMAS WHERE DIGITS(VMVNUM) = " & variable
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetVendorByName(vendorName As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM VNMAS WHERE TRIM(UCASE(VMNAME)) LIKE '" & Trim(UCase(vendorName)) & "%'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartCtpRef(ctpNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim TcpPartNo As String = " "
        Dim columnToChange = "CRPTNO"
        Try
            Sql = "SELECT * FROM CTPREFS WHERE TRIM(UCASE(CRCTPR )) = '" & Trim(UCase(ctpNo)) & "'"
            TcpPartNo = GetSingleDataFromDatabase(Sql, columnToChange)
            Return Trim(TcpPartNo)
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetCTPPartRef(partNo As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim TcpPartNo As String = " "
        Dim columnToChange = "CRCTPR"
        Try
            Sql = "SELECT * FROM CTPREFS WHERE TRIM(UCASE(CRPTNO)) = '" & Trim(UCase(partNo)) & "'"
            TcpPartNo = GetSingleDataFromDatabase(Sql, columnToChange)
            Return Trim(TcpPartNo)
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataFromProdHeaderAndDetail(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD WHERE TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "' ORDER BY PRDVLD.CRDATE DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetDataFromProdHeaderAndDetail2(code As String, partNo As String, vendorNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLH INNER JOIN PRDVLD ON PRDVLH.PRHCOD = PRDVLD.PRHCOD WHERE PRDVLD.VMVNUM = " & Trim(vendorNo) & " AND TRIM(PRDPTN) = '" & Trim(UCase(partNo)) & "' AND PRDVLH.prhcod <> " & Trim(code) & " AND PRDSTS <> 'CS' AND PRDSTS <> 'CN' AND PRDSTS <> 'CL'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetPartInWishList(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from prdwl where TRIM(UCASE(WHLPARTN)) = '" & Trim(UCase(partNo)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetAssignedVendor(vendorAssigned As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorAssigned) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' and pqqdty < 50 
                    ORDER BY PQQDTY DESC, PQQDTM DESC, PQQDTD DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetAllPOQOTA(vendorAssigned As String, partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM POQOTA WHERE PQVND = " & Trim(vendorAssigned) & " AND PQPTN = '" & Trim(UCase(partNo)) & "'  ORDER BY PQSEQ DESC"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString() + ". " + ex.Message + ". " + ex.ToString()
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getUserDataByUsername(userName As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM CSUSER WHERE USUSER = '" & Trim(UCase(userName)) & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getMarketingDataByDate() As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM MACALE WHERE MACADY > 0 AND MACABD >= '" & Format(Now, "yyyy-mm-dd") & "'"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetMenuByUser(userid As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select AMENUCTP.*,DETMENUCTP.dmdimain from MENUCTP inner join DETMENUCTP 
                    on MENUCTP.CODMENU = DETMENUCTP.CODMENU inner join AMENUCTP on AMENUCTP.CODMENU = MENUCTP.CODMENU 
                    where userid = '" & userid & "' and DETMENUCTP.CODDETMENU = AMENUCTP.CODDETMENU order by AMENUCTP.CODMENU"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetInvProdDetailByProject(code As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "SELECT * FROM PRDVLD INNER JOIN INMSTA ON TRIM(PRDVLD.PRDPTN) = TRIM(INMSTA.IMPTN) WHERE PRHCOD = " & code & " ORDER BY PRDPTN"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

#End Region

#Region "Inserts"

    Public Function InsertNewProject(projectno As String, userid As String, dtValue As DateTimePicker, strInfo As String, strName As String, ddlStatus As ComboBox, strUser As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDVLH(PRHCOD,CRUSER,CRDATE,PRDATE,PRINFO,PRNAME,PRSTAT,MOUSER,MODATE,PRPECH) VALUES 
            (" & projectno & ",'" & userid & "','" & Format(Now, "yyyy-MM-dd") & "','" & Format(dtValue.Value, "yyyy-MM-dd") & "',
            '" & Trim(strInfo) & "', '" & Trim(strName) & "','" & Left(ddlStatus.SelectedItem.ToString(), 1) & "','" & userid & "',
            '" & Format(Now, "yyyy-MM-dd") & "','" & Left(Trim(strUser), 10) & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductComment(code As String, partNo As String, comment As String, userId As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDCMH(PRHCOD,PRDPTN,PRDCCO,PRDCDA,PRDCTI,PRDCSU,USUSER) 
                    VALUES(" & Trim(code) & ",'" & Trim(partNo) & "'," & comment & ",'" & Format(DateTime.Now, "yyyy-MM-dd") & "','" & Format(DateTime.Now, "hh:mm:ss") & "',
                            'Person in charge changed','" & userId & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductCommentNew(code As String, partNo As String, comment As String, commentSubject As String, userId As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDCMH(PRHCOD,PRDPTN,PRDCCO,PRDCDA,PRDCTI,PRDCSU,USUSER) 
                    VALUES(" & Trim(code) & ",'" & Trim(partNo) & "'," & comment & ",'" & Format(DateTime.Now, "yyyy-MM-dd") & "',
                    '" & Format(DateTime.Now, "hh:mm:ss") & "','" & commentSubject & "','" & userId & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductCommentDetail(code As String, partNo As String, comment As String, cod_detcomment As String, messcomm As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDCMD(PRHCOD,PRDPTN,PRDCCO,PRDCDC,PRDCTX) 
                    VALUES(" & Trim(code) & ",'" & Trim(partNo) & "'," & comment & "," & cod_detcomment & ",'" & messcomm & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertWishListProduct(maxItem As String, userId As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO PRDWL(WHLCODE,WHLUSER,WHLDATE,WHLPARTN,WHLREASONT,WHLCOMMENT)
                    VALUES(" & maxItem & ",'" & userId & "','" & Format(Now, "yyyy-mm-dd") & "','" & Trim(UCase(partNo)) & "','','No vendor assigned')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewPOQota(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String, strUnitCostNew As String, strMinQty As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(strStsQuote), strStsQuote, If(strStsQuote.Length < maxLength, strStsQuote, strStsQuote.Substring(0, Math.Min(strStsQuote.Length, maxLength))))

            Sql = "INSERT INTO POQOTA (PQPTN,PQVND,PQSEQ,PQQDTY,PQQDTM,PQMPTN,PQQDTD,PQCOMM,SPACE,PQPRC,PQMIN) VALUES 
            ('" & Trim(UCase(partNo)) & "'," & Trim(vendorNo) & "," & maxValue & "," & strYear.Substring(strYear.Length - 2) & ",
            " & strMonth & ",'" & mpnPo & "'," & strDay & ",'" & statusquoteNew & "','" & strSpace & "'," & strUnitCostNew & "," & strMinQty & ")"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewPOQota1(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(strStsQuote), strStsQuote, If(strStsQuote.Length < maxLength, strStsQuote, strStsQuote.Substring(0, Math.Min(strStsQuote.Length, maxLength))))

            Sql = "INSERT INTO POQOTA (PQPTN,PQVND,PQSEQ,PQQDTY,PQQDTM,PQMPTN,PQQDTD,PQCOMM,SPACE) VALUES 
            ('" & Trim(UCase(partNo)) & "'," & Trim(vendorNo) & "," & maxValue & "," & strYear.Substring(strYear.Length - 2) & ",
            " & strMonth & ",'" & mpnPo & "'," & strDay & ",'" & statusquoteNew & "','" & strSpace & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewPOQotaLess(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String, strUnitCostNew As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(strStsQuote), strStsQuote, If(strStsQuote.Length < maxLength, strStsQuote, strStsQuote.Substring(0, Math.Min(strStsQuote.Length, maxLength))))

            Sql = "INSERT INTO POQOTA (PQPTN,PQVND,PQSEQ,PQQDTY,PQQDTM,PQMPTN,PQQDTD,PQCOMM,SPACE,PQPRC) VALUES 
            ('" & Trim(UCase(partNo)) & "'," & Trim(vendorNo) & "," & maxValue & "," & strYear.Substring(strYear.Length - 2) & ",
            " & strMonth & ",'" & mpnPo & "'," & strDay & ",'" & statusquoteNew & "','" & strSpace & "'," & strUnitCostNew & ")"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductDetail(projectno As String, partNo As String, dtValue As DateTimePicker, userid As String, dtValue1 As DateTimePicker, userid1 As String, dtValue2 As DateTimePicker, ctpNo As String, qty As String,
                                        mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, dtValue3 As DateTimePicker, ddlStatus As String, benefits As String,
                                        comments As String, ddlUser As String, chkNew As CheckBox, dtValue4 As DateTimePicker, sampleCost As String, miscCost As String, vendorNo As String,
                                        partsToShow As String, ddlMinorCode As String, toolingCost As String, dtValue5 As DateTimePicker, dtValue6 As DateTimePicker, sampleQty As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            'dtValue6.Value = New DateTime(strDate)
            'Dim chkSelection1 As Integer = If(chkNew.Checked = False, 0, 1)
            Dim chkSelection As Integer = If(getValueCheckTab3(vendorNo, partNo) = -1, 0, 1)

            Sql = "INSERT INTO PRDVLD(PRHCOD,PRDPTN,PRDDAT,CRUSER,CRDATE,MOUSER,MODATE,PRDCTP,PRDQTY,PRDMFR,PRDMFR#,PRDCOS,PRDCON,PRDPO#,PODATE,PRDSTS,PRDBEN,PRDINF,PRDUSR,PRDNEW,
                                        PRDEDD,PRDSCO,PRDTTC,VMVNUM,PRDPTS,PRDMPC,PRDTCO,PRDERD,PRDPDA,PRDSQTY) 
                   VALUES (" & projectno & ",'" & Trim(UCase(partNo)) & "','" & Format(dtValue.Value, "yyyy-MM-dd") & "','" & userid & "','" & Format(dtValue1.Value, "yyyy-MM-dd") & "',
                    '" & userid & "','" & Format(dtValue2.Value, "yyyy-MM-dd") & "','" & Trim(ctpNo) & "'," & qty & ",
            '" & Trim(mfr) & "','" & Trim(mfrNo) & "'," & (unitCost) & ",
                    " & (unitCostNew) & ",'" & Trim(poNo) & "','" & Format(dtValue3.Value, "yyyy-MM-dd") & "',
            '" & Trim(ddlStatus) & "','" & Trim(benefits) & "','" & Trim(comments) & "',
                    '" & Trim(ddlUser) & "'," & chkSelection & ",'" & Format(dtValue4.Value, "yyyy-MM-dd") & "'," & sampleCost & "," & miscCost & "," & Trim(vendorNo) & ",
            '" & partsToShow & "',
                    '" & (ddlMinorCode) & "'," & toolingCost & ",'" & Format(dtValue5.Value, "yyyy-MM-dd") & "', '" & Format(dtValue6.Value, "yyyy-MM-dd") & "' ," & sampleQty & ")"

            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertProductDetailv2(projectno As String, partNo As String, dtValue As DateTimePicker, userid As String, dtValue1 As DateTimePicker, userid1 As String, dtValue2 As DateTimePicker, ctpNo As String, qty As String,
                                        mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, dtValue3 As DateTimePicker, ddlStatus As String, benefits As String,
                                        comments As String, ddlUser As String, chkNew As CheckBox, dtValue4 As DateTimePicker, sampleCost As String, miscCost As String, vendorNo As String,
                                        partsToShow As String, ddlMinorCode As String, toolingCost As String, dtValue5 As DateTimePicker, dtValue6 As DateTimePicker, dtValue7 As DateTimePicker, newValue2 As String) As String
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            'dtValue6.Value = New DateTime(strDate)
            Dim chkSelection As Integer = If(chkNew.Checked = False, 0, 1)

            Sql = "INSERT INTO PRDVLD(PRHCOD,PRDPTN,PRDDAT,CRUSER,CRDATE,MOUSER,MODATE,PRDCTP,PRDQTY,PRDMFR,PRDMFR#,PRDCOS,PRDCON,PRDPO#,PODATE,PRDSTS,PRDBEN,PRDINF,PRDUSR,PRDNEW,
                                        PRDEDD,PRDSCO,PRDTTC,VMVNUM,PRDPTS,PRDMPC,PRDTCO,PRDERD,PRDPDA,PRWLDA,PRWLFL) 
                    VALUES(" & Trim(projectno) & ",'" & Trim(UCase(partNo)) & "','" & Format(Now, "yyyy-MM-dd") & "','" & userid & "','" & Format(Now, "yyyy-MM-dd") & "','" & userid1 & "',
                    '" & Format(Now, "yyyy-MM-dd") & "','" & Trim(ctpNo) & "',0,'',''," & unitCost & ",0,'','1900-01-01','E','','','" & userid & "',0,'1900-01-01',0,0," & Trim(vendorNo) & ",'',
                    '" & ddlMinorCode & "',0,'1900-01-01','1900-01-01','" & Format(dtValue7.Value, "yyyy-MM-dd") & "',1)"

            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertNewInv(strdvlocn As String, strdvpart As String, strdvmjpc As String, strdvmnpc As String, strdvindt As Decimal, strdvunt As String, strdvslr As String, strdvohr As String,
                                    dvprmg As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "insert into dvinva(dvlocn,dvpart,dvmjpc,dvmnpc,dvindt,dvunt$,dvslr,dvohr,dvprmg) 
                    values('01','" & Trim(UCase(strdvpart)) & "','" & Trim(UCase(strdvmjpc)) & "','" & Trim(UCase(strdvmnpc)) & "'," & strdvindt & ",
                            " & strdvunt & ",'99999','99999','" & Trim(dvprmg) & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertIntoLoginTcp(codloginctp As String, userid As String, Versionctp As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "INSERT INTO LOGINCTP VALUES(" & codloginctp & ",'" & userid & "','" & Format(Now, "yyyy-MM-dd") &
                        "','" & Format(Now, "hh:MM:ss") & "','" & Versionctp & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function InsertIntoLogProdDetail(userid As String, code As String, partNo As String, prdDat As DateTime, cruser As String, crDate As DateTime, mouser As String, moDate As DateTime,
                                            ctpNo As String, qty As String, mfrProd As String, mfrProdNo As String, prdCos As String, prdCon As String, poNo As String,
                                            poDate As DateTime, status As String, benefits As String, info As String, prdUsr As String, chkNew As String, prdEdd As DateTime,
                                            sampleCost As String, miscCost As String, vendorNo As String, prdPts As String, prdMpc As String, toolCost As String, prdErd As DateTime,
                                            prdPda As DateTime, prdsQty As String, prwLda As DateTime, prwLfl As String, partNoo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "insert into logprdvld(LGUSRDEL, LGDELDATE, LGTRSTATUS, LGPRHCOD, LGPRDPTN,LGPRDDAT, LGCRUSER, LGCRDATE, LGMOUSER, LGMODATE, 
                    LGPRDCTP,LGPRDQTY, LGPRDMFR, LGPRDMFR#, LGPRDCOS, LGPRDCON, LGPRDPO#,LGPODATE, LGPRDSTS, LGPRDBEN, LGPRDINF, LGPRDUSR, 
                    LGPRDNEW,LGPRDEDD, LGPRDSCO, LGPRDTTC, LGVMVNUM, LGPRDPTS, LGPRDMPC,LGPRDTCO, LGPRDERD, LGPRDPDA, LGPRDSQTY, LGPRWLDA, LGPRWLFL,LGPARTNO) 
                    values('" & userid & "','" & Format(Now, "yyyy-MM-dd") & "','DEL'," & code & ",'" & Trim(partNo) & "','" & Format(prdDat, "yyyy-MM-dd") & "',
                    '" & Trim(cruser) & "','" & Format(crDate, "yyyy-MM-dd") & "','" & Trim(mouser) & "','" & Format(moDate, "yyyy-MM-dd") & "','" & Trim(ctpNo) & "'," & qty & ",'" & Trim(mfrProd) &
                "','" & Trim(mfrProdNo) & "'," & prdCos & "," & prdCon & ",'" & Trim(poNo) & "','" & Format(poDate, "yyyy-MM-dd") & "','" & Trim(status) & "','" & Trim(benefits) & "',
                '" & Trim(info) & "','" & Trim(prdUsr) & "'," & chkNew & ",'" & Format(prdEdd, "yyyy-MM-dd") & "'," & sampleCost & "," & miscCost & "," & vendorNo & ",'" & Trim(prdPts) & "',
                '" & Trim(prdMpc) & "'," & toolCost & ",'" & Format(prdErd, "yyyy-MM-dd") & "','" & Format(prdPda, "yyyy-MM-dd") & "'," & prdsQty & ",'" & Format(prwLda, "yyyy-MM-dd") & "'," & prwLfl & ",'" & Trim(partNoo) & "')"
            QueryResult = InsertDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function


#End Region

#Region "Updates"

    Public Function UpdateGeneralStatus(code As String, status As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Sql = "update qs36f.prdvlh set prstat = '" & status(0) & "' where prhcod = " & Trim(code) & ""
            QueryResult = UpdateDataInDatabase1(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQotaExact(statusquote As String, insertYear As String, insertMonth As String, insertDay As String, vendorNo As String, partNo As String, secuencial As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(statusquote), statusquote, If(statusquote.Length < maxLength, statusquote, statusquote.Substring(0, Math.Min(statusquote.Length, maxLength))))

            Sql = "UPDATE POQOTA SET PQCOMM = '" & statusquoteNew & "', PQQDTY =  " & insertYear.Substring(insertYear.Length - 2) & " ,PQQDTM = " & insertMonth & " ,
                    PQQDTD = " & insertDay & " WHERE PQSEQ = " & secuencial & " AND PQVND  = " & Trim(vendorNo) & " AND PQPTN  = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' " &
                    " AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraRow(mpnopo As String, minQty As String, unitCostNew As String, statusquote As String, insertYear As String, insertMonth As String, insertDay As String,
                                    vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(statusquote), statusquote, If(statusquote.Length < maxLength, statusquote, statusquote.Substring(0, Math.Min(statusquote.Length, maxLength))))

            Sql = "UPDATE POQOTA SET PQMPTN = '" & mpnopo & "',PQMIN  = " & minQty & ",PQPRC  = " & unitCostNew & ",PQCOMM = '" & statusquoteNew & "',
                PQQDTY =  " & insertYear.Substring(insertYear.Length - 2) & " ,PQQDTM = " & insertMonth & " ,PQQDTD = " & insertDay & " 
                WHERE PQVND  = " & Trim(vendorNo) & " AND PQPTN  = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' " &
                " AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraRowNew(statusquote As String, insertYear As String, insertMonth As String, insertDay As String, vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(statusquote), statusquote, If(statusquote.Length < maxLength, statusquote, statusquote.Substring(0, Math.Min(statusquote.Length, maxLength))))

            Sql = "UPDATE POQOTA SET PQCOMM = '" & statusquoteNew & "', PQQDTY =  " & insertYear.Substring(insertYear.Length - 2) & " ,PQQDTM = " & insertMonth & " ,
                    PQQDTD = " & insertDay & " WHERE PQVND  = " & Trim(vendorNo) & " AND PQPTN  = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' " &
                    " AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraRowVendor(oldVendorNo As String, newVendorNo As String, partNo As String, poQotaSeq As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE POQOTA SET PQVND = " & newVendorNo & ", PQPRC = 0 WHERE PQVND = " & oldVendorNo & " AND
                    PQPTN = '" & Trim(UCase(partNo)) & "' AND PQSEQ = " & poQotaSeq & " AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraMfr(mpnopo As String, vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE POQOTA SET PQMPTN = '" & Trim(UCase(mpnopo)) & "' WHERE PQVND = " & vendorNo & " 
                    AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraUC(unitCost As String, vendorNo As String, partNo As String, strYear As String, strMonth As String, strDay As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE POQOTA SET PQPRC = '" & Trim(UCase(unitCost)) & "', PQQDTY = " & strYear.Substring(2, 2) & ", PQQDTM = " & strMonth & ", PQQDTD = " & strDay & "  
                    WHERE PQVND = " & vendorNo & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQoraRow1(mpnopo As String, statusquote As String, insertYear As String, insertMonth As String, insertDay As String,
                                    vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Dim maxLength As Integer = 20
        Try
            Dim statusquoteNew = If(String.IsNullOrEmpty(statusquote), statusquote, If(statusquote.Length < maxLength, statusquote, statusquote.Substring(0, Math.Min(statusquote.Length, maxLength))))

            Sql = "UPDATE POQOTA SET PQMPTN = '" & mpnopo & "',PQCOMM = '" & statusquoteNew + "NEW" & "',
                PQQDTY =  " & insertYear.Substring(insertYear.Length - 2) & " ,PQQDTM = " & insertMonth & " ,PQQDTD = " & insertDay & " 
                WHERE PQVND  = " & Trim(vendorNo) & " AND PQPTN  = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' " &
                " AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdatePoQotaByVendorAndPart(vendorNo As String, oldVendorNo As String, partNo As String, pqSeq As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE POQOTA SET PQVND = " & vendorNo & ", PQPRC = 0 WHERE PQVND = " & oldVendorNo & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND 
                    PQSEQ = " & pqSeq & " AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail(code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPDA = '" & Format(Now, "yyyy-MM-dd") & "' WHERE PRHCOD = " & Trim(code) & " AND PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail1(partstoshow As String, minorCode As String, tooCost As String, strDate1 As Date, jiraTask As String, vendorNo As String, strChkSel As String,
                                        strDate2 As Date, sampleCost As String, miscCost As String, userSelec As String, strDate3 As Date, userid As String, tcpNo As String, sampleQty As String,
                                        qty As String, mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, strDate4 As Date, status As String,
                                        benefits As String, comments As String, code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Dim checkDate1 = Format(strDate1, "yyyy/MM/dd")
            Dim checkDate2 = Format(strDate2, "yyyy/MM/dd")
            Dim checkDate3 = Format(strDate3, "yyyy/MM-dd")
            Dim checkDate4 = Format(Now, "yyyy-MM-dd")



            Sql = "UPDATE PRDVLD SET PRDPTS = '" & partstoshow & "',PRDMPC = '" & minorCode & "',PRDTCO = " & tooCost & ",PRDERD = '" & Format(Today(), "yyyy-MM-dd") & "', 
                    PRDJIRA = '" & Trim(jiraTask) & "', " & "VMVNUM = " & Trim(vendorNo) & ",PRDNEW = " & strChkSel & ",PRDEDD = '" & Format(Today(), "yyyy-MM-dd") & "',
                    PRDSCO = " & sampleCost & ",PRDTTC = " & miscCost & ",PRDUSR = '" & Trim(userSelec) & "',PRDDAT = '" & Format(Today(), "yyyy-MM-dd") & "',MOUSER = '" & userid & "',
                    MODATE = '" & Format(Today(), "yyyy-MM-dd") & "',PRDCTP = '" & Trim(tcpNo) & "',PRDSQTY = " & sampleQty & ", PRDQTY = " & qty & ",PRDMFR = '" & Trim(mfr) & "',
                    PRDMFR# = '" & Trim(mfrNo) & "',PRDCOS = " & unitCost & ",PRDCON = " & unitCostNew & ",PRDPO# = '" & Trim(poNo) & "',PODATE = '" & Format(Today(), "yyyy-MM-dd") & "',
                    PRDSTS = '" & Trim(status) & "',PRDBEN = '" & Trim(benefits) & "',PRDINF = '" & Trim(comments) & "' WHERE PRHCOD = " & Trim(code) & " AND
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDetail2(partstoshow As String, minorCode As String, tooCost As String, strDate1 As Date, vendorNo As String, strChkSel As String,
                                        strDate2 As Date, sampleCost As String, miscCost As String, userSelec As String, strDate3 As Date, userid As String, tcpNo As String, sampleQty As String,
                                        qty As String, mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, strDate4 As Date,
                                        benefits As String, comments As String, code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPTS = '" & partstoshow & "',PRDMPC = '" & minorCode & "',PRDTCO = " & tooCost & ",PRDERD = '" & Format(strDate1, "yyyy-MM-dd") & "', 
                     " & "VMVNUM = " & Trim(vendorNo) & ",PRDNEW = " & strChkSel & ",PRDEDD = '" & Format(strDate2, "yyyy-MM-dd") & "',
                    PRDSCO = " & sampleCost & ",PRDTTC = " & miscCost & ",PRDUSR = '" & Trim(userSelec) & "',PRDDAT = '" & Format(strDate3, "yyyy-MM-dd") & "',MOUSER = '" & userid & "',
                    MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDCTP = '" & Trim(tcpNo) & "',PRDSQTY = " & sampleQty & ", PRDQTY = " & qty & ",PRDMFR = '" & Trim(mfr) & "',
                    PRDMFR# = '" & Trim(mfrNo) & "',PRDCOS = " & unitCost & ",PRDCON = " & unitCostNew & ",PRDPO# = '" & Trim(poNo) & "',PODATE = '" & Format(strDate4, "yyyy-MM-dd") & "',
                    PRDBEN = '" & Trim(benefits) & "',PRDINF = '" & Trim(comments) & "' WHERE PRHCOD = " & Trim(code) & " AND
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProdClosedParts(userid As String, dtvalue As Date, strUser As String, strInfo As String, strName As String, strStatus As String, code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLH SET MOUSER = '" & userid & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDATE = '" & Format(dtvalue, "yyyy-MM-dd") & "',PRPECH = '" & strUser & "',
                    PRINFO = '" & strInfo & "',PRNAME = '" & strName & "',PRSTAT = '" & strStatus & "' WHERE PRHCOD = " & code
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProdOpenParts(userid As String, dtvalue As Date, strUser As String, strInfo As String, strName As String, code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLH SET MOUSER = '" & userid & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDATE = '" & Format(dtvalue, "yyyy-MM-dd") & "',PRPECH = '" & strUser & "',
                    PRINFO = '" & strInfo & "',PRNAME = '" & strName & "' WHERE PRHCOD = " & code
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProductDevHeader(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "update prdvlh set prstat = 'F' where prhcod = " & Trim(code)
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateMarktCampaignData(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE MACALE SET MACADY = 0 WHERE MACACO = " & code
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateProdDetailVendor(partstoshow As String, vendorno As String, code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET PRDPTS = '" & partstoshow & "',VMVNUM = " & vendorno & ", PRDCON = 0 WHERE PRHCOD = " & Trim(code) & " AND PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateChangedVendor(userId As String, vendorNo As String, partNo As String, codeNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET MOUSER = '" & userId & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',VMVNUM = " & vendorNo & ", PRDCON = 0 WHERE PRHCOD = " & codeNo & " AND 
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateChangedUC(userId As String, unitCost As String, partNo As String, codeNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET MOUSER = '" & userId & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "', PRDCON = '" & Trim(unitCost) & "' WHERE PRHCOD = " & codeNo & " AND 
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateChangedMFR(userId As String, mfrNo As String, partNo As String, codeNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET MOUSER = '" & userId & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "', PRDMFR# = '" & Trim(mfrNo) & "' WHERE PRHCOD = " & codeNo & " AND 
                    PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateChangedStatus(userId As String, status As String, partNo As String, codeNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE PRDVLD SET MOUSER = '" & userId & "',MODATE = '" & Format(Now, "yyyy-MM-dd") & "',PRDSTS = '" & Trim(status) & "' WHERE PRHCOD = " & codeNo & " AND PRDPTN = '" & Trim(UCase(partNo)) & "'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

    Public Function UpdateInvByPhotoAddition(partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim QueryResult As Integer = -1
        Try
            Sql = "UPDATE INMSTA SET imfpoe = '7' WHERE IMPTN = '" & Trim(UCase(partNo)) & "' AND IMFPOE <> '6'"
            QueryResult = UpdateDataInDatabase(Sql)
            Return QueryResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return QueryResult
        End Try
    End Function

#End Region

#Region "ComboBoxes"

#End Region

#Region "SelectionFields"

#End Region

#Region "Utils"

    Public Sub killBackgroundProcess()
        Dim exMessage As String = Nothing
        Dim processes As Process() = Process.GetProcesses()
        'Dim lstApp As List(Of String) = New List(Of String)()
        'Dim lstBackground As List(Of String) = New List(Of String)()
        Try

            For Each proc In processes
                If String.IsNullOrEmpty(proc.MainWindowTitle) Then
                    If proc.ProcessName = ProcessName Then
                        proc.Kill()
                        writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Trace, "Process Killed Succesfully.", "")
                        'writeComputerEventLog()
                    End If
                End If
            Next

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, "Exception", exMessage)
            'writeComputerEventLog()
        End Try
    End Sub

    Public Sub writeLog(strLogCadenaCabecera As String, strLevel As VBLog.ErrorTypeEnum, strMessage As String, strDetails As String)
        strLogCadena = strLogCadenaCabecera + " " + System.Reflection.MethodBase.GetCurrentMethod().ToString()

        vblog.WriteLog(strLevel, "CTPSystem" & strLevel, strLogCadena, userid, strMessage, strDetails)
    End Sub

    Public Shared Function CheckForInternetConnection() As Boolean
        Dim exMessage As String = Nothing
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            'writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return False
        End Try
    End Function

    Public Function adjustDatetimeFormat(documentName As String, documentExt As String) As String

        Dim exMessage As String = Nothing
        Try
            Dim name As String = Nothing
            Dim culture As CultureInfo = CultureInfo.CreateSpecificCulture("en-US")
            Dim dtfi As DateTimeFormatInfo = culture.DateTimeFormat
            dtfi.DateSeparator = "."

            Dim now As DateTime = DateTime.Now
            Dim halfName = now.ToString("G", dtfi)
            halfName = halfName.Replace(" ", ".")
            halfName = halfName.Replace(":", "")
            Dim fileName = documentName & "." & halfName & "." & documentExt
            Return fileName
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try

    End Function

    Public Function GetInactiveAlertByUser(userid As String, Optional vendorNo As String = Nothing) As DataSet
        Dim exMessage As String = " "
        Dim sql As String = ""
        Dim ds As DataSet = New DataSet()
        Try
            'sql = " select A1.prhcod As ProjectNo,A1.prdptn As PartNo, A1.vmvnum As VendorNo, A1.crdate As CreationDate, date(now()) As CurrentDate, A1.modate As ModificationDate, (Days(date(now())) - Days(A1.modate)) as DifferenceDays, A2.pqcomm As StatusComment, A1.prdusr As User 
            '        from prdvld A1 inner join poqota A2 on A1.prdptn = A2.pqptn and A1.vmvnum = A2.pqvnd where SUBSTR(UCASE(A2.SPACE),32,3) = 'DEV' AND A2.PQCOMM LIKE 'D%' 
            '        and (Days(date(now())) - Days(A1.modate)) > 30 and (Days(date(now())) - Days(A1.modate)) > 0 and PQCOMM LIKE 'D-Pending%' OR PQCOMM LIKE 'D-Analysis%' 
            '        and A1.prdusr = '" & Trim(UCase(userid)) & "' and A1.cruser = '" & Trim(UCase(userid)) & "'
            '        union
            '        select A0.prhcod  As ProjectNo, A0.prdptn As PartNo, A0.vmvnum As VendorNo, A0.crdate As CreationDate,date(now()) As CurrentDate, A0.modate As ModificationDate, (Days(date(now())) - Days(A0.modate)) as DifferenceDays, 
            '        CASE A0.prdsts
            '           WHEN 'AS' THEN 'D-Analysis of Sample'
            '           WHEN 'PS' THEN 'D-Pending from Suppl'   
            '         END  As StatusComment
            '        , A0.prdusr As User
            '        from prdvld A0 inner join vnmas A1 on A0.vmvnum = A1.vmvnum inner join csuser A2 on A1.vmabb# = A2.uspurc where A2.ususer = '" & Trim(UCase(userid)) & "'
            '        and prdsts in ('PS','AS') and (Days(date(now())) - Days(A0.modate)) > 30 and (Days(date(now())) - Days(A0.modate)) > 0"

            'and (PQCOMM LIKE 'D-Pending%' OR PQCOMM LIKE 'D-Analysis%'  PREVIOUS SENETENCE IN QUERY

            sql = " select A1.prhcod As ProjectNo,A1.prdptn As PartNo, A1.vmvnum As VendorNo, A1.crdate As CreationDate, date(now()) As CurrentDate, A1.modate As ModificationDate, (Days(date(now())) - Days(A1.modate)) as DifferenceDays, A2.pqcomm As StatusComment, A1.prdusr As User 
                    from prdvld A1 inner join poqota A2 on A1.prdptn = A2.pqptn and A1.vmvnum = A2.pqvnd where SUBSTR(UCASE(A2.SPACE),32,3) = 'DEV' AND A2.PQCOMM LIKE 'D%' 
                    and (Days(date(now())) - Days(A1.modate)) > 30 and (Days(date(now())) - Days(A1.modate)) > 0 and ucase(A2.pqcomm)  not like 'D-CLOSED%' 
                    and A1.prdusr = '" & Trim(UCase(userid)) & "' and A1.cruser = '" & Trim(UCase(userid)) & "'"

            If vendorNo IsNot Nothing Then
                Dim addToQuery As String = If(customIsVendorAccepted(vendorNo), " and A1.vmvnum = " & vendorNo, "")
                'Dim addToQuery As String = " and A1.vmvnum = " & vendorNo
                sql += addToQuery
            End If


            'Sql = "SELECT * FROM CSUSER WHERE USUSER = '" & Trim(UCase(userName)) & "'"
            ds = GetDataFromDatabase(sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetMassiveReferences(status As String, userid As String, flag As Boolean, Optional vendorNo As String = Nothing) As DataSet
        Dim exMessage As String = " "
        Dim sql As String = ""
        Dim ds As DataSet = New DataSet()
        Try

            If Not flag Then
                sql = " Select a2.prhcod  As ProjectNo, A1.prname As ProjectName, A2.crdate  As CreationDate,a2.prdptn  As PartNo, a2.prdctp As CTPNo, a2.prdmfr# As ManufacturerNo, a2.prdcon as UnitCost 
                    from prdvlh A1 inner join prdvld A2 On a1.prhcod = a2.prhcod where A2.prdusr = '" & Trim(UCase(userid)) & "' and A2.cruser = '" & Trim(UCase(userid)) & "' 
                    and a2.prdsts = '" & Trim(UCase(status)) & "'"
            Else
                sql = " Select a2.prhcod  As ProjectNo, A1.prname As ProjectName, A2.crdate  As CreationDate,a2.prdptn  As PartNo, a2.prdctp As CTPNo, a2.prdmfr# As ManufacturerNo, a2.prdcon as UnitCost 
                    from prdvlh A1 inner join prdvld A2 On a1.prhcod = a2.prhcod where a2.prdsts = '" & Trim(UCase(status)) & "'"
            End If

            If vendorNo IsNot Nothing Then
                Dim addToQuery As String = If(customIsVendorAccepted(vendorNo), " and A2.vmvnum = " & vendorNo & " order by 1 ", " order by 1 ")
                'Dim addToQuery As String = " and A1.vmvnum = " & vendorNo
                sql += addToQuery
            End If
            'Sql = "SELECT * FROM CSUSER WHERE USUSER = '" & Trim(UCase(userName)) & "'"
            ds = GetDataFromDatabase(sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getValueCheckTab3(vendorNo As String, partno As String)
        Dim exMessage As String = " "

        Try
            Dim listItemCat = VendorWhiteFlagMethod.Split(",")

            Dim dsResult1 = getItemCategoryByVendorAndPart(vendorNo, partno)
            If dsResult1 IsNot Nothing Then
                If dsResult1.Tables(0).Rows.Count > 0 Then
                    For Each item As String In listItemCat
                        If Trim(item).Equals(Trim(vendorNo)) Then
                            Return 1
                        End If
                    Next
                    Return -1
                Else
                    Return 1
                End If
            Else
                Return 1
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return -1
        End Try
    End Function

    Public Sub sendMessageOut(dgv As DataGridView, flag As Boolean)
        Dim exMessage As String = Nothing
        Try
            SendMessage(dgv.Handle, WM_SETREDRAW, flag, 0)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Public Sub OpenOutlookMessage(name As String, partNo As String, subject As String)
        Dim exMessage As String = ""
        Try

            Dim AppOutlook As New Outlook.Application
            Dim OutlookMessage As Object
            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)

            'OutlookMessage.Subject = "Newly Developed Part(s)"
            OutlookMessage.Subject = subject
            OutlookMessage.Body = "Project: " & Trim(name) & "  Part No. " & Trim(partNo)
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Display(True)

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Sub

    Public Function sendEmail1(toemails As String, ByVal userid As String) As Integer
        Dim exMessage As String = " "
        Dim AppOutlook As New Outlook.Application
        'Dim oNS As Outlook.NameSpace
        Dim OutlookMessage As Object
        Dim rsResult As Integer = 0
        Try
            'oNS = AppOutlook.GetNamespace("MAPI")
            'oNS.Logon("Outlokk", "", False, True)
            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
            'Dim Recipents1 As Outlook.Recipients = OutlookMessage.Recipients

            Dim listEmail As New List(Of String)
            Dim strArr() As String
            strArr = toemails.Split(";")
            For Each tt As String In strArr
                If Not String.IsNullOrEmpty(tt) Then
                    listEmail.Add(tt)
                End If
            Next

            For Each ttt As String In listEmail
                Recipents.Add(ttt)
                Recipents.ResolveAll()
            Next

            listEmail.Add("aavila@costex.com")

            'test purpose
            'Dim lenghtRec = Recipents.Count
            'For index As Integer = 1 To lenghtRec
            '    Recipents.Remove(index)
            'Next
            'Recipents1.Add("alexei.ansberto85@gmail.com")
            'Recipents1.Add("ansberto.avila85@gmail.com")
            'test purpose

            OutlookMessage.Subject = "Newly Developed Part(s)"

            OutlookMessage.Body = "User Notified. " & Trim(userid)
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Send() 'must be uncommented to send emails

            'oNS.Logoff()

            Return rsResult = 1
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            MessageBox.Show("Mail could Not be sent") 'if you dont want this message, simply delete this line 
            Return rsResult = -1
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try
    End Function

    Public Function sendEmail(toemails As String, Optional ByVal partNo As String = Nothing) As Integer
        Dim exMessage As String = " "
        Dim AppOutlook As New Outlook.Application
        'Dim oNS As Outlook.NameSpace
        Dim OutlookMessage As Object
        Dim rsResult As Integer = 0
        Try
            'oNS = AppOutlook.GetNamespace("MAPI")
            'oNS.Logon("Outlokk", "", False, True)
            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
            'Dim Recipents1 As Outlook.Recipients = OutlookMessage.Recipients

            If FlagTestEmails = 0 Then
                Dim listEmail As New List(Of String)
                Dim strArr() As String
                strArr = toemails.Split(";")
                For Each tt As String In strArr
                    If Not String.IsNullOrEmpty(tt) Then
                        listEmail.Add(tt)
                    End If
                Next

                listEmail.Add("aavila@costex.com")

                For Each ttt As String In listEmail
                    Recipents.Add(ttt)
                    Recipents.ResolveAll()
                Next
            Else
                'test purpose from config
                Dim listEmail As New List(Of String)
                Dim strArr() As String
                strArr = TestEmailAddresess.Split(";")
                For Each tt As String In strArr
                    If Not String.IsNullOrEmpty(tt) Then
                        listEmail.Add(tt)
                    End If
                Next

                For Each ttt As String In listEmail
                    Recipents.Add(ttt)
                    Recipents.ResolveAll()
                Next
            End If

            OutlookMessage.Subject = "New Report for User"

            OutlookMessage.Body = "Part No. " & Trim(partNo)
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Send() 'must be uncommented to send emails

            'oNS.Logoff()

            Return rsResult = 1
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            MessageBox.Show("Mail could Not be sent") 'if you dont want this message, simply delete this line 
            Return rsResult = -1
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try
    End Function

    Public Function checkfieldsPoQote(partNo As String, vendorNo As String, maxValue As String, strYear As String, strMonth As String, mpnPo As String, strDay As String,
                                    strStsQuote As String, strSpace As String, strUnitCostNew As String, strMinQty As String) As String
        Dim strError As String = String.Empty

#Region "NumericFields"
        If String.IsNullOrEmpty(vendorNo) Then
            strError += "Vendor Number,"
        End If
        If String.IsNullOrEmpty(maxValue) Then
            strError += "Sequencial,"
        End If
        If String.IsNullOrEmpty(strYear) Then
            strError += "Year,"
        End If
        If String.IsNullOrEmpty(strMonth) Then
            strError += "Month,"
        End If
        If String.IsNullOrEmpty(strDay) Then
            strError += "Day,"
        End If
        If String.IsNullOrEmpty(strUnitCostNew) Then
            strError += "Unit Cost New,"
        End If
        If String.IsNullOrEmpty(strMinQty) Then
            strError += "Min Qty,"
        End If
#End Region

        If String.IsNullOrEmpty(strError) Then
            Return ""
        Else
            Return strError
        End If

    End Function

    Public Function checkFields(projectno As String, partNo As String, dtValue As DateTimePicker, userid As String, dtValue1 As DateTimePicker, userid1 As String, dtValue2 As DateTimePicker, ctpNo As String, qty As String,
                                        mfr As String, mfrNo As String, unitCost As String, unitCostNew As String, poNo As String, dtValue3 As DateTimePicker, ddlStatus As String, benefits As String,
                                        comments As String, ddlUser As String, chkNew As CheckBox, dtValue4 As DateTimePicker, sampleCost As String, miscCost As String, vendorNo As String,
                                        partsToShow As String, ddlMinorCode As String, toolingCost As String, dtValue5 As DateTimePicker, strDate As String, sampleQty As String) As String
        Dim strError As String = String.Empty

#Region "TextBoxes"

#End Region

#Region "NumericFields"

        If String.IsNullOrEmpty(partNo) Then
            strError += "Part Number,"
        End If
        If String.IsNullOrEmpty(ctpNo) Then
            strError += "CTP Number,"
        End If
        If String.IsNullOrEmpty(ddlUser) Then
            strError += "Person in Charge,"
        End If
        If String.IsNullOrEmpty(unitCost) Then
            strError += "Unit Cost,"
        End If
        If String.IsNullOrEmpty(unitCostNew) Then
            strError += "Unit Cost New,"
        End If
        If String.IsNullOrEmpty(vendorNo) Then
            strError += "Vendor Number,"
        End If


        If String.IsNullOrEmpty(strError) Then
            Return ""
        Else
            Return strError
        End If

    End Function

    Public Function getmax(table, field)
        Dim exMessage As String = " "
        Dim Sql As String = " "
        Try
            Sql = "Select " & field & " FROM " & table & " ORDER BY " & field & " DESC FETCH FIRST 1 ROW ONLY"
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Using ObjCmd As Odbc.OdbcCommand = New Odbc.OdbcCommand(Sql, ObjConn)
                    ObjConn.Open()
                    ObjCmd.CommandType = CommandType.Text
                    Dim QueryResult = ObjCmd.ExecuteScalar()
                    Return QueryResult
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function getmaxComplex(table, field, strWhereAdd)
        Dim exMessage As String = " "
        Dim Sql As String = " "
        Try
            Sql = "Select " & field & " FROM " & table & " " & strWhereAdd & " ORDER BY " & field & " DESC FETCH FIRST 1 ROW ONLY"
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Using ObjCmd As Odbc.OdbcCommand = New Odbc.OdbcCommand(Sql, ObjConn)
                    ObjConn.Open()
                    ObjCmd.CommandType = CommandType.Text
                    Dim QueryResult = ObjCmd.ExecuteScalar()
                    Return QueryResult
                End Using
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function GetIpAddrTable()
        Dim exMessage As String = " "
        Try
            Dim Buf(0 To 511) As Byte
            Dim BufSize As Long : BufSize = UBound(Buf) + 1
            Dim rc As Long
            Dim ArrayOk As Array

            rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
            'If rc <> 0 Then Err.Raise VBObjectError, , "GetIpAddrTable failed With Return value " & rc
            If rc <> 0 Then Err.Raise(VBObjectError, , "GetIpAddrTable failed With Return value " & rc)
            Dim NrOfEntries As Integer : NrOfEntries = Buf(1) * 256 + Buf(0)
            If NrOfEntries = 0 Then GetIpAddrTable = ArrayOk : Exit Function
            'ReDim IpAddrs(0 To NrOfEntries - 1) As String
            Dim IpAddrs() As String
            ReDim IpAddrs(0 To NrOfEntries - 1)
            Dim i As Integer
            For i = 0 To NrOfEntries - 1
                Dim j As Integer, s As String : s = ""
                For j = 0 To 3 : s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j) : Next
                IpAddrs(i) = s
            Next
            GetIpAddrTable = IpAddrs
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Function

    Public Function LocalIPAddress(Optional ByVal bPreferred As Boolean = False) As String
        'Returns Local/Private IP address from all mapped/bind addresses
        'See the RFC 1918 for IP v4     -> address ranges for private networks
        'https://tools.ietf.org/html/rfc1918
        'and
        'RFC 4193 for IP v6             -> Local IPv6 Unicast Addresses / Unique Local Addresses (ULA)
        'https://tools.ietf.org/html/rfc4193
        Dim i As Long
        Dim IPAddrTable
        Dim C_ClassAddr As String
        Dim Buf(0 To 511) As Byte
        Dim BufSize As Long : BufSize = UBound(Buf) + 1
        Dim exMessage As String = Nothing
        Try
            IPAddrTable = GetIpAddrTable_API(Buf(0), BufSize, 1)

            For i = LBound(IPAddrTable) To UBound(IPAddrTable)
                If Len(IPAddrTable(i)) Then
                    Select Case Left$(IPAddrTable(i), 3)
                        Case "192" '192.168. range
                            C_ClassAddr = Mid$(IPAddrTable(i), 5, 3)
                            Select Case CInt(C_ClassAddr)
                                Case 168
                                    LocalIPAddress = IPAddrTable(i)
                                    Exit For
                            End Select
                        Case "172" '172.16. - 172.31. range
                            C_ClassAddr = Mid$(IPAddrTable(i), 5, 2)
                            Select Case CInt(C_ClassAddr)
                                Case 16 To 31
                                    LocalIPAddress = IPAddrTable(i)
                                    Exit For
                            End Select
                        Case "10." '10.0. - 10.255. range
                            If bPreferred = True Then 'default False, a class 10. addresses not counted as local IP.
                                C_ClassAddr = Mid$(IPAddrTable(i), 4, 3)
                                C_ClassAddr = Replace(C_ClassAddr, ".", "")
                                Select Case CInt(C_ClassAddr)
                                    Case 0 To 255
                                        LocalIPAddress = IPAddrTable(i)
                                        Exit For
                                End Select
                            End If
                    End Select
                End If
            Next i
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Function

    Public Shared Function GetARPTablr() As String
        ' The number of bytes needed.
        Dim bytesNeeded As Integer = 0
        ' The result from the API call.
        Dim result As Integer = GetIpNetTable(IntPtr.Zero, bytesNeeded, False)
        ' Call the function, expecting an insufficient buffer.
        If result <> ERROR_INSUFFICIENT_BUFFER Then
            ' Throw an exception.
            Throw New Win32Exception(result)
        End If
        ' Allocate the memory, do it in a try/finally block, to ensure
        ' that it is released.
        Dim buffer As IntPtr = IntPtr.Zero

        ' Try/finally.
        Try
            ' Allocate the memory.
            buffer = Marshal.AllocCoTaskMem(bytesNeeded)
            ' Make the call again. If it did not succeed, then
            ' raise an error.
            result = GetIpNetTable(buffer, bytesNeeded, False)
            ' If the result is not 0 (no error), then throw an exception.
            If result <> ERROR_SUCCESS Then
                ' Throw an exception.
                Throw New Win32Exception(result)
            End If
            ' Now we have the buffer, we have to marshal it. We can read
            ' the first 4 bytes to get the length of the buffer.
            Dim entries As Integer = Marshal.ReadInt32(buffer)
            ' Increment the memory pointer by the size of the int.
            Dim currentBuffer As New IntPtr(buffer.ToInt64() + Marshal.SizeOf(GetType(Integer)))

            ' Allocate an array of entries.
            Dim table As MIB_IPNETROW() = New MIB_IPNETROW(entries - 1) {}
            ' Cycle through the entries.
            For index As Integer = 0 To entries - 1
                ' Call PtrToStructure, getting the structure information.
                table(index) = DirectCast(Marshal.PtrToStructure(New IntPtr(currentBuffer.ToInt64() + (index * Marshal.SizeOf(GetType(MIB_IPNETROW)))), GetType(MIB_IPNETROW)), MIB_IPNETROW)
            Next
            For index As Integer = 0 To entries - 1
                If table(index).dwType <> DWTYPES.Invalid And table(index).dwType <> DWTYPES.Other Then
                    Dim ip As New IPAddress(table(index).dwAddr)
                    Dim mac As New PhysicalAddress(table(index).bPhysAddr)

                    Dim pepe = table(index).dwType.ToString & vbTab & vbTab & "IP:" + ip.ToString() & vbTab & vbTab & "MAC: " & MACtoString(mac)
                    Return pepe

                    'Console.WriteLine(table(index).dwType.ToString & vbTab & vbTab & "IP:" + ip.ToString() & vbTab & vbTab & "MAC: " & MACtoString(mac))
                End If
            Next
        Finally
            ' Release the memory.
            Marshal.FreeCoTaskMem(buffer)
            '  Marshal.FreeHGlobal(rowptr)
        End Try
    End Function

    Public Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim exMessage As String = " "
        Try
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    GetIPv4Address = ipheal.ToString()
                    Return GetIPv4Address
                End If
            Next
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Function

    Public Shared Function MACtoString(mac As PhysicalAddress, Optional Capital As Boolean = True) As String
        If Capital Then ' In capital Letters
            Return String.Join(":", (From z As Byte In mac.GetAddressBytes Select z.ToString("X2")).ToArray())
        Else
            Return String.Join(":", (From z As Byte In mac.GetAddressBytes Select z.ToString("x2")).ToArray())
        End If
    End Function

    Public Function checkstring(StrInput)
        Dim exMessage As String = Nothing
        Try
            If InStr(1, Trim(StrInput), "'") Or InStr(1, Trim(StrInput), "|") Or InStr(1, Trim(StrInput), "`") Or InStr(1, Trim(StrInput), "~") Or InStr(1, Trim(StrInput), "!") Or InStr(1, Trim(StrInput), "^") Or InStr(1, Trim(StrInput), "_") Or InStr(1, Trim(StrInput), "=") Or InStr(1, Trim(StrInput), "\") Or InStr(1, Trim(StrInput), "%") Or InStr(1, Trim(StrInput), "+") Or InStr(1, Trim(StrInput), "[") Or InStr(1, Trim(StrInput), "]") Or InStr(1, Trim(StrInput), "?") Or InStr(1, Trim(StrInput), "<") Or InStr(1, Trim(StrInput), ">") Then
                checkstring = False
            Else
                checkstring = True
            End If

            Exit Function
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
        'Call gotoerror("general", "checkstring", Err.Number, Err.Description, Err.Source)
    End Function

    Function checkusr(userid, pass)
        'call routine for encrypting password
        Dim as400 As New cwbx.AS400System
        Dim prog As New cwbx.Program
        Dim parms As New cwbx.ProgramParameters
        Dim server As New cwbx.SystemNames
        Dim stringCvtr As New cwbx.StringConverter
        Dim cwbcoPromptNever As New cwbx.cwbcoPromptModeEnum

        Dim wuser, wpass, wswvld
        Dim exMessage As String = " "
        Try

            'Program Parameters
            wuser = Left((Trim(UCase(userid)) & "          "), 10)
            wpass = Left((Trim(UCase(pass)) & "          "), 10)
            wswvld = "0"

            'AS400 Connection Parameters
            as400.Define(server.DefaultSystem)
            as400.UserID = "INTRANET"
            as400.Password = "CTP6100"
            'as400.IPAddress = "SVR400"
            as400.PromptMode = cwbcoPromptNever
            as400.Signon()

            as400.Connect(cwbx.cwbcoServiceEnum.cwbcoServiceODBC)

            If as400.IsConnected(cwbx.cwbcoServiceEnum.cwbcoServiceODBC) = 1 Then

                'Program to call
                prog.system = as400
                prog.LibraryName = "CTPINV"
                prog.ProgramName = "PSWVLDR"

                'Assign Values to Parameters
                parms.Clear()

                parms.Append("USER", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 10)
                parms.Append("PASS", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 20)
                parms.Append("SWVLD", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 1)
                'parms.Append("out", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 10)

                stringCvtr.CodePage = 37

                parms("USER").Value = stringCvtr.ToBytes(wuser)
                parms("PASS").Value = stringCvtr.ToBytes(wpass)
                parms("SWVLD").Value = stringCvtr.ToBytes(wswvld)

                prog.Call(parms)

                checkusr = stringCvtr.FromBytes(parms("SWVLD").Value)

                as400.Disconnect(cwbx.cwbcoServiceEnum.cwbcoServiceAll)

            End If
            Exit Function
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try

    End Function

    Public Sub changeVendor(PartNumber, VendorNumber, User_ID)
        Dim as400 As New cwbx.AS400System
        Dim prog As New cwbx.Program
        Dim parms As New cwbx.ProgramParameters
        Dim server As New cwbx.SystemNames
        Dim pckCvtr As New cwbx.PackedConverter
        Dim stringCvtr As New cwbx.StringConverter
        Dim p1, p2, p3, p4, p5, returnedValue
        Dim cwbcoPromptNever As New cwbx.cwbcoPromptModeEnum
        Dim exMessage As String = Nothing
        Try
            'Program Parameters
            p1 = "01"
            p2 = Left((Trim(UCase(PartNumber)) & "                   "), 19)
            p3 = Right("000000" & Trim(VendorNumber), 6)
            p4 = Left((User_ID & "          "), 10)
            p5 = "CTP "

            'AS400 Connection Parameters
            as400.Define(server.DefaultSystem)
            as400.UserID = "INTRANET"
            as400.Password = "CTP6100"
            'as400.IPAddress = "172.0.0.21"
            as400.PromptMode = cwbcoPromptNever
            as400.Signon()

            as400.Connect(cwbx.cwbcoServiceEnum.cwbcoServiceODBC)

            If as400.IsConnected(cwbx.cwbcoServiceEnum.cwbcoServiceODBC) = 1 Then
                'Program to call
                prog.system = as400
                prog.LibraryName = "CTPINV"
                prog.ProgramName = "UPDDVVNDR"

                parms.Clear()

                'Parameters Definition
                parms.Append("LOC", cwbx.cwbrcParameterTypeEnum.cwbrcInput, 2)
                parms.Append("PART#", cwbx.cwbrcParameterTypeEnum.cwbrcInput, 19)
                parms.Append("VENDOR", cwbx.cwbrcParameterTypeEnum.cwbrcInput, 6)
                parms.Append("USER", cwbx.cwbrcParameterTypeEnum.cwbrcInput, 10)
                parms.Append("WS", cwbx.cwbrcParameterTypeEnum.cwbrcInput, 10)
                parms.Append("ERROR", cwbx.cwbrcParameterTypeEnum.cwbrcInout, 3)

                stringCvtr.CodePage = 37

                'Assign Values to Parameters
                parms("LOC").Value = stringCvtr.ToBytes(p1)
                parms("PART#").Value = stringCvtr.ToBytes(p2)
                parms("VENDOR").Value = stringCvtr.ToBytes(p3)
                parms("USER").Value = stringCvtr.ToBytes(p4)
                parms("WS").Value = stringCvtr.ToBytes(p5)
                parms("ERROR").Value = stringCvtr.ToBytes("   ")

                prog.Call(parms)

                returnedValue = stringCvtr.FromBytes(parms("ERROR").Value)

                If Trim(returnedValue) = "AUT" Then
                    MsgBox("Vendor could not be changed, user not authorized!", vbInformation + vbOKOnly, "CTP System")
                End If
                If Trim(returnedValue) = "VEN" Then
                    MsgBox("Vendor could not be changed, Invalid Vendor!", vbInformation + vbOKOnly, "CTP System")
                End If
                If Trim(returnedValue) = "L/P" Then
                    MsgBox("Vendor could not be changed, Invalid Location or Part Number!", vbInformation + vbOKOnly, "CTP System")
                End If

                as400.Disconnect(cwbx.cwbcoServiceEnum.cwbcoServiceAll)
            End If
            Exit Sub
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Function FillGrid(query As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim PageSize As Integer = 5
        Try
            Dim ObjConn As New Odbc.OdbcConnection(strconnection)
            Dim dataAdapter As New Odbc.OdbcDataAdapter()
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            ObjConn.Open()
            'Sql = "SELECT COUNT(*) TFIELDS FROM PRDVLH " & strwhere
            Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
            dataAdapter = New Odbc.OdbcDataAdapter(cmd)
            dataAdapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            Else
                'message box warning
                Return Nothing
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Sub startProcessOF(pathToOpen As String)
        Dim exMessage As String = " "
        Try
            Dim ProcessProperties As New ProcessStartInfo
            ProcessProperties.FileName = pathToOpen
            ProcessProperties.WindowStyle = ProcessWindowStyle.Maximized
            Dim myProcess As Process = Process.Start(ProcessProperties)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Sub

    Public Function getProgramByExtension(fileName As String) As String
        Dim exMessage As String = " "
        Dim fileNameOk As String = Nothing
        Try
            Select Case fileNameOk
                Case "pdf"
                    Console.WriteLine("Excellent!")
                Case "jpg", "jpeg", "png"
                    Console.WriteLine("Well done")
                Case "txt"
                    fileNameOk = "notepad"
                Case "doc", "docx"
                    fileNameOk = "winword"
                Case Else
                    fileNameOk = ""
            End Select
            Return fileNameOk
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return fileNameOk
        End Try
    End Function

#End Region
#End Region

#Region "Generic Methods"

    'create single class for as400 connection

    Private Function GetDataFromDatabase(query As String) As Data.DataSet
        Dim exMessage As String = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()

                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)

                ObjConn.Close()

                If ds.Tables(0).Rows.Count > 0 Then
                    Return ds
                Else
                    'message box warning
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function GetSingleDataFromDatabase(query As String, columnToChange As String) As String
        Dim exMessage As String = " "
        Dim DescriptionCode As String = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()

                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)

                ObjConn.Close()

                Dim index = ds.Tables(0).Columns(columnToChange).Ordinal
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each RowDs In ds.Tables(0).Rows
                        DescriptionCode = ds.Tables(0).Rows(0).Item(index).ToString()
                        Exit For
                    Next
                    Return DescriptionCode
                Else
                    'message box warning
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function InsertDataInDatabase(query As String) As Integer
        Dim exMessage As String = " "
        Dim result As Integer = -1
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()
                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                result = dataAdapter.Fill(ds)
                Return result
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return result
        End Try
    End Function

    Private Function UpdateDataInDatabase(query As String) As String
        Dim exMessage As String = " "
        'Dim result As Integer = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture

                ObjConn.Open()

                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                dataAdapter = New Odbc.OdbcDataAdapter(cmd)
                dataAdapter.Fill(ds)

                ObjConn.Close()
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function UpdateDataInDatabase1(query As String) As String
        Dim exMessage As String = " "
        'Dim result As Integer = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture
                Dim rows As Integer

                ObjConn.Open()

                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                rows = cmd.ExecuteNonQuery()

                ObjConn.Close()

                Return rows
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    Private Function DeleteRecordFromDatabase(query As String) As Integer
        Dim exMessage As String = " "
        Try
            Using ObjConn As Odbc.OdbcConnection = New Odbc.OdbcConnection(strconnection)
                Dim dataAdapter As New Odbc.OdbcDataAdapter()
                Dim ds As New DataSet()
                ds.Locale = CultureInfo.InvariantCulture
                Dim rows As Integer

                ObjConn.Open()

                Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
                rows = cmd.ExecuteNonQuery()

                ObjConn.Close()

                Return rows
            End Using
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function


#End Region

#Region "Delete"

    Public Function DeleteDataByWSCod(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1

        Try
            Sql = "DELETE FROM PRDWL WHERE WHLCODE = " & Trim(code)
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromProdHead(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1
        Try
            Sql = "delete from prdvlh where prhcod = " & Trim(code)
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromProdDet(code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1
        Try
            Sql = "delete from prdvld where prhcod = " & Trim(code) & " and prdptn = '" & Trim(partNo) & "'"
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromProdDet1(code As String, partNo As String, vendorNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1
        Try
            Sql = "delete from prdvld where prhcod = " & Trim(code) & " and prdptn = '" & Trim(partNo) & "' and vmvnum = " & Trim(vendorNo)
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromPoQota(vendorNo As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1

        Try
            Sql = "delete FROM POQOTA WHERE PQVND = " & Trim(vendorNo) & " AND PQPTN = '" & Trim(UCase(partNo)) & "' AND SUBSTR(UCASE(SPACE),32,3) = 'DEV' AND PQCOMM LIKE 'D%'"
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromProdCommHeader(code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1
        Try
            Sql = "delete from prdcmh where prhcod = " & Trim(code) & " and prdptn = '" & Trim(partNo) & "'"
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteDataFromProdCommDet(code As String, partNo As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1
        Try
            Sql = "delete from prdcmd where prhcod = " & Trim(code) & " and prdptn = '" & Trim(partNo) & "'"
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteRecorFromLoginTcp(code As String) As Integer
        Dim exMessage As String = " "
        Dim Sql As String
        Dim rsConfirm As Integer = -1

        Try
            Sql = "delete from loginctp where codlogin = " & code
            rsConfirm = DeleteRecordFromDatabase(Sql)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

    Public Function DeleteGeneral(table As String, field As String, values As String) As Integer
        Dim exMessage As String = " "
        'Dim Sql As String
        Dim rsConfirm As Integer = -1

        Try
            Dim sqlQuery As String = "DELETE FROM {0} WHERE {1} IN ({2})"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, table, field, values)
            rsConfirm = DeleteRecordFromDatabase(sqlFormattedQuery)
            'rsConfirm = ExecuteNotQueryCommand(sqlFormattedQuery, strconnection)
            Return rsConfirm
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return rsConfirm
        End Try
    End Function

#End Region

#Region "SQL Server Methods"

    'create a single class for sql server connection


    Public Function getmaxSQL(table, field) As Object

        '        Dim intrespond As Long
        '        Dim sentence As Variant



        '        Set RsGeneral = Nothing
        '        Set CMD = Nothing
        '        If ConnSql.State = 1 Then
        '        Else
        '            ConnSql.ConnectionString = strconnSQL
        '            ConnSql.Open()
        '        End If
        '        CMD.ActiveConnection = ConnSql
        '        CMD.CommandText = "spgetmax"
        '        CMD.CommandType = adCmdStoredProc
        '        sentence = "Select Max(" & field & ") As max from " & table
        '        Set RsGeneral = CMD.Execute(, Array(sentence))

        '        If IsNull(RsGeneral.Fields(0)) Then
        '            getmaxSQL = 1
        '        Else
        '            getmaxSQL = RsGeneral.Fields(0) + 1
        '        End If

        '        Exit Function
        'errhandler:
        '        Call gotoerror("general", "getmaxSQL", Err.Number, Err.Description, Err.Source)
    End Function

    Public Function GetMaxCodeDetSql(table As String, field As String) As Data.DataSet
        Try
            Dim sqlQuery As String = "select MAX({1}) from {0}"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, table, field)
            Dim dsResult = ExecuteQueryCommand(sqlFormattedQuery, strconnSQL)
            Return dsResult
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'get data from sql
    Public Function GetDataSqlByUser(table As String, userid As String) As Data.DataSet
        Try
            Dim sqlQuery As String = "SELECT * FROM {0} WHERE USERID = '{1}'"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, table, userid)
            Dim dsResult = ExecuteQueryCommand(sqlFormattedQuery, strconnSQL)
            Return dsResult
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'delete data from sql
    Public Function DeleteDataSqlByUser(table As String, userid As String) As Integer
        Try
            Dim sqlQuery As String = "DELETE FROM {0} WHERE USERID = '{1}'"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, table, userid)
            Dim rsResult = ExecuteNotQueryCommand(sqlFormattedQuery, strconnSQL)
            Return rsResult
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'insert data from sql
    Public Function InsertDataSqlByUser(table As String, userid As String, listData As List(Of String)) As Integer
        Dim codComment As Integer
        Dim comment As String

        Try

            codComment = listData(0)
            comment = listData(1)

            Dim sqlQuery As String = "INSERT INTO {0} (cod_comment, userid, comment) VALUES ({1}, '{2}', '{3}')"
            Dim sqlFormattedQuery As String = String.Format(sqlQuery, table, codComment, userid, comment)
            Dim rsResult = ExecuteNotQueryCommand(sqlFormattedQuery, strconnSQL)
            Return rsResult

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    'query sin devolver resultados
    Public Function ExecuteNotQueryCommand(queryString As String, connectionString As String) As Integer

        Dim exMessage As String = " "
        Dim rsResult As Integer = -1
        Try
            Using connection As New SqlConnection(connectionString)
                Dim command As New SqlCommand(queryString, connection)
                command.CommandType = CommandType.Text

                command.Connection.Open()
                rsResult = command.ExecuteNonQuery()
            End Using
            Return rsResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return rsResult
        Finally

        End Try

    End Function

    'query devolviendo resultados
    Private Function ExecuteQueryCommand(queryString As String, connectionString As String) As Data.DataSet

        Dim exMessage As String = " "
        Dim rsResult As Integer = -1
        Dim dsResult As DataSet = New DataSet()
        Try
            Using connection As New SqlConnection(connectionString)
                Dim command As New SqlCommand(queryString, connection)
                command.Connection.Open()
                Dim tblResult As New DataTable
                tblResult.Load(command.ExecuteReader())
                dsResult.Tables.Add(tblResult)
                Return dsResult
            End Using

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return dsResult
        Finally

        End Try

    End Function

    Public Function FillGridSql(query As String) As Data.DataSet
        Dim exMessage As String = " "
        'Dim rsResult As Integer
        Dim dsResult As DataSet
        Try
            dsResult = ExecuteQueryCommand(query, strconnSQL)

            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            'ObjConn.Open()
            ''Sql = "SELECT COUNT(*) TFIELDS FROM PRDVLH " & strwhere
            'Dim cmd As New Odbc.OdbcCommand(query, ObjConn)
            'dataAdapter = New Odbc.OdbcDataAdapter(cmd)
            'dataAdapter.Fill(ds)

            'If ds.Tables(0).Rows.Count > 0 Then
            '    Return ds
            'Else
            '    'message box warning
            '    Return Nothing
            'End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

#End Region

    Public Function getCell2(code As String) As DataSet
        Dim exMessage As String = " "
        Try
            Dim ds As New DataSet()
            ds.Locale = CultureInfo.InvariantCulture

            Dim Sql = "SELECT PRDDAT,PRDPTN,PRDCTP,PRDMFR#,PRDVLD.VMVNUM,VMNAME,PRDSTS,PRDJIRA,PRDUSR FROM PRDVLD INNER JOIN VNMAS ON PRDVLD.VMVNUM = VNMAS.VMVNUM WHERE PRHCOD = " & code & " "  'DELETE BURNED REFERENCE
            'get the query results
            ds = FillGrid(Sql)
            Return ds
        Catch ex As Exception
            Return Nothing
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
        End Try
    End Function

    Public Function GetTestData(partNo As String) As Data.DataSet
        Dim exMessage As String = " "
        Dim Sql As String
        Dim ds As New DataSet()
        ds.Locale = CultureInfo.InvariantCulture
        Try
            Sql = "select * from inmsta where trim(ucase(imptn)) = '" & Trim(UCase(partNo)) & "' fetch first 10 row only"
            ds = GetDataFromDatabase(Sql)
            Return ds
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            writeLog(strLogCadenaCabecera, VBLog.ErrorTypeEnum.Exception, ex.Message, ex.ToString())
            Return Nothing
        End Try
    End Function

    'Public Sub Generate_Log(Message As String)
    'On Error GoTo Generate_Log_Err
    'Dim LogFile As String
    '   LogFile = ""
    '   LogFile = DirLog + "ErrLog_" + Trim(Str(DatePart("m", Date))) + ".log"
    '  Open LogFile For Append As #2
    'Write #2, Format(Of Date, "mm/dd/yyyy")() + " " + Trim(Str(Time())) + "|" + Message
    'Close #2
    'Exit Sub
    'Generate_Log_Err:
    'Close #2
    'End Sub

    'Public Sub Generate_Log(Message As String)
    'On Error GoTo Generate_Log_Err
    'Dim LogFile As String
    '   LogFile = ""
    '  LogFile = DirLog + "ErrLog_" + Trim(Str(DatePart("m", Date))) + ".log"
    ' Open LogFile For Append As #2
    'Write #2, Format(Of Date, "mm/dd/yyyy")() + " " + Trim(Str(Time())) + "|" + Message
    'Close #2
    'Exit Sub
    'Generate_Log_Err:
    '       Close #2
    'End Sub

    ' Returns an array with the local IP addresses (as strings).
    ' Author: Christian d'Heureuse, www.source-code.biz

End Class
