Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations



<MetadataTypeAttribute(GetType(ProductMetadata))>
Public Class Product

    Sub MySub()

    End Sub

    Private _validationResults As New List(Of ValidationResult)
    Public ReadOnly Property ValidationResults() As List(Of ValidationResult)
        Get
            Return _validationResults
        End Get
    End Property

    Public Function IsValid() As Boolean

        TypeDescriptor.AddProviderTransparent(New AssociatedMetadataTypeTypeDescriptionProvider(GetType(Product), GetType(ProductMetadata)), GetType(Product))

        Dim result As Boolean = True
        Dim context = New ValidationContext(Me, Nothing, Nothing)

        Dim validation = Validator.TryValidateObject(Me, context, _validationResults, True)

        If Not validation Then
            result = False
        End If

        Return result

    End Function

    Public Function IsValid(obj As Object) As Boolean

        TypeDescriptor.AddProviderTransparent(New AssociatedMetadataTypeTypeDescriptionProvider(GetType(ProductHeader), GetType(ProductMetadata)), GetType(ProductHeader))

        Dim result As Boolean = True
        Dim errors As List(Of ValidationResult) = New List(Of ValidationResult)()

        'Dim destinationType = LikeSession.objToCast
        'Dim value As Object = Convert.ChangeType(obj,)

        Dim context = New ValidationContext(obj, Nothing, Nothing)

        Dim validation = Validator.TryValidateObject(obj, context, _validationResults, True)

        If Not validation Then
            result = False
        End If

        Return result

    End Function

End Class

Friend NotInheritable Class ProductMetadata

    Sub MySub()

    End Sub

#Region "Class attributes declaration"

    <Required(ErrorMessage:="Project Name is Required", AllowEmptyStrings:=False)>
    <StringLength(6, ErrorMessage:="Too Long")>
    Private prhcod As String
    Public Property projectNo() As String
        Get
            Return prhcod
        End Get
        Set(ByVal value As String)
            prhcod = value
        End Set
    End Property

    <Required(ErrorMessage:="Project Date is Required", AllowEmptyStrings:=False)>
    <DataType(DataType.Date)>
    Private prdate As String
    Public Property projectDate() As String
        Get
            Return prdate
        End Get
        Set(ByVal value As String)
            prdate = value
        End Set
    End Property

    <StringLength(250, ErrorMessage:="Too Long")>
    Private prinfo As String
    Public Property projectInfo() As String
        Get
            Return prinfo
        End Get
        Set(ByVal value As String)
            prinfo = value
        End Set
    End Property

    <StringLength(100, ErrorMessage:="Too Long")>
    Private prname As String
    Public Property projectName() As String
        Get
            Return prname
        End Get
        Set(ByVal value As String)
            prname = value
        End Set
    End Property

    <StringLength(1, ErrorMessage:="Too Long")>
    Private prstat As String
    Public Property projectStat() As String
        Get
            Return prstat
        End Get
        Set(ByVal value As String)
            prstat = value
        End Set
    End Property

    <Required(ErrorMessage:="Creation Date is Required", AllowEmptyStrings:=False)>
    <DataType(DataType.Date)>
    Private crdate As String
    Public Property creationDate() As String
        Get
            Return crdate
        End Get
        Set(ByVal value As String)
            crdate = value
        End Set
    End Property

    <StringLength(10, ErrorMessage:="Too Long")>
    Private cruser As String
    Public Property creationUser() As String
        Get
            Return cruser
        End Get
        Set(ByVal value As String)
            cruser = value
        End Set
    End Property

    <Required(ErrorMessage:="Modification Date is Required", AllowEmptyStrings:=False)>
    <DataType(DataType.Date)>
    Private modate As String
    Public Property modificationDate() As String
        Get
            Return modate
        End Get
        Set(ByVal value As String)
            modate = value
        End Set
    End Property

    <StringLength(10, ErrorMessage:="Too Long")>
    Private mouser As String
    Public Property modificationUser() As String
        Get
            Return mouser
        End Get
        Set(ByVal value As String)
            mouser = value
        End Set
    End Property

    <StringLength(10, ErrorMessage:="Too Long")>
    Private prpech As String
    Public Property personInCharge() As String
        Get
            Return prpech
        End Get
        Set(ByVal value As String)
            prpech = value
        End Set
    End Property

#End Region

End Class

Friend NotInheritable Class ProductDetailsMetadata

    Sub MySub()

    End Sub

#Region "Class attributes declaration"

    <Required(ErrorMessage:="Project Number is Required", AllowEmptyStrings:=False)>
    <StringLength(6, ErrorMessage:="Too Long")>
    Private prhcod As String
    Public Property ProjectNo() As String
        Get
            Return prhcod
        End Get
        Set(ByVal value As String)
            prhcod = value
        End Set
    End Property

    <Required(ErrorMessage:="Part number is Required", AllowEmptyStrings:=False)>
    <StringLength(19, ErrorMessage:="Too Long")>
    Private prdptn As String
    Public Property PartNo() As String
        Get
            Return prdptn
        End Get
        Set(ByVal value As String)
            prdptn = value
        End Set
    End Property
    '5 properties

    <Required(ErrorMessage:="CTP Number is Required", AllowEmptyStrings:=False)>
    <StringLength(50, ErrorMessage:="Too Long")>
    Private prdctp As String
    Public Property CTPNo() As String
        Get
            Return prdctp
        End Get
        Set(ByVal value As String)
            prdctp = value
        End Set
    End Property

    <Required(ErrorMessage:="Quantity is Required", AllowEmptyStrings:=False)>
    <StringLength(6, ErrorMessage:="Too Long")>
    Private prdqty As String
    Public Property Qty() As String
        Get
            Return prdqty
        End Get
        Set(ByVal value As String)
            prdqty = value
        End Set
    End Property

    <Required(ErrorMessage:="Manufacturer Number is Required", AllowEmptyStrings:=False)>
    <StringLength(50, ErrorMessage:="Too Long")>
    Private prdmfr As String
    Public Property ManufactNo() As String
        Get
            Return prdmfr
        End Get
        Set(ByVal value As String)
            prdmfr = value
        End Set
    End Property

    <Required(ErrorMessage:="Unit Cost is Required", AllowEmptyStrings:=False)>
    <StringLength(11, ErrorMessage:="Too Long")>
    Private prdcos As String
    Public Property UnitCost() As String
        Get
            Return prdcos
        End Get
        Set(ByVal value As String)
            prdcos = value
        End Set
    End Property

    <Required(ErrorMessage:="Unit Cost New is Required", AllowEmptyStrings:=False)>
    <StringLength(11, ErrorMessage:="Too Long")>
    Private prdcon As String
    Public Property UnitCostNew() As String
        Get
            Return prdcon
        End Get
        Set(ByVal value As String)
            prdcon = value
        End Set
    End Property

    <Required(ErrorMessage:="Status is Required", AllowEmptyStrings:=False)>
    <StringLength(2, ErrorMessage:="Too Long")>
    Private prdsts As String
    Public Property Status() As String
        Get
            Return prdsts
        End Get
        Set(ByVal value As String)
            prdsts = value
        End Set
    End Property

    <Required(ErrorMessage:="New Vendor Supplier Required", AllowEmptyStrings:=False)>
    <StringLength(1, ErrorMessage:="Too Long")>
    Private prdnew As String
    Public Property NewOrSupplier() As String
        Get
            Return prdnew
        End Get
        Set(ByVal value As String)
            prdnew = value
        End Set
    End Property

    <Required(ErrorMessage:="Vendor Number is Required", AllowEmptyStrings:=False)>
    <StringLength(6, ErrorMessage:="Too Long")>
    Private vmvnum As String
    Public Property vendorNo() As String
        Get
            Return vmvnum
        End Get
        Set(ByVal value As String)
            vmvnum = value
        End Set
    End Property

    <Required(ErrorMessage:="Minor Code is Required", AllowEmptyStrings:=False)>
    <StringLength(2, ErrorMessage:="Too Long")>
    Private prdmpc As String
    Public Property MinorCode() As String
        Get
            Return prdmpc
        End Get
        Set(ByVal value As String)
            prdmpc = value
        End Set
    End Property

    <Required(ErrorMessage:="Minimun Quantity is Required", AllowEmptyStrings:=False)>
    <StringLength(7, ErrorMessage:="Too Long")>
    Private pqmin As String
    Public Property MinQty() As String
        Get
            Return prdmpc
        End Get
        Set(ByVal value As String)
            prdmpc = value
        End Set
    End Property

#End Region

End Class

Public Class ProductClass

    Sub MySub()

    End Sub

    Dim productHeader As ProductHeader
    Public Property Header() As ProductHeader
        Get
            Return productHeader
        End Get
        Set(ByVal value As ProductHeader)
            productHeader = value
        End Set
    End Property

    Private prodHeaderCollection As ProductHeaderCollection
    Public Property ProdHeaders() As ProductHeaderCollection
        Get
            Return prodHeaderCollection
        End Get
        Set(ByVal value As ProductHeaderCollection)
            prodHeaderCollection = value
        End Set
    End Property

End Class





