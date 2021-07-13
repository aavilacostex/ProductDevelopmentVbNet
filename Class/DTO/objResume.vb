Public Class objResume

    Sub MySub()

        ResumeObject = New objResume()

    End Sub

    Public Enum TypeObj
        notDefined = 0
        addition
        deletion
    End Enum

#Region "Class attributes declaration"

    Private partNo As String
    Public Property PartNumber() As String
        Get
            Return partNo
        End Get
        Set(ByVal value As String)
            partNo = value
        End Set
    End Property

    Private vendorNo As String
    Public Property VendorNumber() As String
        Get
            Return vendorNo
        End Get
        Set(ByVal value As String)
            vendorNo = value
        End Set
    End Property

    Private objDesc As String
    Public Property Description() As String
        Get
            Return objDesc
        End Get
        Set(ByVal value As String)
            objDesc = value
        End Set
    End Property

    Private objType As TypeObj = TypeObj.notDefined
    Public Property TypeResume() As String
        Get
            Return objType
        End Get
        Set(ByVal value As String)
            objType = value
        End Set
    End Property

    Private _objResume As objResume
    Public Property ResumeObject() As objResume
        Get
            Return _objResume
        End Get
        Set(ByVal value As objResume)
            _objResume = value
        End Set
    End Property

#End Region

End Class
