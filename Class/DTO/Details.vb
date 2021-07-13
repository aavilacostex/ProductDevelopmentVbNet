Public Class Details

    Sub MySub()

        _objDetail = New Details
        _objDetail._details = New ProductDetails()

    End Sub

    Private _objDetail As Details
    Public Property ObjDetail() As Details
        Get
            Return _objDetail
        End Get
        Set(ByVal value As Details)
            _objDetail = value
        End Set
    End Property

    Private _details As ProductDetails
    Public Property Details() As ProductDetails
        Get
            Return _details
        End Get
        Set(ByVal value As ProductDetails)
            _details = value
        End Set
    End Property

    Private _lstProdDetails As List(Of ProductDetails)
    Public Property LstProdDetails() As List(Of ProductDetails)
        Get
            Return _lstProdDetails
        End Get
        Set(ByVal value As List(Of ProductDetails))
            _lstProdDetails = value
        End Set
    End Property

End Class
