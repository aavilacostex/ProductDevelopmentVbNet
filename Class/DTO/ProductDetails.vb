

Public Class ProductDetails

        Sub MySub()

            objDetail = New ProductDetails()

        End Sub

#Region "Class attributes declaration"

        Private prhcod As String
        Public Property ProjectNo() As String
            Get
                Return prhcod
            End Get
            Set(ByVal value As String)
                prhcod = value
            End Set
        End Property

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

        Private prdctp As String
        Public Property CTPNo() As String
            Get
                Return prdctp
            End Get
            Set(ByVal value As String)
                prdctp = value
            End Set
        End Property

        Private prdqty As String
        Public Property Qty() As String
            Get
                Return prdqty
            End Get
            Set(ByVal value As String)
                prdqty = value
            End Set
        End Property

        Private prdmfr As String
        Public Property ManufactNo() As String
            Get
                Return prdmfr
            End Get
            Set(ByVal value As String)
                prdmfr = value
            End Set
        End Property

        Private prdcos As String
        Public Property UnitCost() As String
            Get
                Return prdcos
            End Get
            Set(ByVal value As String)
                prdcos = value
            End Set
        End Property

        Private prdcon As String
        Public Property UnitCostNew() As String
            Get
                Return prdcon
            End Get
            Set(ByVal value As String)
                prdcon = value
            End Set
        End Property

        Private prdsts As String
        Public Property Status() As String
            Get
                Return prdsts
            End Get
            Set(ByVal value As String)
                prdsts = value
            End Set
        End Property

        Private prdnew As String
        Public Property NewOrSupplier() As String
            Get
                Return prdnew
            End Get
            Set(ByVal value As String)
                prdnew = value
            End Set
        End Property

        Private vmvnum As String
        Public Property VendorNumber() As String
            Get
                Return vmvnum
            End Get
            Set(ByVal value As String)
                vmvnum = value
            End Set
        End Property

        Private prdmpc As String
        Public Property MinorCode() As String
            Get
                Return prdmpc
            End Get
            Set(ByVal value As String)
                prdmpc = value
            End Set
        End Property

        Private pqmin As String
        Public Property MinQty() As String
            Get
            Return pqmin
        End Get
            Set(ByVal value As String)
            pqmin = value
        End Set
        End Property

        Private objDetail As ProductDetails
    Public Property detailObj() As ProductDetails
        Get
            Return objDetail
        End Get
        Set(ByVal value As ProductDetails)
            objDetail = value
        End Set
    End Property

    Private flagValidationPoqota As String
    Public Property PoqotaValidation() As String
        Get
            Return flagValidationPoqota
        End Get
        Set(ByVal value As String)
            flagValidationPoqota = value
        End Set
    End Property

#End Region

End Class

