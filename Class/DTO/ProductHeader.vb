
Public Class ProductHeader

        Sub MySub()

            _objBase = New ProductHeader()
            _objBase.productDetails = New Details

        End Sub

        'Sub MySub(obj As ProductDetails)

        '    Detail.Add(obj)
        '    'productDetail.Add(obj)

        'End Sub

        'Sub MySub(ByVal objDetail As ProductDetails)

        'End Sub

#Region "Class attributes declaration"

        Dim prhcod As String
        Public Property projectNo() As String
            Get
                Return prhcod
            End Get
            Set(ByVal value As String)
                prhcod = value
            End Set
        End Property

        Dim prdate As String
        Public Property projectDate() As String
            Get
                Return prdate
            End Get
            Set(ByVal value As String)
                prdate = value
            End Set
        End Property

        Private prinfo As String
        Public Property projectInfo() As String
            Get
                Return prinfo
            End Get
            Set(ByVal value As String)
                prinfo = value
            End Set
        End Property

        Private prname As String
        Public Property projectName() As String
            Get
                Return prname
            End Get
            Set(ByVal value As String)
                prname = value
            End Set
        End Property

        Private prstat As String
        Public Property projectStat() As String
            Get
                Return prstat
            End Get
            Set(ByVal value As String)
                prstat = value
            End Set
        End Property

        Private crdate As String
        Public Property creationDate() As String
            Get
                Return crdate
            End Get
            Set(ByVal value As String)
                crdate = value
            End Set
        End Property

        Private cruser As String
        Public Property creationUser() As String
            Get
                Return cruser
            End Get
            Set(ByVal value As String)
                cruser = value
            End Set
        End Property

        Private modate As String
        Public Property modificationDate() As String
            Get
                Return modate
            End Get
            Set(ByVal value As String)
                modate = value
            End Set
        End Property

        Private mouser As String
        Public Property modificationUser() As String
            Get
                Return mouser
            End Get
            Set(ByVal value As String)
                mouser = value
            End Set
        End Property

        Private prpech As String
        Public Property personInCharge() As String
            Get
                Return prpech
            End Get
            Set(ByVal value As String)
                prpech = value
            End Set
        End Property

        Private _objBase As ProductHeader
        Public Property ObjHeader() As ProductHeader
            Get
                Return _objBase
            End Get
            Set(ByVal value As ProductHeader)
                _objBase = value
            End Set
        End Property

        Private productDetails As Details
        Public Property Detail() As Details
            Get
                Return productDetails
            End Get
            Set(ByVal value As Details)
                productDetails = value
            End Set
        End Property

#End Region

    End Class

