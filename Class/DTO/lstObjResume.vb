Public Class lstObjResume

    Sub MySub()
        ResumeObj = New objResume()
        LstObjResume.Add(ResumeObj)
    End Sub

    Sub MySub(pepe As objResume)

        ResumeObj = pepe
        LstObjResume.Add(ResumeObj)
    End Sub

    Private _objResume As objResume
    Public Property ResumeObj() As objResume
        Get
            Return _objResume
        End Get
        Set(ByVal value As objResume)
            _objResume = value
        End Set
    End Property

    Private _lstObjResume As List(Of objResume)
    Public Property LstObjResume() As List(Of objResume)
        Get
            Return _lstObjResume
        End Get
        Set(ByVal value As List(Of objResume))
            _lstObjResume = value
        End Set
    End Property

End Class
