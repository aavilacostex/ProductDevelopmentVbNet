Imports System.ComponentModel

Public Class ProductHeaderCollection
    Inherits CollectionBase

    Public Sub New()

    End Sub
    'Public Sub New(ByVal txt As String)
    '    Dim addresses() As String = txt.Split(New Char() _
    '        {";"})

    '    For Each address_text As String In addresses
    '        Try
    '            Me.List.Add(New ProductHeader(address_text))
    '        Catch
    '            Throw New InvalidCastException(
    '                "Invalid StreetAddress serialization '" _
    '                    &
    '                address_text & "'")
    '        End Try
    '    Next address_text
    'End Sub

    Default Public Property Item(ByVal index As Integer) As _
        ProductHeader
        Get
            Return CType(List(index), ProductHeader)
        End Get
        Set(ByVal Value As ProductHeader)
            List(index) = Value
        End Set
    End Property

    Public Sub Add(ByVal product_header As ProductHeader)
        List.Add(product_header)
    End Sub

    Public Function IndexOf(ByVal value As ProductHeader) _
        As Integer
        Return List.IndexOf(value)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal value _
        As ProductHeader)
        List.Insert(index, value)
    End Sub

    Public Sub Remove(ByVal value As ProductHeader)
        List.Remove(value)
    End Sub

    Public Function Contains(ByVal value As ProductHeader) _
        As Boolean
        Return List.Contains(value)
    End Function

    Protected Overrides Sub OnInsert(ByVal index As _
        Integer, ByVal value As Object)
        If Not value.GetType() Is GetType(ProductHeader) _
            Then
            Throw New ArgumentException("Value must be of " &
                "type ProductHeader.", "value")
        End If
    End Sub

    Protected Overrides Sub OnRemove(ByVal index As _
        Integer, ByVal value As Object)
        If Not value.GetType() Is GetType(ProductHeader) _
            Then
            Throw New ArgumentException("Value must be of " &
                "type StreetAddress.", "value")
        End If
    End Sub

    Protected Overrides Sub OnSet(ByVal index As Integer,
        ByVal oldValue As Object, ByVal newValue As Object)
        If Not newValue.GetType() Is GetType(ProductHeader) _
            Then
            Throw New ArgumentException("New value must be " &
                "of type StreetAddress.", "newValue")
        End If
    End Sub

    Protected Overrides Sub OnValidate(ByVal value As _
        Object)
        If Not value.GetType() Is GetType(ProductHeader) _
            Then
            Throw New ArgumentException("Value must be of " &
                "type StreetAddress.", "value")
        End If
    End Sub

    Public Overrides Function ToString() As String
        Dim txt As String
        For Each street_address As ProductHeader In
            MyBase.List
            txt &= ";" & street_address.ToString()
        Next street_address

        If txt.Length > 0 Then txt = txt.Substring(1)

        Return txt
    End Function

End Class


Public Class ProductHeaderCollectionConverter
    Inherits TypeConverter

    Public Overloads Overrides Function _
        CanConvertFrom(ByVal context As _
        System.ComponentModel.ITypeDescriptorContext, ByVal _
        sourceType As System.Type) As Boolean
        If (sourceType.Equals(GetType(String))) Then
            Return True
        Else
            Return MyBase.CanConvertFrom(context,
                sourceType)
        End If
    End Function

    Public Overloads Overrides Function CanConvertTo(ByVal _
        context As _
        System.ComponentModel.ITypeDescriptorContext, ByVal _
        destinationType As System.Type) As Boolean
        If (destinationType.Equals(GetType(String))) Then
            Return True
        Else
            Return MyBase.CanConvertTo(context,
                destinationType)
        End If
    End Function

    'Public Overloads Overrides Function ConvertFrom(ByVal _
    '    context As _
    '    System.ComponentModel.ITypeDescriptorContext, ByVal _
    '    culture As System.Globalization.CultureInfo, ByVal _
    '    value As Object) As Object
    '    If (TypeOf value Is String) Then
    '        Dim txt As String = CType(value, String)
    '        Return New ProductHeaderCollection(txt)
    '    Else
    '        Return MyBase.ConvertFrom(context, culture,
    '            value)
    '    End If
    'End Function

    Public Overloads Overrides Function ConvertTo(ByVal _
        context As _
        System.ComponentModel.ITypeDescriptorContext, ByVal _
        culture As System.Globalization.CultureInfo, ByVal _
        value As Object, ByVal destinationType As _
        System.Type) As Object
        If (destinationType.Equals(GetType(String))) Then
            Return value.ToString()
        Else
            Return MyBase.ConvertTo(context, culture,
                value, destinationType)
        End If
    End Function

    Public Overloads Overrides Function _
        GetPropertiesSupported(ByVal context As _
        ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overloads Overrides Function GetProperties(ByVal _
        context As ITypeDescriptorContext, ByVal value As _
        Object, ByVal Attribute() As Attribute) As _
        PropertyDescriptorCollection
        Return TypeDescriptor.GetProperties(value)
    End Function
End Class