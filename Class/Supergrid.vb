Imports System.ComponentModel
Imports Castle.Components.DictionaryAdapter

Public Class Supergrid : Inherits DataGridView
    Private pagasize As Integer
    Public Property PageSize() As Integer
        Get
            Return pagasize
        End Get
        Set(ByVal value As Integer)
            pagasize = value
        End Set
    End Property

    Public _pageSize As Integer = 10
    Dim bs As BindingSource = New BindingSource()

    Public tables As System.ComponentModel.BindingList(Of DataTable)

    Public Sub New()
        'InitializeComponent()
        ' Instantiate the needed BindingList
        tables = New System.ComponentModel.BindingList(Of DataTable)
    End Sub


    Public Sub SetPagedDataSource(dataTable As DataTable, bnav As BindingNavigator)
        Try
            Dim dt As DataTable = Nothing
            Dim counter As Integer = 1

            For Each dr As DataRow In dataTable.Rows
                If counter = 1 Then
                    dt = dataTable.Clone()
                    tables.Add(dt)
                End If
                dt.Rows.Add(dr.ItemArray)
                If pagasize < ++counter Then
                    counter = 1
                End If
            Next

            bnav.BindingSource = bs
            bs.DataSource = tables
            AddHandler bs.PositionChanged, AddressOf bs_PositionChanged
            'bs.PositionChanged += bs_PositionChanged
            bs_PositionChanged(bs, EventArgs.Empty)
        Catch ex As Exception
            Dim pepe = ex.Message
        End Try

    End Sub

    Public Sub bs_PositionChanged(sender As Object, e As EventArgs)
        Me.DataSource = tables(bs.Position)
    End Sub

End Class
