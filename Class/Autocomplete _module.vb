Imports System.Globalization

Module Autocomplete__module

    Public Function loadData() As AutoCompleteStringCollection
        Dim exMessage As String = " "
        Try
            Dim gnr As Gn1 = New Gn1()
            Dim lstResult = gnr.getVendorNoAndNameByName()
            Return lstResult
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function


    Public Sub create_textAutocomplete(txtTexbox As TextBox)
        With txtTexbox

            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            .AutoCompleteCustomSource = loadData()

        End With
    End Sub

    Public Sub create_ddlAutocomplete(cmb As ComboBox)
        With cmb

            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            .AutoCompleteCustomSource = loadData()

        End With
    End Sub

End Module
