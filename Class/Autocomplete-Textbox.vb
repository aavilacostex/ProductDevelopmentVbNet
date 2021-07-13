Public Class Autocomplete_Textbox
    Inherits TextBox

    Private _listBox As ListBox
    Dim lstHelper As List(Of String) = New List(Of String)
    Private _isAdded As Boolean
    Private _values As String()
    Private _formerValue As String = Nothing

    Public Sub New()
        InitializeComponent()
        ResetListBox()
    End Sub

    Private Sub InitializeComponent()
        _listBox = New ListBox()
        AddHandler Me.KeyDown, AddressOf txt_KeyDown
        AddHandler Me.KeyUp, AddressOf txt_KeyUp
    End Sub

    'Private _values As String
    Public Property Values() As String()
        Get
            Return _values
        End Get
        Set(ByVal value As String())
            _values = value
        End Set
    End Property

    Private Sub ShowListBox()
        If Not _isAdded Then
            Parent.Controls.Add(_listBox)
            _listBox.Left = Left
            _listBox.Top = Top + Height
            _isAdded = True
        End If
        '_listBox.Visible = True
        '_listBox.BringToFront()
    End Sub

    Private Sub ResetListBox()
        _listBox.Visible = False
    End Sub

    Private Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim exMessage As String = Nothing
        Try
            Select Case e.KeyCode
                Case Keys.Tab
                    If _listBox.Visible Or _listBox.Items.Count > 0 Then
                        InsertWord(_listBox.SelectedItem.ToString())
                        ResetListBox()
                        _formerValue = Text
                    End If
                    Exit Select
                Case Keys.Down
                    If (_listBox.Visible  Or _listBox.Items.Count > 0) And (_listBox.SelectedIndex < _listBox.Items.Count - 1) Then
                        _listBox.SelectedIndex += 1
                    End If
                    Exit Select
                Case Keys.Up
                    If (_listBox.Visible) And (_listBox.SelectedIndex > 0) Then
                        _listBox.SelectedIndex -= 1
                    End If
                    Exit Select
            End Select
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Sub txt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        UpdateListBox()
    End Sub

    Private Function GetItemHeight(lst As ListBox, pos As Integer) As Integer
        Return lst.GetItemHeight(pos)
    End Function

    Private Sub UpdateListBox()
        Dim exMessage As String = Nothing
        Try
            If (Text = _formerValue) Then
                Return
            End If
            _formerValue = Text
            Dim word As String = GetWord()

            If (_values IsNot Nothing And word.Length > 0) Then

                lstHelper = toListMethod(_values, Nothing, 0)
                Dim lstMatches As List(Of String) = New List(Of String)

                'lstMatches = lstHelper.FindAll(Function(x) _
                '                        x.StartsWith(word, StringComparison.OrdinalIgnoreCase) And Not SelectedValues.Contains(x))

                For Each item As String In lstHelper
                    If item IsNot Nothing Then
                        If (item.StartsWith(word, StringComparison.OrdinalIgnoreCase) Or UCase(item).Contains(UCase(word))) And Not SelectedValues.Contains(word) Then
                            lstMatches.Add(item)
                        End If
                    End If
                Next

                'Dim strMatches = toListMethod(Nothing, lstMatches, 1)

                'If (strMatches.Length > 0) Then
                If (lstMatches.Count > 0) Then
                    ShowListBox()
                    _listBox.Items.Clear()

                    'For Each item As String In strMatches
                    For Each item As String In lstMatches
                        _listBox.Items.Add(item)
                    Next

                    'Array.ForEach(strMatches, Function(x) _listBox.Items.Add(x))
                    _listBox.SelectedIndex = 0
                    _listBox.Height = 0
                    _listBox.Width = 0
                    'OnGotFocus()

                    Using graphics As Graphics = _listBox.CreateGraphics()
                        Dim i As Integer = 0
                        _listBox.Name = "lstMatches"
                        For Each item As String In _listBox.Items
                            _listBox.Height += GetItemHeight(_listBox, i)
                            Dim itemWidth As Integer = graphics.MeasureString(item, _listBox.Font).Width
                            _listBox.Width = If(_listBox.Width < itemWidth, itemWidth, _listBox.Width)
                        Next
                    End Using

                    Dim dtBase = DirectCast(frmLoadExcel.ComboBox1.DataSource, DataTable)
                    Dim dtCmb = fromListboxToDatatable(_listBox, dtBase)

                    Dim newRow As DataRow = dtCmb.NewRow
                    newRow("VMNAME") = ""
                    newRow("VMVNUM") = -1
                    'dsUser.Tables(0).Rows.Add(newRow)
                    dtCmb.Rows.InsertAt(newRow, 0)

                    frmLoadExcel.ComboBox2.DataSource = dtCmb
                    frmLoadExcel.ComboBox2.DisplayMember = "VMNAME"
                    frmLoadExcel.ComboBox2.ValueMember = "VMVNUM"

                    frmLoadExcel.ComboBox2.SelectedIndex = 1
                    'frmLoadExcel.lblVendorDesc.Text = Nothing
                Else
                    _listBox.Items.Clear()
                    frmLoadExcel.ComboBox2.DataSource = Nothing
                    frmLoadExcel.txtVendorNo.Text = Nothing
                    frmLoadExcel.lblVendorDesc.Text = Nothing
                    ResetListBox()
                End If
            Else
                _listBox.Items.Clear()
                frmLoadExcel.ComboBox2.DataSource = Nothing
                frmLoadExcel.txtVendorNo.Text = Nothing
                frmLoadExcel.lblVendorDesc.Text = Nothing
                ResetListBox()
            End If
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Function toListMethod(Optional strArray As String() = Nothing, Optional strListString As List(Of String) = Nothing, Optional flag As Integer = 0)

        If flag = 0 Then
            'convert to list of string from array
            Dim lstValues = New List(Of String)
            For Each item As String In strArray
                lstValues.Add(item)
            Next
            Return lstValues
        Else
            'convert to array of string from list
            Dim lenght = strListString.Count
            Dim i As Integer = 0
            Dim strValues = New String(lenght - 1) {}

            For Each item As String In strListString
                strValues(i) = item
                i += 1
            Next
            Return strValues
        End If


    End Function

    Protected Overridable Function IsInputKey(keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Tab
                Return True
            Case Else
                Return Me.IsInputKey(keyData)
        End Select
    End Function

    Private Function GetWord() As String
        Dim exMessage As String = Nothing
        Try
            Dim _text As String = Text
            Dim pos As Integer = Me.SelectionStart
            Dim posStart As Integer = Text.LastIndexOf(" ", If(pos < 1, 0, pos - 1))
            posStart = If(posStart = -1, 0, posStart + 1)
            Dim posEnd As Integer = Text.IndexOf(" ", pos)
            posEnd = If(posEnd = -1, Text.Length, posEnd)

            Dim lenght As Integer = If((posEnd - posStart) < 0, 0, posEnd - posStart)

            Return Text.Substring(posStart, lenght)
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Sub InsertWord(newTag As String)
        Dim exMessage As String = Nothing
        Try
            Dim _text As String = Text
            Dim pos As Integer = Me.SelectionStart
            Dim posStart As Integer = Text.LastIndexOf(" ", If(pos < 1, 0, pos - 1))
            posStart = If(posStart = -1, 0, posStart + 1)
            Dim posEnd As Integer = Text.IndexOf(" ", pos)

            Dim firstPart As String = Text.Substring(0, posStart) + newTag
            Dim updatedText As String = firstPart + (If(posEnd = -1, "", Text.Substring(posEnd, Text.Length - posEnd)))

            Text = updatedText
            Me.SelectionStart = firstPart.Length

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try
    End Sub

    Private Function fromListboxToDatatable(lst As ListBox, Optional dtBase As DataTable = Nothing) As DataTable
        Dim exMessage As String = Nothing
        Try
            Dim dt As DataTable = New DataTable()
            Dim column1 As DataColumn = New DataColumn("VMNAME")
            column1.DataType = System.Type.GetType("System.String")
            Dim column2 As DataColumn = New DataColumn("VMVNUM")
            column2.DataType = System.Type.GetType("System.Decimal")

            dt.Columns.Add(column1)
            dt.Columns.Add(column2)

            If lst.Items.Count > 0 Then
                Dim x As Integer = 0
                Dim vndCode As String = Nothing
                For Each item As String In lst.Items
                    vndCode = If(Integer.TryParse(utilDT(item, dtBase), x), CInt(utilDT(item, dtBase)), 0)
                    If vndCode IsNot Nothing Then
                        Dim row = dt.NewRow()
                        row("VMNAME") = item
                        row("VMVNUM") = vndCode
                        dt.Rows.Add(row)
                    End If
                Next
                dt.AcceptChanges()
            End If
            Return dt
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private Function utilDT(value As String, dtvalues As DataTable) As String
        Dim exMessage As String = Nothing
        Dim code As String = Nothing
        Try
            If value IsNot Nothing And dtvalues IsNot Nothing Then
                For Each item As DataRow In dtvalues.Rows
                    If item.ItemArray(0).ToString().Equals(value) Then
                        code = item.ItemArray(1).ToString()
                        Exit For
                    End If
                Next
            End If
            Return code
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
            Return Nothing
        End Try
    End Function

    Private SelectedValues As List(Of String)
    Public Property lstSelectedValues() As List(Of String)
        Get
            Dim result As String() = Text.Split(New String() {" "}, StringSplitOptions.RemoveEmptyEntries)
            SelectedValues = New List(Of String)(result)
            Return SelectedValues
        End Get
        Set(ByVal value As List(Of String))
            SelectedValues = value
        End Set
    End Property


End Class
