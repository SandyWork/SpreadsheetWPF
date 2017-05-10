Public Class FilterWindow

    Public Property populate As List(Of String) = New List(Of String)

    Dim status As String = ""
    Public Sub New(header As String)

        ' This call is required by the designer.
        InitializeComponent()
        Me.status = header
        lblCursorPosition.Content += Me.status
        textPanel.Visibility = Visibility.Visible
        comboPanel.Visibility = Visibility.Hidden

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    Public Sub New(_populate_ As List(Of String), _status_ As String)
        InitializeComponent()

        _populate_.Sort()
        Me.populate = _populate_
        Me.status = _status_
        cmb_List.ItemsSource = Me.populate

        lblCursorPosition.Content += _status_

        textPanel.Visibility = Visibility.Hidden
        comboPanel.Visibility = Visibility.Visible
    End Sub

    Private Sub button_Click(sender As Object, e As RoutedEventArgs)

        If populate.Count = 0 Then
            If filterValue.Text.Equals("") Then
                errorPanel.Visibility = Visibility.Visible
                errorStatus.Visibility = Visibility.Visible
            Else
                Me.DialogResult = True
            End If
        Else
            If Not cmb_List.SelectedValue.Equals("") Then
                Me.DialogResult = True
            Else
                errorPanel.Visibility = Visibility.Visible
                errorStatus.Visibility = Visibility.Visible
            End If
        End If

    End Sub

    Public Function returnFilterValue() As String
        If populate.Count = 0 Then
            Return filterValue.Text.ToString
        Else
            Return cmb_List.SelectedValue.ToString()
        End If
    End Function

    Public Function returnCaseSensitive() As Boolean
        Return chk_caseSenst.IsChecked
    End Function

    Private Sub filterValue_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs) Handles filterValue.GotKeyboardFocus
        errorStatus.Visibility = Visibility.Hidden
        errorPanel.Visibility = Visibility.Hidden
    End Sub
End Class
