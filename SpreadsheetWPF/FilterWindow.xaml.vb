Public Class FilterWindow
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub button_Click(sender As Object, e As RoutedEventArgs)
        If filterValue.Text.Equals("") Then
            errorStatus.Visibility = Visibility.Visible
        Else
            Me.DialogResult = True
        End If

    End Sub

    Public Function returnFilterValue() As String
        Return filterValue.Text.ToString
    End Function

    Private Sub filterValue_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs) Handles filterValue.GotKeyboardFocus
        errorStatus.Visibility = Visibility.Hidden
    End Sub
End Class
