Public Class ImportExcel
    Private Sub btn_ok_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub

    Public Function getSheetName() As String
        Return sheetNamesList.SelectedItem.ToString
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Public Sub New(sheetNames As List(Of String))

        ' This call is required by the designer.
        InitializeComponent()
        sheetNamesList.ItemsSource = sheetNames
    End Sub

End Class
