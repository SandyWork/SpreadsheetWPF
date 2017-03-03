Imports System.Windows.Interop
Imports System.Windows.Threading


Class userData
    Dim name As String
    Dim selection As String
    Dim attribute1 As String
    Dim attribute2 As String
    Dim attribute3 As String
    Dim attribute4 As String
    Dim attribute5 As String
    Dim attribute6 As String
    Dim attribute7 As String
    Dim attribute8 As String
    Dim attribute9 As String
    Dim attribute10 As String
    Dim unitattri4 As String
    Dim minVal As Integer
    Dim maxVal As Integer
    Dim normVal As Integer

    Public Sub New(name As String, selection As String, attribute1 As String, attribute2 As String, attribute3 As String, attribute4 As String, unitattri4 As String, attribute5 As String, attribute6 As String, attribute7 As String, attribute8 As String, attribute9 As String, attribute10 As String, minVal As Integer, normVal As Integer, maxVal As Integer)

        Me.name = name
        Me.selection = selection
        Me.attribute1 = attribute1
        Me.attribute2 = attribute2
        Me.attribute3 = attribute3
        Me.attribute4 = attribute4
        Me.attribute5 = attribute5
        Me.attribute6 = attribute6
        Me.attribute7 = attribute7
        Me.attribute8 = attribute8
        Me.attribute9 = attribute9
        Me.attribute10 = attribute10
        Me.minVal = minVal
        Me.normVal = normVal
        Me.maxVal = maxVal
        Me.unitattri4 = unitattri4
    End Sub

End Class
Class MainWindow


    Private Sub colSize(sender As Object, e As SizeChangedEventArgs)
        pnl_dock.Width = win_main.ActualWidth
        pnl_dock.Height = win_main.ActualHeight

    End Sub

    Private Sub btn_export_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub btn_import_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub btn_save_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub btn_validate_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub btn_close_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub cntxtmenu_colVertical_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
    End Sub

    Private Sub cntxtmenu_colHorizontal_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)

    End Sub

    Private Sub win_main_Initialized(sender As Object, e As EventArgs)

        Dim obj As userData = New userData("Sundaresh", "none", "1", "2", "3", "3", "3", "3", "3", "3", "3", "3", "3", 1, 2, 3)

    End Sub


End Class