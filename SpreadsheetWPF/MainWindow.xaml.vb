Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading

Class userData : Implements INotifyPropertyChanged
    Public Property name As String
    Public Property selection As String
    Public Property attribute1 As String
    Public Property attribute2 As String
    Public Property attribute3 As String
    Public Property attribute4 As String
    Public Property attribute5 As String
    Public Property attribute6 As String
    Public Property attribute7 As String
    Public Property attribute8 As String
    Public Property attribute9 As String
    Public Property attribute10 As String
    Public Property unitattri4 As String
    Public Property minVal As Nullable(Of Integer)
    Public Property maxVal As Nullable(Of Integer)
    Public Property normVal As Nullable(Of Integer)

    Public Sub New(Optional name As String = "", Optional selection As String = "", Optional attribute1 As String = "", Optional attribute2 As String = "", Optional attribute3 As String = "", Optional attribute4 As String = "", Optional unitattri4 As String = "", Optional attribute5 As String = "", Optional attribute6 As String = "", Optional attribute7 As String = "", Optional attribute8 As String = "", Optional attribute9 As String = "", Optional attribute10 As String = "", Optional minVal As Nullable(Of Integer) = Nothing, Optional normVal As Nullable(Of Integer) = Nothing, Optional maxVal As Nullable(Of Integer) = Nothing)

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

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

End Class

Class PresentData
    Inherits ObservableCollection(Of userData)

    Public Sub New(obj As userData)
        Add(obj)
    End Sub

End Class


Class MainWindow

    Dim collection As PresentData
    Dim obj As userData

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


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub win_main_Initialized(sender As Object, e As EventArgs)

        obj = New userData("Name", "Selc", "1", "2", "3", "dd", "asd", "dd", "ad", "3", "3", "3", "3", 1, 2, 3)
        collection = New PresentData(obj)
        dg_grid1.ItemsSource = collection
    End Sub

    Private Sub dg_grid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dg_grid1.SelectionChanged

    End Sub

    Private Sub btn_edit_row_Click(sender As Object, e As RoutedEventArgs)
        collection.Add(New userData())
    End Sub

    Private Sub btn_delete_row_Click(sender As Object, e As RoutedEventArgs)

    End Sub


    Private Sub detectHeader(sender As Object, e As MouseButtonEventArgs)
        'Dim dg_demoColumn As DataGridColumn = sender
        'If Not dg_demoColumn.Header Is Nothing Then
        '    MsgBox(dg_demoColumn.Header)
        'End If
    End Sub

    Private Sub dg_grid1_LoadingRow(sender As Object, e As DataGridRowEventArgs)

    End Sub

    'Private Sub columnHeader_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseRightButtonUp
    '    If TypeOf sender Is DataGrid Then
    '        MsgBox("Grid")
    '    ElseIf TypeOf sender Is DataGridColumn Then
    '        MsgBox("Column")
    '    ElseIf TypeOf sender Is DataGridRow Then
    '        MsgBox("Row")
    '    Else
    '        MsgBox("GridArea")
    '    End If


    'End Sub
End Class