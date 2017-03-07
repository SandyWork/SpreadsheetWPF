Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading

Namespace gridData
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

        Public Sub New()
            Add(New userData())
        End Sub

    End Class


    Class MainWindow

        Dim collection As PresentData
        Dim obj, obj2 As userData
        Dim rowData(20) As String
        Dim lastCellAddedIndex As Short = 0
        Dim copyActivated As Boolean = False, cutActivated As Boolean = False, pasteActivated As Boolean = False
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

        Private Sub columnHeader_PreviewMouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseRightButtonUp
            If TypeOf sender Is DataGridColumn Then
                Dim header As DataGridColumn = CType(sender, DataGridColumn)
                MsgBox(header.Header)
            ElseIf TypeOf sender Is Primitives.DataGridColumnHeader Then
                Dim header As Primitives.DataGridColumnHeader = CType(sender, Primitives.DataGridColumnHeader)
                MsgBox(header.Content.ToString)
            Else
                Dim header As DataGrid = CType(sender, DataGrid)
            End If

        End Sub

        Public Sub New()

            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.

        End Sub

        Private Sub win_main_Initialized(sender As Object, e As EventArgs)

            obj = New userData("Name", "Selc", "1", "2", "3", "dd", "asd", "dd", "ad", "3", "3", "3", "3", 1, 2, 3)
            collection = New PresentData(obj)
            obj2 = New userData("Name", "Selc", "1", "2", "111", "dd", "asd", "dd", "ad", "3", "3", "3", "3", 1, 2, 3)
            collection.Add(obj2)
            obj2 = New userData("Name", "Selc", "1", "2", "222", "dd", "asd", "dd", "ad", "3", "3", "3", "3", 1, 2, 3)
            collection.Add(obj2)
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


        Private Sub btnCut_Click(sender As Object, e As ExecutedRoutedEventArgs)
            MessageBox.Show("Cut command activated")
            Dim cellsSelected As IList = CType(sender, DataGrid).SelectedCells
            If Not cellsSelected.Count = 0 Then
                If cellsSelected.Count = 1 Then
                    rowData(lastCellAddedIndex) = cellsSelected(0).attribute3
                    MsgBox(rowData(0))
                Else


                End If
            End If

        End Sub

        Private Sub btnPaste_Click(sender As Object, e As ExecutedRoutedEventArgs)
            MessageBox.Show("Paste command activated")
        End Sub

        Private Sub CollectionViewSource_Filter(sender As Object, e As FilterEventArgs)

        End Sub

        Private Sub cbCompleteFilter_Checked(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub cbCompleteFilter_Unchecked(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub CopyCommand(sender As Object, e As ExecutedRoutedEventArgs)

            If Not copyActivated Then
                Dim cellsSelected As Object = CType(sender, DataGrid).SelectedCells
                If Not cellsSelected.Count = 0 Then
                    If cellsSelected.Count = 1 Then
                        If (cellsSelected(0).IsValid) Then
                            Dim header = cellsSelected(0).Column.Header.ToString()
                            Dim userObject As userData = cellsSelected(0).Item

                        End If
                    Else

                    End If
                End If

            End If

        End Sub

        Private Function determineColumn(headerName As String, objectRef As userData) As Object
            Select Case headerName
                Case "Name"
                    Return objectRef.name
                Case "Selection"
                    Return objectRef.selection
                Case "Attribute1"
                    Return objectRef.attribute1
                Case "Attribute2"
                    Return objectRef.attribute2
                Case "Attribute3"
                    Return objectRef.attribute3
                Case "Attribute4"
                    Return objectRef.attribute4
                Case "Unit_Attri4"
                    Return objectRef.unitattri4
                Case "Attribute5"
                    Return objectRef.attribute5
                Case "Attribute6"
                    Return objectRef.attribute6
                Case "Attribute7"
                    Return objectRef.attribute7
                Case "Attribute8"
                    Return objectRef.attribute8
                Case "Attribute9"
                    Return objectRef.attribute9
                Case "Attribute10"
                    Return objectRef.attribute10
                Case "Attribute1"
                    Return objectRef.attribute1
                Case "MinimumVal"
                    Return objectRef.minVal
                Case "NormalVal"
                    Return objectRef.normVal
                Case "MaximumVal"
                    Return objectRef.maxVal
                Case Else
                    MsgBox("Bug Column Name Determination : " + headerName)
                    Return ""
            End Select
        End Function

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
End Namespace
