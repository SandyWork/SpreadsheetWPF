﻿Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading
Imports Microsoft.Win32
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Data

Namespace gridData
    Public Class userData : Implements INotifyPropertyChanged
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
            MsgBox("Editing")
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
        End Sub

    End Class

    Public Class PresentData
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
        Dim rowData(20) As String, headerSelected As String = ""
        Dim lastCellAddedIndex As Short = 0, rowIndex As Short = 0, columnIndex As Short = 0
        Dim rowEditIndex As Integer, colEditIndex As Integer = 0
        Dim filterValue As String = ""
        Dim copyActivated As Boolean = False, cutActivated As Boolean = False, pasteActivated As Boolean = False
        Private filterSelected As Boolean = False


        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

        End Sub

        Private Sub colSize(sender As Object, e As SizeChangedEventArgs)

            pnl_dock.Width = win_main.ActualWidth
            pnl_dock.Height = win_main.ActualHeight


        End Sub

        Private Sub columnHeader_PreviewMouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseRightButtonUp

            Dim dep As DependencyObject = e.OriginalSource

            While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridCell) AndAlso Not (TypeOf dep Is Primitives.DataGridColumnHeader))
                dep = VisualTreeHelper.GetParent(dep)
            End While

            If dep Is Nothing Then
                Return
            End If

            If (TypeOf dep Is Primitives.DataGridColumnHeader) Then
                Dim header As Primitives.DataGridColumnHeader = dep
                headerSelected = header.Content.ToString
            End If

            If (TypeOf dep Is DataGridColumn) Then
                Dim header As DataGridColumn = dep
                headerSelected = header.Header.ToString
            End If

            If (TypeOf dep Is DataGridCell) Then
                Dim cell As DataGridCell = dep
                While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridRow))
                    dep = VisualTreeHelper.GetParent(dep)

                End While
                Dim row As DataGridRow = dep

                rowIndex = FindRowIndex(row)
                columnIndex = cell.Column.DisplayIndex

                ''MsgBox(rowIndex & " " & columnIndex)
                ''If needed to find header of that row/column
                'headerSelected = cell.Column.Header.ToString
            End If
            e.Handled = False
            Return
        End Sub

        ''All the Helper Functions

        Private Function FindRowIndex(row As DataGridRow) As Integer
            Dim dataGrid As DataGrid = ItemsControl.ItemsControlFromItemContainer(row)
            Dim index As Integer = dataGrid.ItemContainerGenerator.IndexFromContainer(row)
            Return index
        End Function


        Private Function determineColumn(headerName As String, objectRef As userData) As Object
            headerName = headerName.Trim
            headerName = headerName.ToLower()
            Select Case headerName
                Case "name"
                    Return objectRef.name
                Case "selection"
                    Return objectRef.selection
                Case "attribute1"
                    Return objectRef.attribute1
                Case "attribute2"
                    Return objectRef.attribute2
                Case "attribute3"
                    Return objectRef.attribute3
                Case "attribute4"
                    Return objectRef.attribute4
                Case "unit_attri4"
                    Return objectRef.unitattri4
                Case "attribute5"
                    Return objectRef.attribute5
                Case "attribute6"
                    Return objectRef.attribute6
                Case "attribute7"
                    Return objectRef.attribute7
                Case "attribute8"
                    Return objectRef.attribute8
                Case "attribute9"
                    Return objectRef.attribute9
                Case "attribute10"
                    Return objectRef.attribute10
                Case "minimumval"
                    Return objectRef.minVal
                Case "normalval"
                    Return objectRef.normVal
                Case "maximumval"
                    Return objectRef.maxVal
                Case Else
                    MsgBox("Bug Column Name Determination : " + headerName)
                    Return ""
            End Select
        End Function

        ''End of Helper Functions
        Private Sub win_main_Initialized(sender As Object, e As EventArgs)
            collection = Me.Resources("presentData")
            collection.Clear()

            obj = New userData("Name", "10", "1", "10", "3", "dd", "asd", "dd", "ad", "2", "20", "3", "3", 1, 2, 3)
            collection.Add(obj)
            obj2 = New userData("Name", "abc", "1", "abc", "111", "dd", "abc", "dd", "abc", "435", "2", "3", "3", 5, 6, 7)
            collection.Add(obj2)
            obj2 = New userData("Something", "12", "12", "2", "222", "dd", "12", "12", "ad", "22", "1", "3", "12", 1, 2, 3)
            collection.Add(obj2)
            ''dg_grid1.ItemsSource = collection


        End Sub

        Private Sub btn_edit_row_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = 16 Then
                collection.Insert(rowEditIndex + 1L, New userData())
            End If

        End Sub

        Private Sub btn_delete_row_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = 17 Then
                collection.RemoveAt(rowEditIndex)
                If collection.Count = 0 Then
                    collection.Add(New userData())
                End If
            End If

        End Sub



        Private Sub CollectionViewSource_Filter(sender As Object, e As FilterEventArgs)
            If filterSelected Then
                Dim obj As userData = e.Item
                If obj IsNot Nothing Then
                    ' If filter is turned on, filter completed items.
                    If Me.cbCompleteFilter.IsChecked = True Then
                        Dim temp As Object = determineColumn(headerSelected, obj)
                        If temp.Equals(filterValue) Then
                            e.Accepted = True

                        Else
                            e.Accepted = False
                        End If
                    End If
                End If
            End If

        End Sub

        Private Sub menuFilter_Click(sender As Object, e As RoutedEventArgs)

            Dim inputDialog As FilterWindow = New FilterWindow()
            inputDialog.ShowInTaskbar = True
            inputDialog.Owner = Me

            If inputDialog.ShowDialog = True Then
                filterSelected = True
                cbCompleteFilter.IsEnabled = True
                cbCompleteFilter.IsChecked = True
                filterValue = inputDialog.returnFilterValue()
                CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                filterStatus.Content = "Currently Filter is applied to Column : " & headerSelected & " with Value : " & filterValue
            End If
        End Sub

        Private Sub CutCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)
            If Not cutActivated Then
                If TypeOf sender Is DataGrid Then
                    Dim cellsSelected As Object = CType(sender, DataGrid).SelectedCells
                    If Not cellsSelected.Count = 0 Then
                        If cellsSelected.Count = 1 Then
                            If (cellsSelected(0).IsValid) Then
                                Dim header = cellsSelected(0).Column.Header.ToString()
                                Dim userObject As userData = cellsSelected(0).Item
                                Dim elementSelected As Object = determineColumn(header, userObject)
                                MsgBox(elementSelected)
                            End If
                        ElseIf cellsSelected > 1 And cellsSelected < dg_grid1.Columns.Count Then

                        End If
                    End If
                End If

            End If

        End Sub

        Private Sub CopyCommand(sender As Object, e As ExecutedRoutedEventArgs)

            If Not copyActivated Then
                If TypeOf sender Is DataGrid Then
                    Dim cellsSelected As Object = CType(sender, DataGrid).SelectedCells
                    If Not cellsSelected.Count = 0 Then
                        If cellsSelected.Count = 1 Then
                            If (cellsSelected(0).IsValid) Then
                                Dim header = cellsSelected(0).Column.Header.ToString()
                                Dim userObject As userData = cellsSelected(0).Item
                                Dim elementSelected As Object = determineColumn(header, userObject)
                                MsgBox(elementSelected)
                            End If
                        ElseIf cellsSelected > 1 And cellsSelected < dg_grid1.Columns.Count Then

                        End If
                    End If

                End If
            End If


        End Sub


        Private Sub PasteCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)

        End Sub

        Private Sub addRowHeader_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = 16 Then
                collection.Insert(rowEditIndex + 1L, New userData())
            End If
        End Sub

        Private Sub deleteRowHeader_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = 17 Then
                collection.RemoveAt(rowEditIndex)
                If collection.Count = 0 Then
                    collection.Add(New userData())
                End If
            End If

        End Sub

        Private Sub copyCells_Click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub pasteCells_Click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub highlightCells_Click(sender As Object, e As RoutedEventArgs)
            If columnIndex = 1 Then
                Dim obj As userData = collection.Item(rowIndex)
                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(1))
                If Not dg_grid1.SelectedCells.Contains(dg_grid1.CurrentCell) Then
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If

                If obj.selection.Equals(obj.attribute1) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(2))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute2) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(3))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute3) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(4))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute4) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(5))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.unitattri4) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(6))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute5) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(7))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute6) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(8))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute7) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(9))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute8) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(10))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute9) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(11))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If
                If obj.selection.Equals(obj.attribute10) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(12))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)
                End If

            End If
        End Sub

        Private Sub CompleteFilter_Changed(sender As Object, e As RoutedEventArgs)
            CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
        End Sub

        Private Sub detectCellClicked(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseLeftButtonUp

            Dim dep As DependencyObject = e.OriginalSource

            While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridCell) AndAlso Not (TypeOf dep Is Primitives.DataGridColumnHeader))
                dep = VisualTreeHelper.GetParent(dep)
            End While

            If dep Is Nothing Then
                Return
            End If

            If (TypeOf dep Is DataGridCell) Then
                Dim cell As DataGridCell = dep
                While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridRow))
                    dep = VisualTreeHelper.GetParent(dep)
                End While
                Dim row As DataGridRow = dep

                rowEditIndex = FindRowIndex(row)
                colEditIndex = cell.Column.DisplayIndex
            End If
            e.Handled = False
            Return

        End Sub

        ''Yet to be Impleted Codes
        Private Sub btn_export_Click(sender As Object, e As RoutedEventArgs)

            '' Main Content
            Dim f As SaveFileDialog = New SaveFileDialog()
            f.Filter = "Excel Workbook (*.xlsx) |*.xlsx|All files (*.*)|*.*"
            Try
                If f.ShowDialog() = True Then
                    Dim xlApp As Excel.Application = New Excel.Application()
                    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add
                    Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(1)
                    Dim colCount = dg_grid1.Columns.Count, rowCount = collection.Count

                    'Create an array with 16 columns and n rows
                    Dim DataArray(rowCount, colCount) As Object
                    Dim headersList(colCount) As Object

                    Dim counter As Integer = 0
                    For Each col In dg_grid1.Columns
                        headersList(counter) = col.Header.ToString
                        counter += 1
                    Next

                    For row As Short = 0 To rowCount - 1
                        For col As Short = 0 To colCount - 3
                            Dim temp As Object = determineColumn(headersList(col), collection.Item(row))
                            DataArray(row, col) = temp
                        Next
                    Next
                    xlWorkSheet.Range("A1").Resize(1, colCount).Value = headersList
                    xlWorkSheet.Range("A2").Resize(rowCount, colCount).Value = DataArray


                    xlWorkSheet.SaveAs(f.FileName)
                    xlWorkBook.Close()
                    xlApp.Quit()

                    releaseObject(xlApp)
                    releaseObject(xlWorkBook)
                    releaseObject(xlWorkSheet)
                    MsgBox("Exported")
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK)
            End Try
        End Sub

        Private Sub releaseObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End Sub

        Private Sub btn_import_Click(sender As Object, e As RoutedEventArgs)
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog()
            openFileDialog.Filter = "Excel 97-2003 Worksheet (*.xls)|*.xls|Excel Workbook (*.xlsx) |*.xlsx|All files (*.*)|*.*"
            openFileDialog.Multiselect = False

            If (openFileDialog.ShowDialog() = True) Then
                If openFileDialog.CheckFileExists = True Then
                    Dim filename = openFileDialog.FileName


                    Dim sheetNames As List(Of String) = GetExcelSheetNames(filename)
                    Dim excel As ImportExcel = New ImportExcel(sheetNames)
                    excel.ShowInTaskbar = True
                    excel.Owner = Me

                    If excel.ShowDialog = True Then
                        Dim sheetName As String = excel.getSheetName()
                        displayExcelFile(filename, sheetName)
                    End If

                    ''displayExcelFile(filename)
                End If

            End If
        End Sub

        Private Sub btn_save_Click(sender As Object, e As RoutedEventArgs)
        End Sub

        Private Function GetExcelSheetNames(ByVal fileName As String) As List(Of String)
            Dim strconn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
          fileName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            Dim conn As New OleDb.OleDbConnection(strconn)

            conn.Open()

            Dim dtSheets As System.Data.DataTable =
              conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim listSheet As New List(Of String)
            Dim drSheet As DataRow

            For Each drSheet In dtSheets.Rows
                Dim temp As String = drSheet("TABLE_NAME").ToString()
                listSheet.Add(temp.Substring(0, temp.Length - 1))
            Next
            conn.Close()
            Return listSheet

        End Function

        Private Sub displayExcelFile(filename As String, sheetName As String)

            Dim excelApp As Excel.Application = New Excel.Application()
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(filename)
            Dim worksheet As Excel.Worksheet = workbook.Sheets(sheetName)

            Dim col As Integer = 0, row As Integer = 0
            Dim range As Excel.Range = worksheet.UsedRange
            MsgBox(range.Rows.Count)

            workbook.Save()
            workbook.Close()
            excelApp.Quit()
            'For row = 2 To range.Rows.Count
            '    For col = 1 To range.Columns.Count
            '        workbook.Close(True, Missing.Value, Missing.Value);
            '            excelApp.Quit();
            '            Return dt.DefaultView;
            '    Next
            'Next

        End Sub


        Private Sub btn_validate_Click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub btn_close_Click(sender As Object, e As RoutedEventArgs)

        End Sub

    End Class
End Namespace
