Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading
Imports Microsoft.Win32
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Data
Imports System.IO
Imports System.Collections.Specialized
'This is comment
Namespace gridData
    Public Class userData : Implements INotifyPropertyChanged, IEditableObject
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

        'Public Property PhoneNumber() As String
        '    Get
        '        Return Me.attribute10
        '    End Get
        '    Set
        '        If Value <> Me.attribute10 Then
        '            Me.attribute10 = Value
        '            OnPropertyChanged("Attribute10")
        '        End If
        '    End Set
        'End Property

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

        Public Sub BeginEdit() Implements IEditableObject.BeginEdit
            ''Console.WriteLine("Editing")
        End Sub

        Public Sub EndEdit() Implements IEditableObject.EndEdit
            ''Console.WriteLine("End")
        End Sub

        Public Sub CancelEdit() Implements IEditableObject.CancelEdit
            ''Console.WriteLine("Cancel Editing")
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

        Protected Overrides Sub OnCollectionChanged(e As NotifyCollectionChangedEventArgs)
            MyBase.OnCollectionChanged(e)
        End Sub
    End Class

    Class MainWindow

        Dim collection As PresentData

        'header of the column where user right clicked
        Dim headerSelected As String = "", configurationFileName As String = "..\selectionConfigFile.txt"

        'rowIndex and columnIndex stores the location of the cell where user right clicked to open the context Menu
        Dim rowIndex As Short = 0, columnIndex As Short = 0

        'rowEditIndex, colEditIndex stores the location of the cell where user left clicked
        'This is used for Add Row and Delete Row Buttons to know exactly where users wants to perform Row operations

        Dim rowEditIndex As Integer, colEditIndex As Integer = 0

        'Filter Value that user entered when prompted
        Dim filterValue As String = ""
        Dim copyActivated As Boolean = False, cutActivated As Boolean = False, pasteActivated As Boolean = False
        Private filterSelected As Boolean = False
        Dim configHeaderList As List(Of List(Of String)) = New List(Of List(Of String))()
        Dim configIndexCount As Integer = 0
        Dim redcellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim greencellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim blueCellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim errorHighlight As Boolean = False, valHighlight As Boolean = False, blueHighlight As Boolean = False, blueFlag As Boolean = False
        Dim btnNamesArray() As String = {"btn_filter_name", "btn_filter_sel", "btn_filter_attri1", "btn_filter_attri2", "btn_filter_attri3", "btn_filter_attri4", "btn_filter_unitattri4", "btn_filter_attri5", "btn_filter_attri6", "btn_filter_attri7", "btn_filter_attri8", "btn_filter_attri9", "btn_filter_attri10", "btn_filter_minval", "btn_filter_normval", "btn_filter_maxval"}

        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            collection = Me.Resources("presentData")
            collection.Clear()

            Dim obj = New userData("Name", "10", "1", "10", "3", "dd", "asd", "dd", "ad", "2", "20", "3", "3", 1, 2, 3)
            collection.Add(obj)
            Dim obj2 = New userData("Name", "abc", "1", "abc", "111", "dd", "abc", "dd", "abc", "435", "2", "3", "3", 5, 6, 7)
            collection.Add(obj2)
            obj2 = New userData("Something", "12", "12", "2", "222", "dd", "12", "12", "ad", "22", "1", "3", "12", 1, 2, 3)
            collection.Add(obj2)
        End Sub

        Public Sub New(list As userData())

            ' This call is required by the designer.
            InitializeComponent()
            collection = Me.Resources("presentData")
            collection.Clear()

            If list.Count > 0 Then
                For Each element In list
                    collection.Add(element)
                Next
            Else
                Dim obj = New userData("Name", "10", "1", "10", "3", "dd", "asd", "dd", "ad", "2", "20", "3", "3", 1, 2, 3)
                collection.Add(obj)
                Dim obj2 = New userData("Name", "abc", "1", "abc", "111", "dd", "abc", "dd", "abc", "435", "2", "3", "3", 5, 6, 7)
                collection.Add(obj2)
                obj2 = New userData("Something", "12", "12", "2", "222", "dd", "12", "12", "ad", "22", "1", "3", "12", 1, 2, 3)
                collection.Add(obj2)
            End If

        End Sub

        'This method is used to automatically resize the dock panel when user resizes Window
        Private Sub resizeWindow(sender As Object, e As SizeChangedEventArgs)

            pnl_dock.Width = win_main.ActualWidth
            pnl_dock.Height = win_main.ActualHeight


        End Sub


        ''**********All the Helper Functions

        Private Function FindRowIndex(row As DataGridRow) As Integer
            Dim dataGrid As DataGrid = ItemsControl.ItemsControlFromItemContainer(row)
            Dim index As Integer = dataGrid.ItemContainerGenerator.IndexFromContainer(row)
            Return index
        End Function

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
                    MsgBox("Bug in determineColumn : " + headerName)
                    Return ""
            End Select
        End Function


        Private Sub fileRead()
            Console.WriteLine("fileread")
            Try
                'Check if an existing configuration file exists.
                'If not create a new one with default Values
                If File.Exists(configurationFileName) = False Then

                    Console.WriteLine("No Configuration File")
                    Console.WriteLine("Creating One With Default Values")

                    Using sw As StreamWriter = File.CreateText(configurationFileName)
                        sw.WriteLine("<-- Format to specify columns to highlight for specific Value is : -->")
                        sw.WriteLine("<-- [Value] [List of Columns separated by "" "" ]                  -->")
                        sw.WriteLine("<-- e.g.  1 Attribute1 Attribute2 Attribute3                       -->")
                        sw.WriteLine("1 Attribute8 Attribute9 Attribute10")
                        sw.WriteLine("2 Attribute6 Attribute7 Attribute8")
                        sw.WriteLine("3 Attribute1 Attribute4 Attribute10")
                        sw.Flush()
                    End Using
                End If
                Console.WriteLine("Already there")
                ' Open the file to read from.
                Using sr As StreamReader = File.OpenText(configurationFileName)
                    Dim lineCount As Integer = File.ReadLines(configurationFileName).Count()
                    Dim temp As String = ""
                    Dim attributeList As String()
                    While sr.Peek() >= 0
                        temp = sr.ReadLine()
                        If Not temp.Contains("<--") Then
                            configHeaderList.Add(New List(Of String)())

                            attributeList = temp.Split(" ")
                            For counter As Integer = 0 To attributeList.Count - 1
                                configHeaderList(configIndexCount).Add(attributeList(counter))
                            Next
                            configIndexCount += 1
                        End If
                    End While
                End Using

            Catch e As Exception
                ' Let the user know what went wrong.
                Console.WriteLine("The file could not be read:")
                Console.WriteLine(e.Message)
            End Try
        End Sub

        Public Function GetDataGridCell(cellInfo As DataGridCellInfo) As DataGridCell
            Dim cellContent As Object = cellInfo.Column.GetCellContent(cellInfo.Item)
            If (cellContent IsNot Nothing) Then
                Return CType(cellContent.Parent, DataGridCell)
            End If
            Return Nothing
        End Function

        Public Sub changeCellColor(dataCellInfo As DataGridCellInfo, bgColor As Color, fgColor As Color)
            Dim dataCell As DataGridCell = GetDataGridCell(dataCellInfo)

            If bgColor.Equals(Colors.Black) Then
                dataCell.BorderThickness = New Thickness(0.0)
            Else
                dataCell.BorderThickness = New Thickness(3.0)
            End If

            If bgColor = Colors.DarkRed Then
                greencellsColored.Add(dataCell)
                valHighlight = True
            End If

            If bgColor = Colors.Blue Then
                blueCellsColored.Add(dataCell)
                valHighlight = True
            End If

            If bgColor = Colors.Red Then
                redcellsColored.Add(dataCell)
                errorHighlight = True
            End If

            dataCell.BorderBrush = New SolidColorBrush(bgColor)
        End Sub

        ''**********End of Helper Functions

        Private Sub win_main_Initialized(sender As Object, e As EventArgs)

        End Sub


        'This method determines the where user performed mouse right click and stores the location of click.
        Private Sub columnHeader_PreviewMouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseRightButtonUp

            HighlightClear()
            Dim dep As DependencyObject = e.OriginalSource

            'No Matter where user click inside the Datagrid, be it row or column or header, the sender is always DataGrid.
            'Hence we must find the parent element where user actually clicked.
            While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridCell) AndAlso Not (TypeOf dep Is Primitives.DataGridColumnHeader))
                dep = VisualTreeHelper.GetParent(dep)
            End While


            If dep Is Nothing Then
                Return

            ElseIf (TypeOf dep Is Primitives.DataGridColumnHeader) Then
                Dim header As Primitives.DataGridColumnHeader = dep
                Dim inTextBlock As TextBlock = header.Content.Children.Item(0)

                headerSelected = inTextBlock.Text

            ElseIf (TypeOf dep Is DataGridColumn) Then
                Dim header As DataGridColumn = dep
                Dim inTextBlock As TextBlock = header.Header.Children.Item(0).Text
                headerSelected = inTextBlock.Text

                'If user clicks anywhere in the rows, store the row and column index
            ElseIf (TypeOf dep Is DataGridCell) Then
                Dim cell As DataGridCell = dep
                While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridRow))
                    dep = VisualTreeHelper.GetParent(dep)

                End While
                Dim row As DataGridRow = dep

                rowIndex = FindRowIndex(row)
                columnIndex = cell.Column.DisplayIndex

            End If

            'By default e.Handled is True. It signifies that right click has been handled, and sometimes context menu don't pop up
            'To avoid this, set it to False, so context menu gets visible
            e.Handled = False

            Return

        End Sub

        'This method determines the where user performed mouse right click and stores the location of click.
        'More specifically, it is used to determine on which row, the user has clicked the Add Row Button
        Private Sub HighlightClear()
            If blueHighlight = True Then
                If blueFlag = False Then
                    For Each cell In blueCellsColored
                        cell.BorderBrush = New SolidColorBrush(Colors.Black)
                        cell.BorderThickness = New Thickness(0.0)
                    Next
                    blueCellsColored.Clear()
                    blueFlag = True
                End If
                blueHighlight = False
            End If

        End Sub

        Private Sub detectCellClicked(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseLeftButtonUp

            HighlightClear()

            Dim dep As DependencyObject = e.OriginalSource

            While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridCell) AndAlso Not (TypeOf dep Is Primitives.DataGridColumnHeader))
                dep = VisualTreeHelper.GetParent(dep)
            End While

            If dep Is Nothing Then
                Return


            ElseIf (TypeOf dep Is DataGridCell) Then
                Dim cell As DataGridCell = dep
                While (Not (dep Is Nothing) AndAlso Not (TypeOf dep Is DataGridRow))
                    dep = VisualTreeHelper.GetParent(dep)
                End While
                Dim row As DataGridRow = dep

                rowEditIndex = FindRowIndex(row)
                colEditIndex = cell.Column.DisplayIndex
            End If

            'By default e.Handled is True. It signifies that right click has been handled, and sometimes context menu don't pop up
            'To avoid this, set it to False, so context menu gets visible
            e.Handled = False

            Return

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
                filterStatus.Visibility = Visibility.Visible
            End If
        End Sub

        ''*********Miscellaneous Tasks. Copy Cut Paste

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

        ''*********End of Miscellaneous tasks



        ''********Row Header Context Menu Tasks
        'Performs the same Function as Add Row Button. i.e. Adding Row
        'This is based on context menu click of Row Header
        Private Sub addRowHeader_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = 16 Then
                collection.Insert(rowEditIndex + 1L, New userData())
            End If
        End Sub


        'Performs the same Function as Delete Row Button. i.e. Deleting Row
        'This is based on context menu click of Row Header
        Private Sub deleteRowHeader_Click(sender As Object, e As RoutedEventArgs)
            If columnIndex = 17 Then
                collection.RemoveAt(rowIndex)
                If collection.Count = 0 Then
                    collection.Add(New userData())
                End If
            End If

        End Sub
        ''********End of Row Header Context Menu Tasks


        ''********Row Context Menu Tasks
        Private Sub copyCells_Click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub pasteCells_Click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Sub highlightCells_Click(sender As Object, e As RoutedEventArgs)
            blueFlag = False
            blueHighlight = False
            ''Highlight option is available for every row cell
            ''This is to ensure that it works only when it is clicked on Selection Column
            Dim foundSelection As Boolean = False
            If columnIndex = 1 Then
                '' First Read the Configuration file, to get the current Selection configuration
                '' In such Way, even if the config file is changed in middle of application, it won't be affected
                fileRead()
                Dim obj As userData = collection.Item(rowIndex)
                For Each List In configHeaderList

                    If obj.selection.Equals(List.Item(0)) Then
                        foundSelection = True
                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(1))
                        changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                        For i As Integer = 1 To List.Count - 1

                            If List.Item(i).Equals("Attribute1") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(2))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute2") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(3))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute3") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(4))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute4") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(5))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Unit_Attri4") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(6))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute5") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(7))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute6") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(8))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute7") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(9))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute8") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(10))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute9") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(11))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)

                            ElseIf List.Item(i).Equals("Attribute10") Then
                                dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(12))
                                changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                            End If
                        Next
                        Exit For
                    End If
                Next

                '' If none of the selection value in Configuration File Matches, default selection is selected
                '' i.e only Selection cell and Name Cell
                If foundSelection = False Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(1))
                    changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(0))
                    changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                End If

                blueHighlight = True

            End If
        End Sub
        ''********End Of Row Context Menu tasks



        Private Sub CompleteFilter_Changed(sender As Object, e As RoutedEventArgs)
            If cbCompleteFilter.IsChecked = False Then
                CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                filterStatus.Content = "Currently No Filters Have been Applied"
            Else
                CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                filterStatus.Content = "Currently Filter is applied to Column : " & headerSelected & " with Value : " & filterValue
            End If
            filterStatus.Visibility = Visibility.Visible
        End Sub


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
                    Dim headersList(colCount) As String

                    For counter As Integer = 0 To colCount - 3
                        Dim col = dg_grid1.Columns.Item(counter)
                        headersList(counter) = col.Header.Children.Item(0).Text
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
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Warning", MessageBoxButton.OK)
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
                End If

            End If
        End Sub

        Private Sub displayExcelFile(filename As String, sheetName As String)

            Dim excelApp As Excel.Application = New Excel.Application()
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(filename)
            Dim worksheet As Excel.Worksheet = workbook.Sheets(sheetName)

            Dim col As Integer = 0, row As Integer = 0
            Dim range As Excel.Range = worksheet.UsedRange
            Dim warning As String = "This will close the existing opened spreadhsheet" & vbNewLine & " Do you want to save it before you import ?"


            If (MessageBox.Show(warning, "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) = MessageBoxResult.Yes) Then
                btn_export_Click(New Object(), New Object())
            Else
                ' Load all cells into 2d array.
                Dim array(,) As Object = range.Value(XlRangeValueDataType.xlRangeValueDefault)

                ' Scan the cells.
                If array IsNot Nothing Then
                    Console.WriteLine("Length: {0}", array.Length)

                    ' Get bounds of the array.
                    Dim bound0 As Integer = array.GetUpperBound(0)
                    Dim bound1 As Integer = array.GetUpperBound(1)

                    Console.WriteLine("Dimension 0: {0}", bound0)
                    Console.WriteLine("Dimension 1: {0}", bound1)

                    ' Loop over all elements.
                    For rowCount As Integer = 1 To bound0
                        For colCount As Integer = 1 To bound1
                            'array(rowCount, colCount)
                        Next
                        Console.WriteLine()
                    Next
                End If

            End If

            workbook.Save()
            workbook.Close()
            excelApp.Quit()

        End Sub

        Private Function arraytoUserDataObject(list As Object(,)) As userData()
            Dim temp(dg_grid1.Columns.Count) As Object

            For i As Integer = 0 To temp.Length
                temp(i) = vbNull
            Next
            For i As Integer = 1 To list.GetUpperBound(0)

                temp = arraySlice(list, i, list.GetUpperBound(1))


            Next

        End Function

        Private Function arraySlice(list As Object(,), currRow As Integer, nCols As Integer) As Object()
            Dim temp() As Object = New Object()
            temp(0) = Nothing
            For j As Integer = 1 To nCols
                temp(j) = list(currRow, j)
            Next
            Return temp
        End Function

        Private Sub btn_filter_Click(sender As Object, e As RoutedEventArgs)
            If TypeOf sender Is System.Windows.Controls.Button Then
                Dim btn As System.Windows.Controls.Button = sender
                Dim colCount As Integer = dg_grid1.Columns.Count
                Dim headersList(colCount) As String
                For counter As Integer = 0 To colCount - 3
                    Dim col = dg_grid1.Columns.Item(counter)
                    headersList(counter) = col.Header.Children.Item(0).Text
                Next

                For counter As Integer = 0 To colCount - 3
                    If btn.Name.Equals(btnNamesArray(counter)) Then
                        headerSelected = headersList(counter)
                    End If
                Next
            End If

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
                filterStatus.Visibility = Visibility.Visible
            End If

        End Sub


        Private Sub btn_save_click(sender As Object, e As RoutedEventArgs)

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


        Private Sub btn_validate_Click(sender As Object, e As RoutedEventArgs)
            validate_Mandatory(collection.Count)
            validate_Integers(collection.Count)
            filterStatus.Visibility = Visibility.Hidden
            errorStatus.Content = ""
            If errorHighlight = True Then
                errorStatus.Content = errorStatus.Content & vbNewLine & "Cells Highlighted in Red Must not be left Blank"
            End If

            If valHighlight = True Then
                errorStatus.Content = errorStatus.Content & vbNewLine & "Value highlighted in Dark Red don't match with validation conditions"
            End If


        End Sub

        Private Sub validate_Mandatory(nRows As Integer)
            fileRead()

            For Each cell In redcellsColored
                cell.BorderBrush = New SolidColorBrush(Colors.Black)
                cell.BorderThickness = New Thickness(1.0)
            Next
            redcellsColored.Clear()
            errorHighlight = False

            For counter As Integer = 0 To nRows - 1
                Dim temp_userData As userData = collection.Item(counter)
                For Each List In configHeaderList

                    If temp_userData.selection.Equals(List.Item(0)) Then
                        For i As Integer = 1 To List.Count - 1

                            If List.Item(i).Equals("Attribute1") Then
                                If temp_userData.attribute1.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(2))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            ElseIf List.Item(i).Equals("Attribute2") Then
                                If temp_userData.attribute2.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(3))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            ElseIf List.Item(i).Equals("Attribute3") Then
                                If temp_userData.attribute3.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(4))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            ElseIf List.Item(i).Equals("Attribute4") Then
                                If temp_userData.attribute4.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(5))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            ElseIf List.Item(i).Equals("Unit_Attri4") Then
                                If temp_userData.unitattri4.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(6))
                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            ElseIf List.Item(i).Equals("Attribute5") Then
                                If temp_userData.attribute5.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(7))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)

                                End If

                            ElseIf List.Item(i).Equals("Attribute6") Then
                                If temp_userData.attribute6.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(8))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If

                            ElseIf List.Item(i).Equals("Attribute7") Then
                                If temp_userData.attribute7.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(9))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If

                            ElseIf List.Item(i).Equals("Attribute8") Then
                                If temp_userData.attribute8.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(10))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If

                            ElseIf List.Item(i).Equals("Attribute9") Then
                                If temp_userData.attribute9.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(11))

                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If

                            ElseIf List.Item(i).Equals("Attribute10") Then
                                If temp_userData.attribute10.Equals("") Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(12))
                                    changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                End If
                            End If
                        Next
                        Exit For
                    End If
                Next
            Next
            dg_grid1.SelectedCells.Clear()
        End Sub

        Private Sub validate_Integers(nRows As Integer)
            fileRead()
            For Each cell In greencellsColored
                cell.BorderBrush = New SolidColorBrush(Colors.Black)
                cell.BorderThickness = New Thickness(1.0)
            Next
            greencellsColored.Clear()

            For counter As Integer = 0 To nRows - 1
                Dim temp_userData As userData = collection.Item(counter)

                If (temp_userData.minVal > temp_userData.normVal) Or (temp_userData.normVal > temp_userData.maxVal) Or (temp_userData.minVal > temp_userData.maxVal) Then
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(13))
                    changeCellColor(dg_grid1.CurrentCell, Colors.DarkRed, Colors.White)
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(14))
                    changeCellColor(dg_grid1.CurrentCell, Colors.DarkRed, Colors.White)
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(15))
                    changeCellColor(dg_grid1.CurrentCell, Colors.DarkRed, Colors.White)
                End If
            Next
        End Sub

        Private Sub btn_close_Click(sender As Object, e As RoutedEventArgs)
            'Dim colCount As Integer = dg_grid1.Columns.Count
            'Dim headersList(colCount) As String
            'For counter As Integer = 0 To colCount - 3
            '    Dim col = dg_grid1.Columns.Item(counter)
            '    headersList(counter) = col.Header.Children.Item(0).Text
            '    Console.WriteLine(headersList(counter))
            'Next
        End Sub

    End Class
End Namespace
