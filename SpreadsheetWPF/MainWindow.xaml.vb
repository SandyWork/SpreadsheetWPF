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

        Public Property col_list As List(Of String)

        Public Sub New(list As List(Of String))

            If list IsNot Nothing Then
                If list.Count > 0 Then
                    col_list = New List(Of String)()
                    For Each value In list
                        col_list.Add(value)
                    Next
                End If
            End If
        End Sub

        Public Sub New(list() As String, Optional columnCount As Integer = 0)

            If list Is Nothing Then

                If columnCount = 0 Then
                    Console.WriteLine("Invalid Parameters passed to userData constructor ")
                ElseIf columnCount > 0
                    col_list = New List(Of String)()
                    For i As Integer = 0 To columnCount - 1
                        col_list.Add("")
                    Next
                End If
            Else
                If list.Length < columnCount Then
                    columnCount = columnCount - list.Length

                    col_list = New List(Of String)()
                    For i As Integer = 0 To list.Length - 1
                        col_list.Add(list(i))
                    Next
                    For i As Integer = 0 To columnCount - 1
                        col_list.Add("")
                    Next
                Else
                    col_list = New List(Of String)()
                    For i As Integer = 0 To list.Length - 1
                        col_list.Add(list(i))
                    Next
                End If
            End If
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

        End Sub
        Protected Overrides Sub OnCollectionChanged(e As NotifyCollectionChangedEventArgs)
            MyBase.OnCollectionChanged(e)
        End Sub
    End Class

    Class MainWindow

        Dim collection As PresentData
        Dim headerList() As String = {"Item Tag Name", "Measuring Principle", "Measuring/Adjust Location", "PID Sheet Number", "Construction Status", "pressure P1 minimum", "pressure P1 in operation", "pressure P1 maximum", "unit of pressure P1", "temperature minimum", "temperature in operation", "temperature maximum", "unit of temperature", "differential pressure minimum", "differential pressure in operation", "differential pressure maximum", "unit of differential pressure"}
        Dim btnNamesArray() As String = {"btn_filter_name", "btn_filter_sel", "btn_filter_attri1", "btn_filter_attri2", "btn_filter_attri3", "btn_filter_attri4", "btn_filter_unitattri4", "btn_filter_attri5", "btn_filter_attri6", "btn_filter_attri7", "btn_filter_attri8", "btn_filter_attri9", "btn_filter_attri10", "btn_filter_minval", "btn_filter_normval", "btn_filter_maxval", "btn_filter_unitofdifferentialpressue"}
        Dim colNames() As String = {"dgtxtcol_name", "dgtxtcol_sel", "dgtxtcol_attri1", "dgtxtcol_attri2", "dgtxtcol_attri3", "dgtxtcol_attri4", "dgtxtcol_attri5", "dgtxtcol_attri6", "dgtxtcol_attri7", "dgtxtcol_attri8", "dgtxtcol_attri9", "dgtxtcol_attri10", "dgtxtcol_attri11", "dgtxtcol_minval", "dgtxtcol_normval", "dgtxtcol_maxval", "dgtxtcol_unitofdifferentalpressure"}

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

        Dim previousSelectedCells As List(Of DataGridCellInfo) = New List(Of DataGridCellInfo)()
        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()
            collection = Me.Resources("presentData")
            collection.Clear()
            AddColumns()
            defaultData_dgGrid()
        End Sub

        Private Sub defaultData_dgGrid()
            'For i As Integer = 0 To 200 Step 100
            '    Dim templist As List(Of String) = New List(Of String)()
            '    For j As Integer = 0 To dg_grid1.Columns.Count - 3
            '        templist.Add(i + j)
            '    Next
            '    Dim obj2 As userData = New userData(templist)
            '    collection.Add(obj2)
            'Next

            Dim arrayData(dg_grid1.Columns.Count - 2) As String
            Dim obj As userData

            For i As Integer = 0 To arrayData.Length - 1
                arrayData(i) = ""
            Next

            ''Insert the data you need to insert below. 
            ''If you Then need multiple rows To be inserted,clear the array, insert new values,and add to collection
            ''Given below is an example

            'Adding Values to array, you can skip to enter rightmost values. 
            'If you are skipping values, please do pass the number of columns as an arguemnt
            'dg_grid1.Columns.Count will give you the number of columns ( including add row and delete row columns )
            'So actual number of columns will be dg_grid1.Columns.Count - 2  
            'But If you want an empty value to be present in between, you must enter it Like below
            arrayData = {"F11001", "MFM", "M1", "0002", "New", "", "6", "12", "bara", "50", "120", "200", "°C", "", "0.4", "0.8", "bar"}

            'Adding object to collection
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            'Method to clear Array
            Array.Clear(arrayData, 0, arrayData.Length)

            arrayData = {"H11001", "BAV", "S1", "0002", "New", "", "", "12", "bara", "50", "120", "200", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            'Method to clear Array
            Array.Clear(arrayData, 0, arrayData.Length)

            'In Below Example, I have included all the values to be entered for all the columns.
            'So no need to pass the columns count in this case. Just pass the array in such cases
            arrayData = {"H16601", "BUC", "S1", "0001", "New", "0.8", "1.2", "1.5", "bar(pe)", "20", "25", "50", "°C", "100", "150", "300", "mbar"}
            obj = New userData(arrayData)
            collection.Add(obj)

            arrayData = {"H16632", "BUV", "S1", "0001", "New", "", "", "4", "bar(pe)", "12", "28", "50", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            arrayData = {"L11001", "DPT", "M1", "0002", "New", "1", "1.3", "2.4", "bar(abs)", "", "94.8", "200", "°C", "", "1.3", "1.5", "bar"}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            arrayData = {"L11003", "LST", "M1", "0002", "New", "", "", "12", "bara", "", "120", "200", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            arrayData = {"L16608", "LCT", "M1", "0001", "New", "", "", "10.5", "bar(abs)", "", "", "45", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 2)
            collection.Add(obj)

            Console.WriteLine("Columns Count" & dg_grid1.Columns.Count)
            Console.WriteLine("Attributes Count" & obj.col_list.Count)
        End Sub

        Public Sub New(list As userData)

            ' This call is required by the designer.
            InitializeComponent()
            collection = Me.Resources("presentData")
            collection.Clear()
            AddColumns()

            If list.col_list.Count > 0 Then
                collection.Add(list)
            Else
                defaultData_dgGrid()
            End If
        End Sub

        Public Sub New(list() As userData)

            ' This call is required by the designer.
            InitializeComponent()
            collection = Me.Resources("presentData")
            collection.Clear()
            AddColumns()

            If list.Count > 0 Then
                For Each item In list
                    collection.Add(item)
                Next
            Else
                defaultData_dgGrid()
            End If
        End Sub

        Public Sub New(arr() As String, colCount As Integer)

            ' This call is required by the designer.
            InitializeComponent()
            collection = Me.Resources("presentData")
            collection.Clear()
            AddColumns()
            ' Add any initialization after the InitializeComponent() call.


            If arr.Length > 0 Then
                Dim obj As userData = New userData(arr)

            End If

        End Sub

        Private Sub AddColumns()

            For i As Integer = 0 To dg_grid1.Columns.Count - 3
                Dim col = dg_grid1.Columns.Item(i)
                col.Header.Children.Item(0).Text = headerList(i)


            Next

        End Sub
        'This method is used to automatically resize the dock panel when user resizes Window
        Private Sub resizeWindow(sender As Object, e As SizeChangedEventArgs)
            pnl_dock.Width = win_main.ActualWidth
            pnl_dock.Height = win_main.ActualHeight
        End Sub


        ''**********All the Helper Functions

        Private Function FindRowIndexFromRow(row As DataGridRow) As Integer
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

        Private Sub fileRead()
            Try
                'Check if an existing configuration file exists.
                'If not create a new one with default Values
                If File.Exists(configurationFileName) = False Then

                    Console.WriteLine("No Configuration File")
                    Console.WriteLine("Creating One With Default Values")


                    Using sw As StreamWriter = File.CreateText(configurationFileName)
                        sw.WriteLine("<-- Format to specify columns to highlight for specific Value Is :     -->")
                        sw.WriteLine("<-- *[Value] -->" & vbNewLine & "Column Names in separate line" & "-->")
                        sw.WriteLine("**@ 1" & vbNewLine & "Measuring Principle" & vbNewLine & "Construction Status" & vbNewLine & "PID Sheet Number")
                        sw.WriteLine("**@ 2" & vbNewLine & "pressure P1 minimum" & vbNewLine & "pressure P1 maximum" & vbNewLine & "Measuring Principle")
                        sw.WriteLine("**@ 3" & vbNewLine & "pressure P1 maximum" & vbNewLine & "pressure P1 in operation" & vbNewLine & "Measuring Principle")
                        sw.Flush()
                    End Using
                End If
                ' Open the file to read from.
                Using sr As StreamReader = File.OpenText(configurationFileName)
                    configHeaderList.Clear()
                    Dim counter As Integer = -1
                    Dim temp As String = ""
                    While sr.Peek() >= 0
                        temp = sr.ReadLine()
                        temp = temp.Trim()

                        If temp.Contains("**@") Then
                            configHeaderList.Add(New List(Of String)())
                            temp = temp.Substring(4)
                            counter = counter + 1
                            configHeaderList(counter).Add(temp)
                            Continue While
                        Else
                            If Not temp.Equals("") Then
                                If counter >= 0 Then
                                    configHeaderList(counter).Add(temp)
                                End If
                            Else
                                Continue While
                            End If
                        End If
                    End While
                End Using

                For i As Integer = 0 To configHeaderList.Count - 1
                    For j As Integer = 0 To configHeaderList(i).Count - 1
                        Console.WriteLine(configHeaderList(i).Item(j))
                    Next

                Next

            Catch e As Exception
                ' Let the user know what went wrong.
                Console.WriteLine("The file could Not be read:")
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

        Public Shared Function GetRowIndexFromCell(dataGrid As DataGrid, dataGridCellInfo As DataGridCellInfo) As Integer
            Dim dgrow As DataGridRow = DirectCast(dataGrid.ItemContainerGenerator.ContainerFromItem(dataGridCellInfo.Item), DataGridRow)
            If dgrow IsNot Nothing Then
                Return dgrow.GetIndex()
            Else
                Return -1
            End If

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
        Private Sub collectionviewsource_filter(sender As Object, e As FilterEventArgs)
            If filterSelected Then
                Dim obj As userData = e.Item
                If obj IsNot Nothing Then
                    ' if filter is turned on, filter completed items.
                    If Me.cbCompleteFilter.IsChecked = True Then
                        Dim headerIndex As Integer = determineIndex(headerSelected)
                        If obj.col_list.Item(headerIndex).Equals(filterValue) Then
                            e.Accepted = True
                        Else
                            e.Accepted = False
                        End If
                    End If
                End If
            End If
        End Sub

        Private Function determineIndex(header As String) As Integer

            For i As Integer = 0 To dg_grid1.Columns.Count
                If headerList(i).Equals(header) Then
                    Return i
                End If
            Next
            Return 0
        End Function
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

                rowIndex = FindRowIndexFromRow(row)
                columnIndex = cell.Column.DisplayIndex

            End If

            'By default e.Handled is True. It signifies that right click has been handled, and sometimes context menu don't pop up
            'To avoid this, set it to False, so context menu gets visible
            e.Handled = False

            Return

        End Sub

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

        'This method determines the where user performed mouse right click and stores the location of click.
        'More specifically, it is used to determine on which row, the user has clicked the Add Row Button
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

                rowEditIndex = FindRowIndexFromRow(row)
                colEditIndex = cell.Column.DisplayIndex
            End If

            'By default e.Handled is True. It signifies that right click has been handled, and sometimes context menu don't pop up
            'To avoid this, set it to False, so context menu gets visible
            e.Handled = False

            Return

        End Sub

        Private Sub btn_edit_row_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = dg_grid1.Columns.Count - 2 Then
                collection.Insert(rowEditIndex + 1L, New userData(Nothing, dg_grid1.Columns.Count))
            End If
        End Sub

        Private Sub btn_delete_row_Click(sender As Object, e As RoutedEventArgs)
            If colEditIndex = dg_grid1.Columns.Count - 1 Then
                collection.RemoveAt(rowEditIndex)
                If collection.Count = 0 Then
                    collection.Add(New userData(Nothing, dg_grid1.Columns.Count))
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
                filterStatus.Content = "Currently Filter Is applied to Column  " & headerSelected & " with Value : " & filterValue
                filterStatus.Visibility = Visibility.Visible
            End If
        End Sub

        ''*********Miscellaneous Tasks. Copy Cut Paste

        Private Sub CutCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)
            If copyActivated Then
                copyActivated = False
            End If
            HoldCells(sender, e)
            cutActivated = True
        End Sub

        Private Sub HoldCells(sender As Object, e As ExecutedRoutedEventArgs)
            Dim cellsSelected As IList(Of DataGridCellInfo) = dg_grid1.SelectedCells

            If cellsSelected.Count > 0 Then
                previousSelectedCells.Clear()
                For Each cell In cellsSelected
                    previousSelectedCells.Add(cell)
                Next
            End If
        End Sub

        Private Sub CopyCommand(sender As Object, e As ExecutedRoutedEventArgs)
            If cutActivated Then
                cutActivated = False
            End If
            copyActivated = True
            HoldCells(sender, e)
        End Sub

        Private Sub RetrieveCells(sender As Object, e As ExecutedRoutedEventArgs)
            Dim ifCut As Boolean = False
            Dim currentCell As DataGridCellInfo = dg_grid1.SelectedCells.Item(0)
            Dim currColIndex = currentCell.Column.DisplayIndex
            Dim currRowIndex = GetRowIndexFromCell(dg_grid1, currentCell)
            Dim curruserData As userData = collection.Item(currRowIndex)

            If previousSelectedCells.Count > 0 Then
                Dim rowIndex = GetRowIndexFromCell(dg_grid1, previousSelectedCells.Item(0))
                Dim prevuserData As userData = collection.Item(rowIndex)
                Dim colIndex As Integer = 0
                For i As Integer = 0 To previousSelectedCells.Count - 1
                    Dim cell As DataGridCellInfo = previousSelectedCells.Item(i)
                    colIndex = cell.Column.DisplayIndex
                    curruserData.col_list.Item(currColIndex) = prevuserData.col_list.Item(colIndex)

                    If cutActivated Then
                        prevuserData.col_list.Item(colIndex) = ""
                        ifCut = True
                    End If
                    currColIndex += 1
                Next

                If ifCut Then
                    cutActivated = False
                Else
                    copyActivated = False
                End If

                If rowIndex > currRowIndex Then
                    collection.RemoveAt(rowIndex)
                    collection.RemoveAt(currRowIndex)

                    collection.Insert(currRowIndex, curruserData)
                    collection.Insert(rowIndex, prevuserData)
                ElseIf rowIndex < currRowIndex Then
                    collection.RemoveAt(currRowIndex)
                    collection.RemoveAt(rowIndex)

                    collection.Insert(rowIndex, prevuserData)
                    collection.Insert(currRowIndex, curruserData)
                Else
                    collection.RemoveAt(rowIndex)
                    collection.Insert(currRowIndex, curruserData)
                End If

            End If
        End Sub
        Private Sub DeleteCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)
            deleteItem(sender, e)
        End Sub

        Private Sub deleteItem(sender As Object, e As ExecutedRoutedEventArgs)

            Dim cellsSelected As IList(Of DataGridCellInfo) = dg_grid1.SelectedCells
            Dim currentCell As DataGridCellInfo = dg_grid1.SelectedCells.Item(0)
            Dim currColIndex = currentCell.Column.DisplayIndex
            Dim currRowIndex = GetRowIndexFromCell(dg_grid1, currentCell)
            Dim curruserData As userData = collection.Item(currRowIndex)
            Dim cell As DataGridCellInfo = Nothing
            If cellsSelected.Count > 0 Then
                For i As Integer = 0 To cellsSelected.Count - 1
                    cell = cellsSelected.Item(i)
                    currColIndex = cell.Column.DisplayIndex
                    curruserData.col_list.Item(currColIndex) = ""

                Next
            End If
            collection.RemoveAt(currRowIndex)
            collection.Insert(currRowIndex, curruserData)
            dg_grid1.SelectedCells.Add(cell)
            dg_grid1.CurrentCell = cell
        End Sub

        Private Sub PasteCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)

            RetrieveCells(sender, e)
            pasteActivated = True
            previousSelectedCells.Clear()
        End Sub

        ''*********End of Miscellaneous tasks

        Private Sub copyCells_Click(sender As Object, e As RoutedEventArgs)
            CopyCommand(sender, Nothing)
        End Sub


        Private Sub cutCells_Click(sender As Object, e As RoutedEventArgs)
            CutCommand_Executed(sender, Nothing)
        End Sub


        Private Sub pasteCells_Click(sender As Object, e As RoutedEventArgs)
            RetrieveCells(sender, Nothing)
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
                    If List.Item(0) IsNot Nothing Then
                        If obj.col_list.Item(1).Equals(List.Item(0)) Then
                            foundSelection = True
                            dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(1))
                            changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                            For i As Integer = 1 To List.Count - 1
                                For j As Integer = 2 To dg_grid1.Columns.Count - 3
                                    If List.Item(i).Equals(headerList(j)) Then
                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(rowIndex), dg_grid1.Columns.Item(j))
                                        changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                                    End If
                                Next
                            Next
                            Exit For
                        End If
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
                filterStatus.Content = "Currently Filter Is applied to Column  " & headerSelected & " with Value : " & filterValue
            End If
            filterStatus.Visibility = Visibility.Visible
        End Sub


        Private Sub btn_export_Click(sender As Object, e As RoutedEventArgs)
            exportExcel()
        End Sub

        Private Sub exportExcel()
            '' Main Content
            Dim f As SaveFileDialog = New SaveFileDialog()
            f.Filter = "Excel Workbook (*.xlsx) |*.xlsx|All files (*.*)|*.*"
            Try
                If f.ShowDialog() = True Then
                    Dim xlApp As Excel.Application = New Excel.Application()
                    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add
                    Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(1)
                    Dim colCount = dg_grid1.Columns.Count - 2, rowCount = collection.Count

                    'Create an array with 16 columns and n rows
                    Dim DataArray(rowCount, colCount) As Object

                    For row As Short = 0 To rowCount - 1
                        For col As Short = 0 To colCount - 3
                            Dim index = determineIndex(headerList(col))
                            DataArray(row, col) = collection.Item(row).col_list.Item(index)
                        Next
                    Next
                    xlWorkSheet.Range("A1").Resize(1, colCount).Value = headerList
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

        Private Sub btn_import_click(sender As Object, e As RoutedEventArgs)
            Dim openfiledialog As OpenFileDialog = New OpenFileDialog()
            openfiledialog.Filter = "excel 97-2003 worksheet (*.xls)|*.xls|excel workbook (*.xlsx) |*.xlsx|all files (*.*)|*.*"
            openfiledialog.Multiselect = False

            If (openfiledialog.ShowDialog() = True) Then
                If openfiledialog.CheckFileExists = True Then
                    Dim filename = openfiledialog.FileName


                    Dim sheetnames As List(Of String) = getexcelsheetnames(filename)
                    Dim excel As ImportExcel = New ImportExcel(sheetnames)
                    excel.ShowInTaskbar = True
                    excel.Owner = Me

                    If excel.ShowDialog = True Then
                        Dim sheetname As String = excel.getSheetName()
                        displayexcelfile(filename, sheetname)
                    End If
                End If

            End If
        End Sub



        Private Sub displayexcelfile(filename As String, sheetname As String)

            Dim excelapp As Excel.Application = New Excel.Application()
            Dim workbook As Excel.Workbook = excelapp.Workbooks.Open(filename)
            Dim worksheet As Excel.Worksheet = workbook.Sheets(sheetname)

            Dim col As Integer = 0, row As Integer = 0
            Dim range As Excel.Range = worksheet.UsedRange
            Dim warning As String = "This will close the existing opened sheet" & vbNewLine & " Do you want to save the current file before you import ?"


            If (MessageBox.Show(warning, "warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) = MessageBoxResult.Yes) Then
                exportExcel()
            End If
            ' load all cells into 2d array.
            Dim array(,) As Object = range.Value(XlRangeValueDataType.xlRangeValueDefault)

            ' scan the cells.
            If array IsNot Nothing Then
                Console.WriteLine("length {0}", array.Length)

                ' get bounds of the array.
                Dim bound0 As Integer = array.GetUpperBound(0)
                Dim bound1 As Integer = array.GetUpperBound(1)

                Console.WriteLine("dimension 0 {0}", bound0)
                Console.WriteLine("dimension 1 {0}", bound1)


                collection.Clear()

                ' loop over all elements.
                For rowcount As Integer = 1 To bound0
                    Dim temp As List(Of String) = New List(Of String)()

                    For colcount As Integer = 1 To bound1
                        temp.Add(array(rowcount, colcount))
                    Next
                    Dim obj As userData = New userData(temp)
                    collection.Add(obj)
                Next
            End If

            workbook.Save()
            workbook.Close()
            excelapp.Quit()

        End Sub

        Private Sub btn_filter_Click(sender As Object, e As RoutedEventArgs)
            If TypeOf sender Is System.Windows.Controls.Button Then
                Dim btn As System.Windows.Controls.Button = sender

                For counter As Integer = 0 To dg_grid1.Columns.Count - 3
                    If btn.Name.Equals(btnNamesArray(counter)) Then
                        headerSelected = headerList(counter)
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
                filterStatus.Content = "Currently Filter Is applied to Column  " & headerSelected & " with Value : " & filterValue
                filterStatus.Visibility = Visibility.Visible
            End If

        End Sub


        Private Sub btn_save_click(sender As Object, e As RoutedEventArgs)

        End Sub

        Private Function getexcelsheetnames(ByVal filename As String) As List(Of String)
            Dim strconn As String = "provider=microsoft.ace.oledb.12.0;data source=" &
          filename & ";extended properties=""excel 12.0 xml;hdr=yes"";"
            Dim conn As New OleDb.OleDbConnection(strconn)

            conn.Open()

            Dim dtsheets As System.Data.DataTable =
              conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim listsheet As New List(Of String)
            Dim drsheet As DataRow

            For Each drsheet In dtsheets.Rows
                Dim temp As String = drsheet("table_name").ToString()
                listsheet.Add(temp.Substring(0, temp.Length - 1))
            Next
            conn.Close()
            Return listsheet

        End Function


        Private Sub btn_validate_Click(sender As Object, e As RoutedEventArgs)
            validate_Mandatory(collection.Count)
            validate_Integers(collection.Count)
            filterStatus.Visibility = Visibility.Hidden
            errorStatus.Content = ""
            If errorHighlight = True Then
                errorStatus.Content = errorStatus.Content & vbNewLine & "Cells Highlighted in Red Must Not be left Blank"
            End If

            If valHighlight = True Then
                errorStatus.Content = errorStatus.Content & vbNewLine & "Value highlighted in Dark Red don't match with validation conditions"
            End If


        End Sub

        Private Sub validate_Mandatory(nRows As Integer)
            fileRead()

            For Each cell In redcellsColored
                cell.BorderBrush = New SolidColorBrush(Colors.Black)
                cell.BorderThickness = New Thickness(0.0)
            Next
            redcellsColored.Clear()
            errorHighlight = False

            For counter As Integer = 0 To nRows - 1
                Dim temp_userdata As userData = collection.Item(counter)
                For Each list In configHeaderList
                    If temp_userdata.col_list.Item(1).Equals(list.Item(0)) Then
                        Console.WriteLine("Selection Equals " & list.Item(0))
                        For i As Integer = 1 To list.Count - 1
                            For j As Integer = 2 To dg_grid1.Columns.Count - 3
                                If list.Item(i).Equals(headerList(j)) Then
                                    If temp_userdata.col_list.Item(j).Equals("") Then
                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(j))
                                        changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                    End If
                                End If
                            Next
                        Next
                        Exit For
                    End If
                Next
            Next
        End Sub

        Private Sub validate_Integers(nRows As Integer)
            fileRead()
            For Each cell In greencellsColored
                cell.BorderBrush = New SolidColorBrush(Colors.Black)
                cell.BorderThickness = New Thickness(0.0)
            Next
            greencellsColored.Clear()
            Dim minVal As Integer, normVal As Integer, maxVal As Integer
            Dim minIndex As Integer, normIndex As Integer, maxIndex As Integer
            For i As Integer = 0 To dg_grid1.Columns.Count - 1
                If colNames(i).Equals("dgtxtcol_minval") Then
                    minIndex = i
                    If colNames(i + 1).Equals("dgtxtcol_normval") Then
                        normIndex = i + 1
                        If colNames(i + 2).Equals("dgtxtcol_maxval") Then
                            maxIndex = i + 2
                            Exit For
                        End If
                    End If
                End If
            Next
            For counter As Integer = 0 To nRows - 1
                Dim temp_userdata As userData = collection.Item(counter)


                minVal = CInt(temp_userdata.col_list.Item(minIndex))
                normVal = CInt(temp_userdata.col_list.Item(normIndex))
                maxVal = CInt(temp_userdata.col_list.Item(maxIndex))
                If (minVal > normVal) Or (normVal > maxVal) Or (minVal > maxVal) Then
                    For i As Integer = 0 To 2
                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(minIndex + i))
                        changeCellColor(dg_grid1.CurrentCell, Colors.DarkRed, Colors.White)
                    Next
                End If
            Next
        End Sub

        Private Sub btn_close_Click(sender As Object, e As RoutedEventArgs)
            fileRead()
        End Sub

    End Class
End Namespace
