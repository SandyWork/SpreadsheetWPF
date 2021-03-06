﻿Imports System.Collections.ObjectModel
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

        Dim _tempUnit_() As String = {"", "°C"}
        Dim _pressureUnit_() As String = {"", "bar", "mbar", "pascal"}
        Dim _diffPressureUnit_() As String = {"", "bar", "mbar", "pascal"}
        Dim _measuringprinciple_() As String = {"", "BAV", "BUC", "BUV", "DPT", "LCT", "LST", "MFM", "PG", "PGS", "RTD", "TE", "THE"}


        Public Property col_list As List(Of String)
        Public Property pressureUnit As List(Of String)
        Public Property diffPressureUnit As List(Of String)
        Public Property tempUnit As List(Of String)
        Public Property measuringprinciple As List(Of String)
        Public Property include As Boolean = True


        Private Sub addComboData(obj As userData)
            pressureUnit = New List(Of String)()
            diffPressureUnit = New List(Of String)()
            tempUnit = New List(Of String)()
            measuringprinciple = New List(Of String)


            For Each item In _tempUnit_
                obj.tempUnit.Add(item)
            Next

            For Each item In _pressureUnit_
                obj.pressureUnit.Add(item)
            Next

            For Each item In _diffPressureUnit_
                obj.diffPressureUnit.Add(item)
            Next

            For Each item In _measuringprinciple_
                obj.measuringprinciple.Add(item)
            Next

        End Sub

        Public Sub New(list As List(Of String))

            If list IsNot Nothing Then
                If list.Count > 0 Then
                    col_list = New List(Of String)()
                    For Each value In list
                        col_list.Add(value)
                    Next
                End If
            End If
            addComboData(Me)
        End Sub

        Public Sub New(list() As String, Optional columnCount As Integer = 0)

            Try
                col_list = New List(Of String)()
                If list Is Nothing Then
                    If columnCount = 0 Then
                        Console.WriteLine("Error : Invalid Parameters passed to userData String array constructor")
                        MsgBox("Encountered Error! Exiting Application. Check Output Window for more details")
                        Application.Current.Shutdown(1)
                    ElseIf columnCount > 0 Then
                        For i As Integer = 0 To columnCount - 1
                            col_list.Add("")
                        Next
                    End If
                ElseIf list.Length > 1 Then
                    If columnCount = 0 Or list.Length = columnCount Then
                        For i As Integer = 0 To list.Length - 1
                            col_list.Add(list(i))
                        Next
                    Else
                        If list.Length < columnCount Then
                            columnCount = columnCount - list.Length
                            For i As Integer = 0 To list.Length - 1
                                col_list.Add(list(i))
                            Next

                            'Fill empty Values
                            For i As Integer = 0 To columnCount - 1
                                col_list.Add("")
                            Next
                        ElseIf list.Length > columnCount Then
                            For i As Integer = 0 To columnCount - 1
                                col_list.Add(list(i))
                            Next
                        End If
                    End If
                End If

                addComboData(Me)
            Catch ex As Exception
                MsgBox("Encountered Error in Adding Values to userData object. Exiting Application.")
                Application.Current.Shutdown(1)
            End Try

        End Sub

        Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(PropertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
        End Sub

        Public Sub BeginEdit() Implements IEditableObject.BeginEdit
            'Console.WriteLine("Editing")
        End Sub

        Public Sub EndEdit() Implements IEditableObject.EndEdit
            'Console.WriteLine("End")
        End Sub

        Public Sub CancelEdit() Implements IEditableObject.CancelEdit
            'Console.WriteLine("Cancel Editing")
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

    Public Module MyExtensions

        <System.Runtime.CompilerServices.Extension()>
        Public Function convertDouble(value As String) As Double
            If String.IsNullOrEmpty(value) Then
                Return 0.0
            Else
                If Double.TryParse(value, Nothing) Then
                    Return CDbl(value)
                End If
            End If
            Return 0.0
        End Function

    End Module

    Class MainWindow

        Dim collection As PresentData
        Dim headerList() As String = {"Include", "Item Tag Name", "Measuring Principle", "Measuring/Adjust Location", "PID Sheet Number", "Construction Status", "Pressure P1 Minimum", "Pressure P1 In Operation", "Pressure P1 Maximum", "Unit Of Pressure P1", "Temperature Minimum", "Temperature In Operation", "Temperature Maximum", "Unit Of Temperature", "Differential Pressure Minimum", "Differential Pressure In Operation", "Differential Pressure Maximum", "Unit of Differential Pressure"}

        Dim btnNamesArray() As String = {"", "btn_filter_name", "btn_filter_sel", "btn_filter_attri1", "btn_filter_attri2", "btn_filter_attri3", "btn_filter_attri4", "btn_filter_unitattri4", "btn_filter_attri5", "btn_filter_attri6", "btn_filter_attri7", "btn_filter_attri8", "btn_filter_attri9", "btn_filter_attri10", "btn_filter_minval", "btn_filter_normval", "btn_filter_maxval", "btn_filter_unitofdifferentialpressure"}
        Dim colNames() As String = {"dgtxtcol_include", "dgtxtcol_name", "dgtxtcol_sel", "dgtxtcol_attri1", "dgtxtcol_attri2", "dgtxtcol_attri3", "dgtxtcol_attri4", "dgtxtcol_unitattri4", "dgtxtcol_attri5", "dgtxtcol_attri6", "dgtxtcol_attri7", "dgtxtcol_attri8", "dgtxtcol_attri9", "dgtxtcol_attri10", "dgtxtcol_minval", "dgtxtcol_normval", "dgtxtcol_maxval", "dgtxtcol_unitofdifferentialpressure"}
        Dim discardList As List(Of Integer) = New List(Of Integer)()


        'header of the column where user right clicked
        Dim headerSelected As String = "", configurationFileName As String = "..\selectionConfigFile.txt"

        'rowIndex and columnIndex stores the location of the cell where user right clicked to open the context Menu
        Dim rowIndex As Short = 0, columnIndex As Short = 0, measuringIndex As Short = 0, itemNameIndex As Short = 0
        'rowEditIndex, colEditIndex stores the location of the cell where user left clicked
        'This is used for Add Row and Delete Row Buttons to know exactly where users wants to perform Row operations

        Dim rowEditIndex As Integer, colEditIndex As Integer = 0
        Dim oldWindowHeight As Double = 600
        'Filter Value that user entered when prompted
        Dim filterValue As String = ""
        Dim copyActivated As Boolean = False, cutActivated As Boolean = False, pasteActivated As Boolean = False, rowAdd = False, allHighlight = True
        Private filterSelected As Boolean = False, caseSensitive As Boolean = False
        Dim configHeaderList As List(Of List(Of String)) = New List(Of List(Of String))()
        Dim configIndexCount As Integer = 0
        Dim redcellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim darkRedCellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim violetcellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim blueCellsColored As List(Of DataGridCell) = New List(Of DataGridCell)
        Dim errorHighlight As Boolean = False, valHighlight As Boolean = False,
            blueHighlight As Boolean = False, blueFlag As Boolean = False, intHighlight As Boolean = False
        Dim previousSelectedCells As List(Of DataGridCellInfo) = New List(Of DataGridCellInfo)()

        Private Class XMLSpreadsheetCellData
            Public Property CellId As Int32
            Public Property RowIndex As Int32
            Public Property ColumnIndex As Int32
            Public Property DataType As String
            Public Property DataValue As String
        End Class

        Private Function ParseClipboard() As Object(,)
            Dim clipboardData = Clipboard.GetDataObject
            If clipboardData IsNot Nothing Then
                If clipboardData.GetFormats.Contains("XML Spreadsheet") Then
                    Dim spreadsheet = New DataSet
                    spreadsheet.ReadXml(clipboardData.GetData("XML Spreadsheet"))
                    Dim rowCount = spreadsheet.Tables("Table").Rows(0)("ExpandedRowCount")
                    Dim columnCount = spreadsheet.Tables("Table").Rows(0)("ExpandedColumnCount")
                    If rowCount > 0 AndAlso columnCount > 0 Then
                        Dim result(rowCount - 1, columnCount - 1) As Object
                        If spreadsheet.Tables.Contains("Data") Then 'if there is no "Data" table then all cells are empty and all array elements will be nothing
                            If Not spreadsheet.Tables("Cell").Columns.Contains("Index") Then spreadsheet.Tables("Cell").Columns.Add(New DataColumn("Index", GetType(Int32)))
                            If Not spreadsheet.Tables("Row").Columns.Contains("Index") Then spreadsheet.Tables("Row").Columns.Add(New DataColumn("Index", GetType(Int32)))

                            'Iterate through the Row table and set the row indexes
                            Dim rowIndex = 1
                            With spreadsheet.Tables("Row")
                                For i = 0 To .Rows.Count - 1
                                    If IsDBNull(.Rows(i)("Index")) Then
                                        .Rows(i)("Index") = rowIndex
                                        rowIndex += 1
                                    Else
                                        rowIndex = .Rows(i)("Index") + 1
                                    End If
                                Next
                            End With

                            'Iterate through the cell table and set the column indexes
                            rowIndex = -1
                            Dim columnIndex = 0
                            With spreadsheet.Tables("Cell")
                                For i = 0 To .Rows.Count - 1
                                    If .Rows(i)("Row_Id") <> rowIndex Then columnIndex = 1
                                    rowIndex = .Rows(i)("Row_Id")
                                    If IsDBNull(.Rows(i)("Index")) Then
                                        .Rows(i)("Index") = columnIndex
                                        columnIndex += 1
                                    Else
                                        columnIndex = .Rows(i)("Index") + 1
                                    End If
                                Next
                            End With

                            Dim cells = (From cellRecord In spreadsheet.Tables("Cell") Join rowRecord In spreadsheet.Tables("Row")
                                 On cellRecord("Row_Id") Equals rowRecord("Row_Id") Join dataRecord In spreadsheet.Tables("Data")
                                 On cellRecord("Cell_Id") Equals dataRecord("Cell_Id")
                                         Select New XMLSpreadsheetCellData With {.CellId = cellRecord("Cell_Id"),
                                                                         .RowIndex = rowRecord("Index") - 1,
                                                                         .ColumnIndex = cellRecord("Index") - 1,
                                                                         .DataType = dataRecord("Type"),
                                                                         .DataValue = dataRecord("Data_Text")})

                            For Each cell In (From entry In cells
                                              Order By entry.RowIndex, entry.CellId)
                                rowIndex = cell.RowIndex
                                columnIndex = cell.ColumnIndex
                                Select Case cell.DataType
                                    Case "String"
                                        result(rowIndex, columnIndex) = cell.DataValue
                                    Case "DateTime"
                                        result(rowIndex, columnIndex) = DateTime.Parse(cell.DataValue)
                                    Case "Number"
                                        result(rowIndex, columnIndex) = Decimal.Parse(cell.DataValue)
                                        If Decimal.Floor(result(rowIndex, columnIndex)) = result(rowIndex, columnIndex) Then
                                            result(rowIndex, columnIndex) = Integer.Parse(result(rowIndex, columnIndex))
                                        End If
                                    Case Else
                                        Throw New DataException(String.Format("XML Spreadsheet Type {0} not recognized.", cell.DataType))
                                End Select
                            Next
                        End If
                        Return result
                    End If
                    Return Nothing
                End If
            End If
            Return Nothing
        End Function

        'Initial Call Required by Constructor(s)
        Private Sub initialization()
            Try
                InitializeComponent()
                collection = Me.Resources("presentData")
                collection.Clear()

                If dg_grid1.Columns.Count > 0 Then
                    assignColumnHeader()
                Else
                    Console.WriteLine("No Columns Present in DataGrid. Create Some Columns")
                    shutdown()
                End If

            Catch ex As Exception
                Console.WriteLine("Error : MainWindow Initialization Error")
                shutdown()
            End Try

        End Sub

        Public Sub New()
            ' This call is required by the designer.
            initialization()
            defaultData_dgGrid()
        End Sub

        Private Sub defaultData_dgGrid()

            If dg_grid1.Columns.Count < 3 Then
                Console.WriteLine("Number of Columns Count is less than 1.")
                shutdown()
            End If

            Dim arrayData(dg_grid1.Columns.Count - 1) As String
            Dim obj As userData

            ''Insert the data you need to insert below. 
            ''If you Then need multiple rows To be inserted,clear the array, insert new values,and add to collection
            ''Given below is an example

            'While adding Values to array, you can skip to enter rightmost values. You cannot skip in between
            'If you want to skip a value, specify a null string in that place
            'If you are skipping values, please do pass the number of columns as an arguemnt
            'dg_grid1.Columns.Count - 1 will give you the actual number of columns present
            'IF you are specifying all the values, no need to pass the datagrid columns count

            'Initialize array Values
            arrayData = {"", "F11001", "MFM", "M1", "0002", "New", "", "6", "12", "bar", "50", "120", "200", "°C", "", "0.4", "0.8", "bar"}

            'Adding object to collection
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)

            'Method to clear Array
            Array.Clear(arrayData, 0, arrayData.Length)

            arrayData = {"", "H11001", "BAV", "S1", "0002", "New", "", "", "12", "bar", "50", "120", "200", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)
            Array.Clear(arrayData, 0, arrayData.Length)

            arrayData = {"", "H16601", "BUC", "S1", "0001", "New", "0.8", "1.2", "1.5", "bar", "20", "25", "50", "°C", "100", "150", "300", "bar"}
            obj = New userData(arrayData)
            collection.Add(obj)
            arrayData = {"", "H16632", "BUV", "S1", "0001", "New", "", "", "4", "bar", "12", "28", "50", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)

            arrayData = {"", "L11001", "DPT", "M1", "0002", "New", "1", "1.3", "2.4", "bar", "", "94.8", "200", "°C", "", "1.3", "1.5", "pascal"}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)

            arrayData = {"", "L11003", "LST", "M1", "0002", "New", "", "", "12", "bar", "", "120", "200", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)

            arrayData = {"", "L16608", "LCT", "M1", "0001", "New", "", "", "10.5", "pascal", "", "", "45", "°C", "", "", "", ""}
            obj = New userData(arrayData, dg_grid1.Columns.Count - 1)
            collection.Add(obj)

            Console.WriteLine("Columns count - " & dg_grid1.Columns.Count)
            Console.WriteLine("Attributes count - " & obj.col_list.Count)
        End Sub

        Public Sub New(_userObject_ As userData)

            ' This call is required by the designer.
            initialization()

            If _userObject_.col_list.Count > 0 Then
                collection.Add(_userObject_)
            Else
                defaultData_dgGrid()
            End If
        End Sub

        Private Sub shutdown()
            MessageBox.Show("Encountered Error! Exiting Application. Check Output Window ( Ctrl + Alt + O) for more details", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Application.Current.Shutdown(1)
        End Sub

        Public Sub New(_userObjectArray_() As userData)

            ' This call is required by the designer.
            initialization()

            If _userObjectArray_ IsNot Nothing Then
                If _userObjectArray_.Count > 0 Then
                    For Each item In _userObjectArray_
                        If item.col_list.Count = dg_grid1.Columns.Count - 1 Then
                            collection.Add(item)
                        Else
                            Console.WriteLine("Base: userData Array in MainWindow Constructor" & vbNewLine & "userData object contains too less values.")
                            Console.WriteLine("Possible Error if object created using a list of String" & vbNewLine & "Please specify all values")
                            shutdown()
                        End If
                    Next
                Else
                    Console.WriteLine("Error : Empty userData Array passed to MainWindow Constructor." & vbNewLine & "Initializing Default Values")
                    defaultData_dgGrid()
                End If
            Else
                Console.WriteLine("Error : Null Object Array Passed To MainWindow Constructor" & vbNewLine & "Initializing Default Values")
                defaultData_dgGrid()
            End If
        End Sub

        Public Sub New(arr() As String, colCount As Integer)

            ' This call is required by the designer.
            initialization()

            If colCount < dg_grid1.Columns.Count - 1 Then
                colCount = dg_grid1.Columns.Count
            End If

            Dim obj As userData = New userData(arr, colCount)
            collection.Add(obj)
        End Sub

        Private Sub assignColumnHeader()
            Try
                For i As Integer = 0 To dg_grid1.Columns.Count - 2
                    Dim col = dg_grid1.Columns.Item(i)
                    col.Header.Children.Item(0).Text = headerList(i)
                Next
            Catch ex As Exception
                Console.WriteLine("Base: MainWindow.assignColumnHeader" & vbNewLine & "Error: In assigning column headers to Datagrid Columns")
            End Try
        End Sub

        'This method is used to automatically resize the dock panel when user resizes Window
        Private Sub resizeWindow(sender As Object, e As SizeChangedEventArgs)
            pnl_dock.Width = win_main.ActualWidth
            pnl_dock.Height = win_main.ActualHeight
            tabControl1.Width = win_main.ActualWidth
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
                Console.WriteLine("Error : Unable to Release Object. Possible Memory Lock")
                Console.WriteLine(ex.Message)
                shutdown()
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

                If configHeaderList.Count <= 0 Then
                    Console.WriteLine("Base : fileRead" & vbNewLine & "Empty configHeaderlist. File format doesn't match with the one specified. Please Check")
                    shutdown()
                End If

            Catch e As FileNotFoundException
                Console.WriteLine("Unable to find or Create configuration file" & vbNewLine & "Not enough access to the current Directory")
                'shutdown()
            Catch e As Exception
                Console.WriteLine("Base: MainWindow.fileRead " & vbNewLine & "Error : While Reading Configuration file from Directory")
                Console.WriteLine(e.Message)
                'shutdown()
            End Try
        End Sub

        Private Function setValues(ByRef obj As userData) As userData
            For i As Integer = 0 To obj.col_list.Count - 1
                If obj.col_list.Item(i) Is Nothing Then
                    obj.col_list.Item(i) = ""
                End If
            Next
            Return obj
        End Function

        Public Function GetDataGridCell(cellInfo As DataGridCellInfo) As DataGridCell
            Try
                Dim cellContent As Object = cellInfo.Column.GetCellContent(cellInfo.Item)
                If (cellContent IsNot Nothing) Then
                    Return CType(cellContent.Parent, DataGridCell)
                End If
                Return Nothing
            Catch ex As InvalidCastException
                Console.WriteLine("Base: GetDataGridCell" & vbNewLine & "Error : Invalid Cast Exeption")
                Return Nothing
                shutdown()
            Catch ex As Exception
                Console.WriteLine("Error: GetDataGridCell")
                Console.WriteLine(ex.Message)
                Return Nothing
                shutdown()
            End Try
        End Function

        Public Function GetRowIndexFromCell(dataGrid As DataGrid, dataGridCellInfo As DataGridCellInfo) As Integer
            Try
                Dim dgrow As DataGridRow = DirectCast(dataGrid.ItemContainerGenerator.ContainerFromItem(dataGridCellInfo.Item), DataGridRow)
                If dgrow IsNot Nothing Then
                    Return dgrow.GetIndex()
                Else
                    Return -1
                End If
            Catch ex As InvalidCastException
                Console.WriteLine("Base: GetRowIndexFromCell" & vbNewLine & "Error : Invalid DirectCast Exeption")
                Return -1
                shutdown()

            Catch ex As Exception
                Console.WriteLine("Error: GetRowIndexFromCell")
                Console.WriteLine(ex.Message)
                Return -1
                shutdown()
            End Try
        End Function

        Public Sub changeCellColor(dataCellInfo As DataGridCellInfo, bgColor As Color, fgColor As Color)
            Dim dataCell As DataGridCell = GetDataGridCell(dataCellInfo)

            If bgColor.Equals(Colors.Black) Then
                dataCell.BorderThickness = New Thickness(0.0)
            Else
                dataCell.BorderThickness = New Thickness(3.0)
            End If

            If bgColor = Colors.DarkRed Then
                darkRedCellsColored.Add(dataCell)
                valHighlight = True
            End If

            If bgColor = Colors.Violet Then
                violetcellsColored.Add(dataCell)
                intHighlight = True
            End If

            If bgColor = Colors.Blue Then
                blueCellsColored.Add(dataCell)
                blueHighlight = True
            End If

            If bgColor = Colors.Red Then
                redcellsColored.Add(dataCell)
                errorHighlight = True
            End If

            dataCell.BorderBrush = New SolidColorBrush(bgColor)
        End Sub

        Private Function determineIndex(header As String, Optional value As Integer = 0) As Integer
            Console.WriteLine("Header = " & header)
            Dim list() As String = headerList
            If value = 1 Then
                list = btnNamesArray
            End If

            For i As Integer = 0 To dg_grid1.Columns.Count - 2
                If (list(i).ToLower()).Equals(header) Then
                    Console.WriteLine("i = " & i)
                    Return i
                End If
            Next
            Return -1
        End Function


        Private Sub win_main_Initialized(sender As Object, e As EventArgs)

            If dg_grid1.Columns.Count < 3 Then
                MessageBox.Show("Critical Error : No Columns Present in Datagrid. Please Add Some", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Application.Current.Shutdown(1)
            End If

            If headerList.Length <> dg_grid1.Columns.Count - 1 Or (btnNamesArray.Length) <> dg_grid1.Columns.Count - 1 Or colNames.Length <> dg_grid1.Columns.Count - 1 Then
                MessageBox.Show("Critical Error : Insufficient Column ids or Button Name IDs or Headers List", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Application.Current.Shutdown(1)
            End If

            For i As Integer = 0 To headerList.Length - 1
                If (headerList(i).ToLower()).Equals("measuring principle") Then
                    measuringIndex = i
                ElseIf (headerList(i).ToLower()).Equals("item tag name") Then
                    itemNameIndex = i
                End If
            Next

        End Sub

        ''**********End of Helper Functions

        ''Methods Pertaining to FilterData in DataGrid

        Private Sub openFilterWindow(sender As Object, e As RoutedEventArgs)

            Try
                Dim dummyUserData As userData = New userData(Nothing, dg_grid1.Columns.Count - 1)
                Dim list As List(Of String) = New List(Of String)
                If headerSelected.ToLower.Equals("measuring principle") Then
                    list = dummyUserData.measuringprinciple
                ElseIf headerSelected.ToLower.Equals("unit of pressure p1") Then
                    list = dummyUserData.pressureUnit
                ElseIf headerSelected.ToLower.Equals("unit of temperature") Then
                    list = dummyUserData.tempUnit
                ElseIf headerSelected.ToLower.Equals("unit of differential pressure") Then
                    list = dummyUserData.diffPressureUnit
                End If
                Dim inputDialog As FilterWindow
                If list.Count = 0 Then
                    inputDialog = New FilterWindow(headerSelected)
                Else
                    inputDialog = New FilterWindow(list, headerSelected)
                End If
                inputDialog.ShowInTaskbar = True
                inputDialog.Owner = Me

                If inputDialog.ShowDialog = True Then
                    filterSelected = True
                    caseSensitive = inputDialog.returnCaseSensitive()
                    cbCompleteFilter.IsEnabled = True
                    cbCompleteFilter.IsChecked = True
                    filterValue = inputDialog.returnFilterValue()

                    If Not filterValue.Equals("") Then
                        CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                        filterStatus.Content = "Currently Filter Is applied to Column : " & headerSelected & " with Value : " & filterValue
                        filterStatus.Visibility = Visibility.Visible
                    Else
                        Console.WriteLine("Base: openFilterWindow" & vbNewLine & "Invalid filter Value obtained from FilterWindow")
                        shutdown()
                    End If
                Else
                    Console.WriteLine("Base : openFilterWindow" & vbNewLine & "FilterWindow encountered an error. ShowDialog Returned False")
                    'shutdown()
                End If
            Catch ex As Exception
                Console.WriteLine("Base: openFilterWindow")
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try

        End Sub

        ''Called When user access the filter Window by using the context Menu for headers
        Private Sub menuFilter_Click(sender As Object, e As RoutedEventArgs)
            openFilterWindow(sender, e)
        End Sub

        Private Sub btn_filter_Click(sender As Object, e As RoutedEventArgs)
            If TypeOf sender Is System.Windows.Controls.Button Then
                Dim btn As System.Windows.Controls.Button = sender
                If btn.Name = "" Then
                    btn.Name = "btn_filter_name"
                End If
                Dim counter As Integer = determineIndex(btn.Name.ToLower(), 1)
                If counter = -1 Then
                    Console.WriteLine("Base : btn_filter_Click" & vbNewLine & "Error : Button Names List Initialized in DataGridColumns (xaml) doesn't match with btnNamesArray in MainWindow")
                    shutdown()
                Else
                    headerSelected = headerList(counter)
                End If

                'Call FilterWindow
                openFilterWindow(sender, e)
            Else
                Console.WriteLine("Sender is Not Button for btn_filter_Click" & vbNewLine & "Invalid Invokation")
                shutdown()
            End If
        End Sub

        Private Sub CompleteFilter_Changed(sender As Object, e As RoutedEventArgs)
            Try
                If cbCompleteFilter.IsChecked = False Then
                    CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                    filterStatus.Content = "Currently No Filters Have been Applied"
                Else
                    CollectionViewSource.GetDefaultView(dg_grid1.ItemsSource).Refresh()
                    filterStatus.Content = "Currently Filter Is applied to Column  " & headerSelected & " with Value : " & filterValue
                End If
                filterStatus.Visibility = Visibility.Visible
            Catch ex As Exception
                Console.WriteLine("Base: CompleteFilter_Changed" & vbNewLine & "Error : While Refreshing COllectionViewSource")
                Console.WriteLine("Details : " & ex.Message)
                shutdown()
            End Try
        End Sub

        Private Sub collectionviewsource_filter(sender As Object, e As FilterEventArgs)
            Try

                If filterSelected Then
                    Dim obj As userData = e.Item
                    If obj IsNot Nothing Then
                        obj = setValues(obj)
                        If obj.col_list.Count > 0 Then
                            If cbCompleteFilter.IsChecked = True Then
                                Dim headerIndex As Integer = determineIndex(headerSelected.ToLower())
                                If headerIndex = -1 Then
                                    Console.WriteLine("Base : collectionviewsource_filter" & vbNewLine & "Error : Header Name Initialized to DataGrid columns doesn't match with headerList")
                                    shutdown()
                                End If
                                If caseSensitive = True Then
                                    If obj.col_list.Item(headerIndex).Equals(filterValue) Then
                                        e.Accepted = True
                                    Else
                                        e.Accepted = False
                                    End If
                                Else
                                    If (obj.col_list.Item(headerIndex).ToLower()).Equals(filterValue.ToLower) Then
                                        e.Accepted = True
                                    Else
                                        e.Accepted = False
                                    End If
                                End If
                            End If
                        Else
                            Console.WriteLine("Base : collectionviewsource_filter" & vbNewLine & "Error : userData object contains empty or invalid List")
                            Console.WriteLine("Possibly Empty DataGrid")
                            shutdown()
                        End If
                    Else
                        Console.WriteLine("Base : collectionviewsource_filter" & vbNewLine & "Error : userData object obtained from e.Item is Nothing")
                        Console.WriteLine("Possibly Empty DataGrid")
                        shutdown()
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                shutdown()
            End Try
        End Sub

        ''End of Methods pertaining to Filtering Data in DataGrid

        ''This method determines the where user performed mouse right click And stores the location Of click.
        Private Sub columnHeader_PreviewMouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseRightButtonUp
            'HighlightClear()
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

        'This method determines the where user performed mouse left click and stores the location of click.
        'More specifically, it is used to determine on which row, the user has clicked the Add Row Button
        Private Sub detectCellClicked(sender As Object, e As MouseButtonEventArgs) Handles dg_grid1.PreviewMouseLeftButtonUp

            'HighlightClear()

            Dim dep As DependencyObject = e.OriginalSource

            'Sender is always datagrid, no matter where user click in the table ( datagrid)
            'To get the exact sender, we must traverse the DataTree to DatagridCell

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
            ElseIf (TypeOf dep Is Primitives.DataGridColumnHeader) Then

            Else
                Console.WriteLine("Unable to determine cell clicked")
            End If

            'By default e.Handled is True. It signifies that right click has been handled, and sometimes context menu don't pop up
            'To avoid this, set it to False, so context menu gets visible
            e.Handled = False

            If colEditIndex = dg_grid1.Columns.Count - 1 Then
                rowAdd = True
            Else
                rowAdd = False
            End If

            Return
        End Sub

        Private Sub btn_edit_row_Click(sender As Object, e As RoutedEventArgs)
            Console.WriteLine("Adding Row " & collection.Count)
            If colEditIndex = dg_grid1.Columns.Count - 1 Then
                collection.Insert(rowEditIndex + 1L, New userData(Nothing, dg_grid1.Columns.Count - 1))
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
            Try
                Dim ifCut As Boolean = False
                Dim currentCell As DataGridCellInfo = dg_grid1.SelectedCells.Item(0)
                Dim _currColIndx = currentCell.Column.DisplayIndex
                Dim _currRowIndx = GetRowIndexFromCell(dg_grid1, currentCell)
                Dim curruserData As userData = collection.Item(_currRowIndx)

                If previousSelectedCells.Count > 0 Then
                    Dim _rowIndx = GetRowIndexFromCell(dg_grid1, previousSelectedCells.Item(0))
                    Dim prevuserData As userData = collection.Item(_rowIndx)
                    Dim _colIndx As Integer = 0
                    For i As Integer = 0 To previousSelectedCells.Count - 1
                        Dim cell As DataGridCellInfo = previousSelectedCells.Item(i)
                        _colIndx = cell.Column.DisplayIndex
                        curruserData.col_list.Item(_currColIndx) = prevuserData.col_list.Item(_colIndx)

                        If cutActivated Then
                            prevuserData.col_list.Item(_colIndx) = ""
                            ifCut = True
                        End If
                        _currColIndx += 1
                    Next

                    If ifCut Then
                        cutActivated = False
                    End If

                    If _rowIndx > _currRowIndx Then
                        collection.RemoveAt(_rowIndx)
                        collection.RemoveAt(_currRowIndx)

                        collection.Insert(_currRowIndx, curruserData)
                        collection.Insert(_rowIndx, prevuserData)
                    ElseIf _rowIndx < _currRowIndx Then
                        collection.RemoveAt(_currRowIndx)
                        collection.RemoveAt(_rowIndx)

                        collection.Insert(_rowIndx, prevuserData)
                        collection.Insert(_currRowIndx, curruserData)
                    Else
                        collection.RemoveAt(_rowIndx)
                        collection.Insert(_currRowIndx, curruserData)
                    End If

                    dg_grid1.Focus()
                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(_currRowIndx), dg_grid1.Columns(_currColIndx - 1))
                    dg_grid1.SelectedCells.Add(dg_grid1.CurrentCell)


                End If
            Catch ex As Exception
                Console.WriteLine("Base : RetrieveCells" & vbNewLine & "Error in Pasting values!")
                shutdown()
            End Try

        End Sub
        Private Sub DeleteCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)
            deleteItem(sender, e)
        End Sub

        Private Sub deleteItem(sender As Object, e As ExecutedRoutedEventArgs)
            Try
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
                dg_grid1.SelectedCells.Clear()
                dg_grid1.SelectedCells.Add(cell)
                dg_grid1.CurrentCell = cell
                dg_grid1.Focus()
            Catch ex As Exception
                Console.WriteLine("Base : deleteItem" & vbNewLine & "Unable to delete Values")
                shutdown()
            End Try

        End Sub

        Private Sub PasteCommand_Executed(sender As Object, e As ExecutedRoutedEventArgs)

            RetrieveCells(sender, e)
            pasteActivated = True
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

        Private Sub highlightAllRows_Click(sender As Object, e As RoutedEventArgs)
            allHighlight = True
            _highlightCells(sender, e, True)
        End Sub

        Private Sub clearHighlightCells_Click(sender As Object, e As RoutedEventArgs)
            HighlightClear()
        End Sub

        Private Sub _highlightCells(sender As Object, e As RoutedEventArgs, nRows As Boolean)
            blueHighlight = False
            blueFlag = False
            Dim _tempRow As Integer = rowEditIndex
            Dim _end As Integer = _tempRow
            If nRows Then
                _end = collection.Count - 1
                _tempRow = 0
            End If
            Try
                fileRead()
                For _rowIndx As Integer = _tempRow To _end
                    Dim obj As userData = collection.Item(_rowIndx)
                    If obj IsNot Nothing Then
                        obj = setValues(obj)
                        If obj.col_list.Count > 0 Then
                            Try
                                For Each List In configHeaderList
                                    If List.Item(0) IsNot Nothing Then
                                        If obj.col_list.Item(measuringIndex).Equals(List.Item(0)) Then
                                            dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(_rowIndx), dg_grid1.Columns.Item(measuringIndex))
                                            changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                                            For i As Integer = 1 To List.Count - 1
                                                For j As Integer = 2 To dg_grid1.Columns.Count - 2
                                                    If (List.Item(i).ToLower()).Equals(headerList(j).ToLower()) Then
                                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(_rowIndx), dg_grid1.Columns.Item(j))
                                                        changeCellColor(dg_grid1.CurrentCell, Colors.Blue, Colors.White)
                                                    End If
                                                Next
                                            Next
                                            Exit For
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                Console.WriteLine("Base : highlightCells_Click" & vbNewLine & "Exception while highlighting Cells")
                                'shutdown()
                            End Try
                        Else
                            Console.WriteLine("Base : highlightCells_Click" & vbNewLine & "Object doesn't contain any column Values. Invalid!!")
                            'shutdown()
                        End If
                    Else
                        Console.WriteLine("Base : highlightCells_Click" & vbNewLine & "Null Object obtained. Possbile Empty Datagrid")
                        'shutdown()
                    End If
                Next
            Catch ex As Exception
                Console.WriteLine("Base : highlightCells_Click" & vbNewLine & "Exception ")
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try
        End Sub

        Private Sub highlightCells_Click(sender As Object, e As RoutedEventArgs)
            _highlightCells(sender, e, False)
        End Sub

        ''********End Of Row Context Menu tasks

        ''ComboBox Selection Changed
        Private Sub changeComboBoxValue(sender As Object, e As SelectionChangedEventArgs)
            If rowAdd = False Then
                If TypeOf sender Is ComboBox Then
                    Dim comboBox As ComboBox = CType(sender, ComboBox)
                    If comboBox.SelectedValue IsNot Nothing Then
                        If rowEditIndex >= 0 Then
                            If colEditIndex >= 0 AndAlso colEditIndex < dg_grid1.Columns.Count - 1 Then
                                collection.Item(rowEditIndex).col_list.Item(colEditIndex) = comboBox.SelectedValue.ToString
                            End If
                            If colEditIndex = measuringIndex Then
                                highlightCells_Click(New Object(), New RoutedEventArgs())
                            End If
                        End If
                    End If
                End If
            End If

        End Sub
        ''End of ComboBox Event Handlers

        ''Excel Related Files
        Private Sub btn_export_Click(sender As Object, e As RoutedEventArgs)
            exportExcel()
        End Sub

        Private Sub comboLoaded(sender As Object, e As RoutedEventArgs)
            If sender.SelectedValue Is Nothing Then
                sender.SelectedValue = ""
            End If
        End Sub

        Private Sub exportExcel()
            '' Main Content
            Dim f As SaveFileDialog = New SaveFileDialog()
            f.Filter = "Excel Workbook (*.xlsx) |*.xlsx|All files (*.*)|*.*"

            For i As Integer = 0 To collection.Count - 1
                setValues(collection.Item(i))
            Next
            Try
                If f.ShowDialog() = True Then
                    Dim xlApp As Excel.Application = New Excel.Application()
                    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add
                    Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(1)
                    Dim colCount = dg_grid1.Columns.Count, rowCount = collection.Count
                    Dim temp As Integer = 0, added As Boolean = True
                    'Create an array with n columns and n rows
                    Dim DataArray(rowCount, colCount - 2) As Object

                    For row As Short = 0 To rowCount - 1
                        If Not discardList.Contains(row) Then
                            For col As Short = 0 To colCount - 2
                                DataArray(temp, col) = collection.Item(row).col_list.Item(col)
                                added = True
                            Next
                        End If
                        If added Then
                            temp = temp + 1
                            added = False
                        End If
                    Next

                    xlWorkSheet.Range("A1").Resize(1, colCount).Value = headerList
                    xlWorkSheet.Range("A2").Resize(temp, colCount).Value = DataArray


                    xlWorkSheet.SaveAs(f.FileName)
                    xlWorkBook.Close()
                    xlApp.Quit()

                    MessageBox.Show("Published!.")

                    releaseObject(xlApp)
                    releaseObject(xlWorkBook)
                    releaseObject(xlWorkSheet)
                End If
            Catch ex As Exception
                MessageBox.Show("Unable to export", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try
        End Sub



        Private Sub chk_include_Checked(sender As Object, e As RoutedEventArgs)
            collection.Item(rowEditIndex).include = True
            If discardList.Contains(rowEditIndex) Then
                discardList.Remove(rowEditIndex)
            End If
        End Sub

        Private Sub chk_include_Unchecked(sender As Object, e As RoutedEventArgs)
            collection.Item(rowEditIndex).include = False
            If Not discardList.Contains(rowEditIndex) Then
                discardList.Add(rowEditIndex)
            End If

        End Sub

        Private Sub btn_import_click(sender As Object, e As RoutedEventArgs)
            Try
                Dim openfiledialog As OpenFileDialog = New OpenFileDialog()
                openfiledialog.Filter = "Excel workbook (*.xlsx) |*.xlsx|All files (*.*)|*.*"
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
            Catch ex As Exception
                Console.WriteLine("Base: btn_import_click")
                Console.WriteLine(ex.Message)
                shutdown()
            End Try

        End Sub

        Private Sub displayexcelfile(filename As String, sheetname As String)
            Try
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
                    'Console.WriteLine("length {0}", array.Length)

                    ' get bounds of the array.
                    Dim bound0 As Integer = array.GetUpperBound(0)
                    Dim bound1 As Integer = array.GetUpperBound(1)

                    'Console.WriteLine("dimension 0 {0}", bound0)
                    'Console.WriteLine("dimension 1 {0}", bound1)
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
            Catch ex As Exception
                Console.WriteLine("Base: displayExcelFile" & vbNewLine & "Error while importing file")
                Console.WriteLine(ex.Message)
                shutdown()
            End Try

        End Sub
        ''Excel Related Functions

        Private Sub btn_save_click(sender As Object, e As RoutedEventArgs)
            Console.WriteLine(collection.Item(0).include.ToString)
        End Sub

        Private Function getexcelsheetnames(ByVal filename As String) As List(Of String)
            Try
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
            Catch ex As Exception
                Console.WriteLine("Base: getexcelsheetnames" & vbNewLine & "Error while getting excel sheet names")
                Console.WriteLine(ex.Message)
                Return Nothing
                shutdown()
            End Try
        End Function


        Private Sub btn_validate_Click(sender As Object, e As RoutedEventArgs)
            Try
                validate_Mandatory2(collection.Count)

                If errorHighlight = False Then
                    validate_integerValue(collection.Count)
                End If

                If intHighlight = False Then
                    validate_comparison(collection.Count)
                End If

                filterStatus.Visibility = Visibility.Hidden
                errorStatus.Content = ""
                If errorHighlight = True Then
                    errorStatus.Content = errorStatus.Content & vbTab & "Cells Highlighted in Red Must Not be left Blank"
                End If

                If valHighlight = True Then
                    errorStatus.Content = errorStatus.Content & vbTab & "Value highlighted in Dark Red should be in order: Min < Operation < Max"
                End If

                If intHighlight = True Then
                    errorStatus.Content = errorStatus.Content & vbTab & "Value highlighted in Violet must be a numeric value"
                End If

                If intHighlight = False AndAlso valHighlight = False AndAlso errorHighlight = False Then
                    MsgBox("Validation Ok")
                Else
                    MsgBox("Validation not Ok")
                End If

            Catch ex As Exception
                MsgBox("Base : btn_validate_click" & vbNewLine)
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try

        End Sub

        Private Sub validate_integerValue(nRows As Integer)
            For Each cell In violetcellsColored
                If blueCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Blue)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf darkRedCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.DarkRed)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf redcellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Red)
                    cell.BorderThickness = New Thickness(3.0)
                Else
                    cell.BorderBrush = New SolidColorBrush(Colors.Black)
                    cell.BorderThickness = New Thickness(0.0)
                End If
            Next
            violetcellsColored.Clear()
            intHighlight = False
            Dim result As Integer = 0
            Dim columnsToValidate() As String =
            {
            "Pressure P1 Minimum", "Pressure P1 In Operation", "Pressure P1 Maximum",
             "Temperature Minimum", "Temperature In Operation", "Temperature Maximum",
            "Differential Pressure Minimum", "Differential Pressure In Operation", "Differential Pressure Maximum"}

            Dim indexArray As List(Of Integer) = New List(Of Integer)()

            Try
                For i As Integer = 0 To headerList.Length - 1
                    For j As Integer = 0 To columnsToValidate.Length - 1
                        If (headerList(i).ToLower()).Equals(columnsToValidate(j).ToLower()) Then
                            indexArray.Add(i)
                            Exit For
                        End If
                    Next
                Next
                For counter As Integer = 0 To nRows - 1
                    Dim temp_userdata As userData = collection.Item(counter)
                    If temp_userdata IsNot Nothing Then
                        temp_userdata = setValues(temp_userdata)
                        If temp_userdata.col_list.Count > 0 Then
                            For i As Integer = 0 To indexArray.Count - 1
                                If temp_userdata.col_list.Item(indexArray.Item(i)).Equals("") Then
                                    Continue For
                                End If
                                If Not IsNumeric(temp_userdata.col_list.Item(indexArray.Item(i))) Then
                                    dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(indexArray.Item(i)))
                                    changeCellColor(dg_grid1.CurrentCell, Colors.Violet, Colors.White)
                                End If
                            Next
                        Else
                            MsgBox("Base : validate_integerValue" & vbNewLine & "Object doesn't contain any column Values. Invalid!!")
                            'shutdown()
                        End If
                    Else
                        MsgBox("Base : validate_integerValue" & vbNewLine & "Null Object obtained. Possbile Empty Datagrid")
                        'shutdown()
                    End If
                Next
            Catch ex As Exception
                MsgBox("Base : validate_integerValue" & vbNewLine & "Exception Thrown")
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try

        End Sub

        Private Sub validate_Mandatory(nRows As Integer)
            fileRead()
            For Each cell In redcellsColored
                If blueCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Blue)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf darkRedCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.DarkRed)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf violetcellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Violet)
                    cell.BorderThickness = New Thickness(3.0)
                Else
                    cell.BorderBrush = New SolidColorBrush(Colors.Black)
                    cell.BorderThickness = New Thickness(0.0)
                End If
            Next
            redcellsColored.Clear()
            errorHighlight = False
            Try
                Dim indexList As List(Of Integer) = New List(Of Integer)()
                Dim pgotIt As Boolean = False
                'Get the index of Unit of Pressure and Unit of Temperature to check for values

                For i As Integer = 0 To headerList.Length - 1
                    If (headerList(i).ToLower()).Contains("unit of pressure") Then
                        indexList.Add(i)
                        Continue For
                    End If
                    If (headerList(i).ToLower()).Contains("unit of temperature") Then
                        indexList.Add(i)
                        Continue For
                    End If
                Next

                If indexList.Count < 2 Then
                    MsgBox("Error : Header Name Mismatch" & vbNewLine & "Check column header names for Unit of Pressure & Temperature in both xaml")
                    'shutdown()
                End If

                For counter As Integer = 0 To nRows - 1
                    'Get the current Row userData object
                    Dim temp_userdata As userData = collection.Item(counter)
                    If temp_userdata IsNot Nothing Then
                        temp_userdata = setValues(temp_userdata)
                        If temp_userdata.col_list.Count > 0 Then
                            If Not temp_userdata.col_list.Item(measuringIndex).Equals("") Then
                                For Each list In configHeaderList
                                    If temp_userdata.col_list.Item(measuringIndex).Equals(list.Item(0)) Then
                                        'Console.WriteLine("Selection Equals " & list.Item(0))
                                        For i As Integer = 1 To list.Count - 1
                                            For j As Integer = 0 To dg_grid1.Columns.Count - 2
                                                If (list.Item(i).ToLower()).Equals(headerList(j).ToLower()) Then
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
                            Else
                                errorHighlight = True
                            End If
                            ' Check if Unit of Pressure / Unit of Temperature is Empty or Not
                            'If empty highlight it 
                            If Not temp_userdata.col_list.Item(measuringIndex).Equals("") Then
                                For Each item In indexList
                                    If temp_userdata.col_list(item).Equals("") Then
                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(item))
                                        changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                    End If
                                Next
                            End If
                        Else
                            MsgBox("Base : validate_Mandatory" & vbNewLine & "Object doesn't contain any column Values. Invalid!!")
                            'shutdown()
                        End If
                    Else
                        MsgBox("Base : validate_Mandatory" & vbNewLine & "Null Object obtained. Possbile Empty Datagrid")
                        'shutdown()
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show("Enter values in highlighted fields", "Validation Message", MessageBoxButton.OK, MessageBoxImage.Error)
                'MsgBox("Base : validate_Mandatory" & vbNewLine & "Exception Thrown")
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try

        End Sub

        Private Sub validate_Mandatory2(nRows As Integer)
            fileRead()
            For Each cell In redcellsColored
                If blueCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Blue)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf darkRedCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.DarkRed)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf violetcellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Violet)
                    cell.BorderThickness = New Thickness(3.0)
                Else
                    cell.BorderBrush = New SolidColorBrush(Colors.Black)
                    cell.BorderThickness = New Thickness(0.0)
                End If
            Next
            redcellsColored.Clear()
            errorHighlight = False

            Try
                Dim indexList As List(Of Integer) = New List(Of Integer)()
                Dim pgotIt As Boolean = False
                'Get the index of Unit of Pressure and Unit of Temperature to check for values

                For i As Integer = 0 To headerList.Length - 1
                    If (headerList(i).ToLower()).Contains("unit of pressure") Then
                        indexList.Add(i)
                        Continue For
                    End If
                    If (headerList(i).ToLower()).Contains("unit of temperature") Then
                        indexList.Add(i)
                        Continue For
                    End If
                    If (headerList(i).ToLower()).Contains("unit of differential pressure") Then
                        indexList.Add(i)
                        Continue For
                    End If
                Next

                If indexList.Count < 3 Then
                    MsgBox("Error : Header Name Mismatch" & vbNewLine & "Check column header names for Unit of Pressure & Temperature in both xaml")
                    'shutdown()
                End If

                For counter As Integer = 0 To nRows - 1
                    'Get the current Row userData object
                    Dim temp_userdata As userData = collection.Item(counter)
                    If temp_userdata IsNot Nothing Then
                        temp_userdata = setValues(temp_userdata)
                        If temp_userdata.col_list.Count > 0 Then
                            For Each item In indexList
                                If temp_userdata.col_list(item).Equals("") Then
                                    If Not (temp_userdata.col_list(item - 1).Equals("") AndAlso temp_userdata.col_list(item - 2).Equals("") AndAlso temp_userdata.col_list(item - 3).Equals("")) Then
                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(item))
                                        changeCellColor(dg_grid1.CurrentCell, Colors.Red, Colors.White)
                                    End If
                                End If
                            Next
                        End If
                    Else
                        MsgBox("Base : validate_Mandatory" & vbNewLine & "Null Object obtained. Possbile Empty Datagrid")
                        'shutdown()
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show("Enter values in highlighted fields", "Validation Message", MessageBoxButton.OK, MessageBoxImage.Error)
                Console.WriteLine(ex.Message)
                'shutdown()
            End Try

        End Sub

        Private Sub validate_comparison(nRows As Integer)
            fileRead()
            For Each cell In darkRedCellsColored
                If blueCellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Blue)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf redcellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Red)
                    cell.BorderThickness = New Thickness(3.0)
                ElseIf violetcellsColored.Contains(cell) Then
                    cell.BorderBrush = New SolidColorBrush(Colors.Violet)
                    cell.BorderThickness = New Thickness(3.0)
                Else
                    cell.BorderBrush = New SolidColorBrush(Colors.Black)
                    cell.BorderThickness = New Thickness(0.0)
                End If
            Next
            darkRedCellsColored.Clear()
            valHighlight = False

            Try
                For j As Integer = 0 To 3
                    Dim minVal As Double, normVal As Double, maxVal As Double
                    Dim minIndex As Integer = -1, normIndex As Integer = -1, maxIndex As Integer = -1

                    If j = 0 Then
                        For i As Integer = 0 To dg_grid1.Columns.Count - 2
                            If headerList(i).Equals("Pressure P1 Minimum") Then
                                minIndex = i
                                If headerList(i + 1).Equals("Pressure P1 In Operation") Then
                                    normIndex = i + 1
                                    If headerList(i + 2).Equals("Pressure P1 Maximum") Then
                                        maxIndex = i + 2
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    ElseIf j = 1 Then
                        For i As Integer = 0 To dg_grid1.Columns.Count - 2
                            If headerList(i).Equals("Temperature Minimum") Then
                                minIndex = i
                                If headerList(i + 1).Equals("Temperature In Operation") Then
                                    normIndex = i + 1
                                    If headerList(i + 2).Equals("Temperature Maximum") Then
                                        maxIndex = i + 2
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    ElseIf j = 2 Then
                        For i As Integer = 0 To dg_grid1.Columns.Count - 2
                            If headerList(i).Equals("Differential Pressure Minimum") Then
                                minIndex = i
                                If headerList(i + 1).Equals("Differential Pressure In Operation") Then
                                    normIndex = i + 1
                                    If headerList(i + 2).Equals("Differential Pressure Maximum") Then
                                        maxIndex = i + 2
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    Else
                        Exit For
                    End If

                    If minIndex = -1 Or maxIndex = -1 Or normIndex = -1 Then
                        MsgBox("Error : Header Name Mismatch" & vbNewLine & "Check column header names for Pressure, Temperature & Differential Pressure")
                        'shutdown()
                    End If

                    For counter As Integer = 0 To nRows - 1
                        Dim temp_userdata As userData = collection.Item(counter)

                        If temp_userdata IsNot Nothing Then
                            temp_userdata = setValues(temp_userdata)
                            If temp_userdata.col_list.Count > 0 Then
                                Dim _tempmin As String, _tempnorm As String, _tempmax As String
                                _tempmin = temp_userdata.col_list.Item(minIndex)
                                _tempnorm = temp_userdata.col_list.Item(normIndex)
                                _tempmax = temp_userdata.col_list.Item(maxIndex)

                                If _tempmin.Equals("") AndAlso _tempnorm.Equals("") AndAlso _tempmax.Equals("") Then
                                    Continue For
                                End If

                                minVal = convertDouble(_tempmin)
                                normVal = convertDouble(_tempnorm)
                                maxVal = convertDouble(_tempmax)

                                If (minVal > normVal) Or (normVal > maxVal) Or (minVal > maxVal) Then
                                    For i As Integer = 0 To 2
                                        dg_grid1.CurrentCell = New DataGridCellInfo(dg_grid1.Items(counter), dg_grid1.Columns.Item(minIndex + i))
                                        changeCellColor(dg_grid1.CurrentCell, Colors.DarkRed, Colors.White)
                                    Next
                                End If
                            Else
                                MsgBox("Base : validate_comparison" & vbNewLine & "Object doesn't contain any column Values. Invalid!!")
                                'shutdown()
                            End If
                        Else
                            Console.WriteLine("Base : validate_comparison" & vbNewLine & "Null Object obtained. Possbile Empty Datagrid")
                            shutdown()
                        End If
                    Next
                Next
            Catch ex As Exception
                Console.WriteLine("Base : validate_comparison" & vbNewLine & "Exception Thrown")
                Console.WriteLine(ex.Message)
                shutdown()
            End Try
        End Sub

        Private Sub btn_close_Click(sender As Object, e As RoutedEventArgs)
            Application.Current.MainWindow.Close()

        End Sub

    End Class
End Namespace