Imports System.Windows.Interop
Imports System.Windows.Threading

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

    Private Sub header_rotate(sender As Object, e As MouseButtonEventArgs) Handles
             menu_colVertical.MouseLeftButtonDown



    End Sub

    'Code to rotate text'
    'Private void Form1_Load(Object sender, EventArgs e)
    '    {
    '        DataGridViewTextBoxColumn tc = New DataGridViewTextBoxColumn();
    '        tc.HeaderText = "Hello\nWorld";
    '        tc.Width = 50;

    '        dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
    '        dataGridView1.ColumnHeadersHeight = 100;

    '        dataGridView1.Columns.Add(tc);
    '        For (int x = 0; x < 10; x++)
    '        {
    '            dataGridView1.Rows.Add();
    '            dataGridView1[0, x].Value = x.ToString();
    '        }
    '    }

    '    Private void dataGridView1_CellPainting(Object sender, DataGridViewCellPaintingEventArgs e)
    '    {
    '        If (e.RowIndex == -1 && e.ColumnIndex == 0)
    '        {
    '            e.PaintBackground(e.CellBounds, true);
    '            e.Graphics.TranslateTransform(e.CellBounds.Left , e.CellBounds.Bottom);
    '            e.Graphics.RotateTransform(270);
    '            e.Graphics.DrawString(e.FormattedValue.ToString(),e.CellStyle.Font,Brushes.Black,5,5);
    '            e.Graphics.ResetTransform();
    '            e.Handled=true;
    '        }
    '    }




End Class
