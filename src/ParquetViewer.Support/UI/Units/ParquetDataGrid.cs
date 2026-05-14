using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace ParquetViewer.Support.UI.Units
{
    public class ParquetDataGrid : DataGrid
    {
        public static readonly DependencyProperty RowFilterProperty =
            DependencyProperty.Register(nameof(RowFilter), typeof(string), typeof(ParquetDataGrid),
                new PropertyMetadata(string.Empty, OnRowFilterChanged));

        public string RowFilter
        {
            get => (string)GetValue(RowFilterProperty);
            set => SetValue(RowFilterProperty, value);
        }

        public ParquetDataGrid()
        {
            DefaultStyleKey = typeof(ParquetDataGrid);
            AutoGenerateColumns = true;
            IsReadOnly = true;
            CanUserAddRows = false;
            CanUserDeleteRows = false;
        }

        private static void OnRowFilterChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ParquetDataGrid grid && grid.ItemsSource is DataView view)
            {
                view.RowFilter = e.NewValue as string;
            }
        }

        protected override void OnAutoGeneratingColumn(DataGridAutoGeneratingColumnEventArgs e)
        {
            base.OnAutoGeneratingColumn(e);
            e.Column.MaxWidth = 400;
        }

        protected override void OnItemsSourceChanged(System.Collections.IEnumerable oldValue, System.Collections.IEnumerable newValue)
        {
            base.OnItemsSourceChanged(oldValue, newValue);

            if (newValue is DataView view && !string.IsNullOrEmpty(RowFilter))
            {
                view.RowFilter = RowFilter;
            }
        }
    }
}
