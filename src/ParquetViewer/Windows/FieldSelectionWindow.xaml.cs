using ParquetViewer.Controls;
using ParquetViewer.Form.Local.ViewModels;
using System.Collections.Generic;

namespace ParquetViewer.Windows
{
    public partial class FieldSelectionWindow : WindowBase
    {
        public FieldSelectionViewModel ViewModel { get; }

        public FieldSelectionWindow(IEnumerable<string> availableFields, IEnumerable<string> preSelectedFields)
        {
            InitializeComponent();
            ViewModel = new FieldSelectionViewModel(availableFields, preSelectedFields);
            DataContext = ViewModel;
        }
    }
}
