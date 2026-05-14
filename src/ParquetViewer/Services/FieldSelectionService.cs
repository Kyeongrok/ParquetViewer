using ParquetViewer.Services;
using ParquetViewer.Windows;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;

namespace ParquetViewer.Services
{
    public class FieldSelectionService : IFieldSelectionService
    {
        private readonly Window _owner;

        public FieldSelectionService(Window owner)
        {
            _owner = owner;
        }

        public Task<List<string>?> ShowAsync(IEnumerable<string> availableFields, IEnumerable<string> preSelectedFields)
        {
            var window = new FieldSelectionWindow(availableFields, preSelectedFields)
            {
                Owner = _owner
            };
            window.ShowDialog();
            return Task.FromResult(window.ViewModel.ResultFields);
        }
    }
}
