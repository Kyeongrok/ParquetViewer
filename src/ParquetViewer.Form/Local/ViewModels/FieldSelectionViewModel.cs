using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ParquetViewer;
using ParquetViewer.Helpers;
using ParquetViewer.Resources;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace ParquetViewer.Form.Local.ViewModels
{
    public sealed partial class FieldSelectionViewModel : ObservableObject
    {
        private readonly List<FieldItem> _allFields;

        public ObservableCollection<FieldItem> DisplayFields { get; } = new();

        [ObservableProperty] private string _filterText = string.Empty;
        [ObservableProperty] private bool _loadAllFields = true;
        [ObservableProperty] private bool _rememberChoice;

        private bool _selectAll;
        public bool SelectAll
        {
            get => _selectAll;
            set
            {
                if (!SetProperty(ref _selectAll, value)) return;
                foreach (var item in DisplayFields.Where(f => f.IsEnabled))
                    item.IsSelected = value;
                OnPropertyChanged(nameof(SelectedFieldCountText));
                OnPropertyChanged(nameof(SelectAllCheckboxText));
            }
        }

        public bool LoadSelectedFields
        {
            get => !LoadAllFields;
            set => LoadAllFields = !value;
        }

        public string SelectedFieldCountText
        {
            get
            {
                var count = _allFields.Count(f => f.IsEnabled && f.IsSelected);
                return $"Select specific fields ({count} selected)";
            }
        }

        public string SelectAllCheckboxText
            => SelectAll
                ? Strings.DeselectAllCheckmarkTextFormat.Format(_allFields.Count(f => f.IsEnabled))
                : Strings.SelectAllCheckmarkTextFormat.Format(_allFields.Count(f => f.IsEnabled));

        public List<string>? ResultFields { get; private set; }

        public FieldSelectionViewModel(IEnumerable<string> availableFields, IEnumerable<string> preSelectedFields)
        {
            var allAvailable = availableFields.ToList();
            var preSelected = preSelectedFields.ToList();

            var duplicates = allAvailable
                .GroupBy(f => f.ToUpperInvariant())
                .Where(g => g.Count() > 1)
                .SelectMany(g => g)
                .ToHashSet(System.StringComparer.OrdinalIgnoreCase);

            _allFields = allAvailable
                .Select(f => new FieldItem(f, preSelected.Contains(f), !duplicates.Contains(f)))
                .ToList();

            RememberChoice = AppSettings.AlwaysSelectAllFields;
            LoadAllFields = preSelected.Count == 0;

            if (!_loadAllFields)
                SubscribeToFieldChanges();

            ApplyFilter(string.Empty);
        }

        partial void OnFilterTextChanged(string value) => ApplyFilter(value);

        partial void OnLoadAllFieldsChanged(bool value)
        {
            OnPropertyChanged(nameof(LoadSelectedFields));
            if (!value)
                SubscribeToFieldChanges();
        }

        private void ApplyFilter(string filter)
        {
            DisplayFields.Clear();

            var filtered = string.IsNullOrWhiteSpace(filter)
                ? _allFields
                : FilterByText(filter);

            foreach (var item in filtered)
                DisplayFields.Add(item);
        }

        private IEnumerable<FieldItem> FilterByText(string filter)
        {
            var parts = filter.Split(',');
            if (parts.Length == 1)
            {
                return _allFields.Where(f =>
                    f.Name.Contains(filter.Trim(), System.StringComparison.OrdinalIgnoreCase));
            }

            var exactNames = parts
                .Select(p => p.Trim(' ', '"', '\''))
                .Where(p => p.Length > 0)
                .ToHashSet(System.StringComparer.OrdinalIgnoreCase);
            return _allFields.Where(f => exactNames.Contains(f.Name));
        }

        private void SubscribeToFieldChanges()
        {
            foreach (var item in _allFields)
                item.PropertyChanged += (_, _) => OnPropertyChanged(nameof(SelectedFieldCountText));
        }

        [RelayCommand]
        private void ClearFilter() => FilterText = string.Empty;

        [RelayCommand]
        private void Confirm(Window window)
        {
            List<string> result;

            if (LoadAllFields)
            {
                AppSettings.AlwaysSelectAllFields = RememberChoice;
                result = _allFields.Select(f => f.Name).ToList();
            }
            else
            {
                result = _allFields.Where(f => f.IsEnabled && f.IsSelected).Select(f => f.Name).ToList();
                if (result.Count == 0)
                {
                    MessageBox.Show(
                        Errors.SelectAtLeastOneFieldErrorMessage,
                        Errors.SelectAtLeastOneFieldErrorTitle,
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AppSettings.AlwaysSelectAllFields = false;
            }

            ResultFields = result;
            window.DialogResult = true;
            window.Close();
        }

        [RelayCommand]
        private void Cancel(Window window)
        {
            ResultFields = null;
            window.Close();
        }
    }
}
