using CommunityToolkit.Mvvm.ComponentModel;

namespace ParquetViewer.Form.Local.ViewModels
{
    public sealed partial class FieldItem : ObservableObject
    {
        public string Name { get; }
        public bool IsEnabled { get; }

        [ObservableProperty] private bool _isSelected;

        public FieldItem(string name, bool isSelected, bool isEnabled)
        {
            Name = name;
            _isSelected = isSelected;
            IsEnabled = isEnabled;
        }
    }
}
