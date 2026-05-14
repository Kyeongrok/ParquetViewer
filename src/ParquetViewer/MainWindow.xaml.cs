using ParquetViewer.Analytics;
using ParquetViewer.Controls;
using ParquetViewer.Main.Local.ViewModels;
using System.Windows;

namespace ParquetViewer
{
    public partial class MainWindow : WindowBase
    {
        private readonly MainViewModel _viewModel;

        public MainWindow(string? initialPath = null)
        {
            InitializeComponent();

            var service = new Services.FieldSelectionService(this);
            _viewModel = new MainViewModel(service);
            DataContext = _viewModel;

            AllowDrop = true;
            Drop += OnDrop;
            DragEnter += OnDragEnter;

            if (initialPath is not null)
                Loaded += async (_, _) => await _viewModel.OpenNewFileOrFolderAsync(initialPath);

            Loaded += (_, _) =>
            {
                App.GetUserConsentToGatherAnalytics();
                App.AskUserIfTheyWantToSwitchToDarkMode();
            };
        }

        private async void OnDrop(object sender, DragEventArgs e)
        {
            var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files?.Length > 0)
            {
                MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.DragDrop);
                await _viewModel.OpenNewFileOrFolderAsync(files[0]);
            }
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effects = DragDropEffects.Copy;
        }
    }
}
