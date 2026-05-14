using Microsoft.Win32;
using ParquetViewer.Controls;
using ParquetViewer.Engine;
using ParquetViewer.Helpers;
using ParquetViewer.Resources;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ParquetViewer.Windows
{
    public partial class MetadataWindow : WindowBase
    {
        private const string ThriftMetadataTabName = "Thrift Metadata";
        private const string ApacheArrowSchemaKey = "ARROW:schema";
        private const string PandasSchemaKey = "pandas";

        public MetadataWindow(IParquetEngine engine)
        {
            InitializeComponent();
            Loaded += async (_, _) => await LoadMetadataAsync(engine);
        }

        private async Task LoadMetadataAsync(IParquetEngine engine)
        {
            List<(string Name, string Content)> tabs;
            try
            {
                tabs = await Task.Run(() => BuildTabs(engine));
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(this,
                    Errors.MetadataReadErrorMessage + System.Environment.NewLine + ex,
                    Errors.MetadataReadErrorTitle,
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Close();
                return;
            }

            foreach (var (name, content) in tabs)
            {
                var textBox = new TextBox
                {
                    IsReadOnly = true,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = 12,
                    Text = content,
                    TextWrapping = TextWrapping.NoWrap,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    AcceptsReturn = true,
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(4),
                };
                MetadataTabControl.Items.Add(new TabItem { Header = name, Content = textBox });
            }

            LoadingPanel.Visibility = Visibility.Collapsed;
            MetadataTabControl.Visibility = Visibility.Visible;
        }

        private static List<(string Name, string Content)> BuildTabs(IParquetEngine engine)
        {
            var result = new List<(string, string)>();

            var json = ParquetMetadataAnalyzers.ThriftMetadataToJSON(engine, engine.RecordCount, engine.Fields.Count);
            result.Add((ThriftMetadataTabName, json));

            if (engine.CustomMetadata is not null)
            {
                foreach (var kv in engine.CustomMetadata)
                {
                    var value = kv.Key switch
                    {
                        PandasSchemaKey => ParquetMetadataAnalyzers.TryFormatJSON(kv.Value),
                        ApacheArrowSchemaKey => ParquetMetadataAnalyzers.ApacheArrowToJSON(kv.Value),
                        _ => ParquetMetadataAnalyzers.TryFormatJSON(kv.Value)
                    };
                    result.Add((kv.Key, value));
                }
            }

            return result;
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (MetadataTabControl.SelectedContent is not TextBox tb) return;
            try
            {
                Clipboard.SetText(tb.Text);
            }
            catch
            {
                if (MessageBox.Show(this,
                    Errors.CopyRawMetadataFailedErrorMessage,
                    Errors.CopyRawMetadataFailedErrorTitle,
                    MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    SaveToFile(tb.Text);
                }
            }
        }

        private void SaveToFile(string content)
        {
            var dialog = new SaveFileDialog { Filter = "JSON file (*.json)|*.json|Text file (*.txt)|*.txt" };
            if (dialog.ShowDialog(this) == true)
            {
                try { File.WriteAllText(dialog.FileName, content); }
                catch (System.Exception ex)
                {
                    MessageBox.Show(this, ex.Message, Errors.GenericErrorMessage,
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
    }
}
