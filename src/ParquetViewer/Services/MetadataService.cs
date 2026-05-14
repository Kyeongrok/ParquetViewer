using ParquetViewer.Engine;
using ParquetViewer.Services;
using ParquetViewer.Windows;
using System.Windows;

namespace ParquetViewer.Services
{
    public class MetadataService : IMetadataService
    {
        private readonly Window _owner;

        public MetadataService(Window owner) => _owner = owner;

        public void Show(IParquetEngine engine)
        {
            var window = new MetadataWindow(engine) { Owner = _owner };
            window.Show();
        }
    }
}
