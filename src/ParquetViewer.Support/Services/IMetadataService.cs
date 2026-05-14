using ParquetViewer.Engine;

namespace ParquetViewer.Services
{
    public interface IMetadataService
    {
        void Show(IParquetEngine engine);
    }
}
