using System.Collections.Generic;
using System.Threading.Tasks;

namespace ParquetViewer.Services
{
    public interface IFieldSelectionService
    {
        Task<List<string>?> ShowAsync(IEnumerable<string> availableFields, IEnumerable<string> preSelectedFields);
    }
}
