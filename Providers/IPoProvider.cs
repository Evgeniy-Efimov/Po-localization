using System.Collections.Generic;
using System.Text;

namespace LocalizePo.Providers
{
    public interface IPoProvider
    {
        IEnumerable<PoObject<TModel>> ReadFile<TModel>(string filePath) where TModel : new();
        void ExportToFile<TModel>(IEnumerable<PoObject<TModel>> data, string filePath, Encoding encoding) where TModel : new();
    }
}
