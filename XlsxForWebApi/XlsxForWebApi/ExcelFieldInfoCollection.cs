using System.Collections.ObjectModel;

namespace XlsxForWebApi
{
    public class ExcelFieldInfoCollection : KeyedCollection<string, ExcelFieldInfo>
    {
        protected override string GetKeyForItem(ExcelFieldInfo item)
        {
            return item.PropertyName;
        }
    }
}
