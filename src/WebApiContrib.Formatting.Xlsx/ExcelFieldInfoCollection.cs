using System.Collections.ObjectModel;

namespace WebApiContrib.Formatting.Xlsx
{
    public class ExcelFieldInfoCollection : KeyedCollection<string, ExcelFieldInfo>
    {
        protected override string GetKeyForItem(ExcelFieldInfo item)
        {
            return item.PropertyName;
        }
    }
}
