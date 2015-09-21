using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// A collection of column information for an Excel document, keyed by field/property name.
    /// </summary>
    public class ExcelColumnInfoCollection : KeyedCollection<string, ExcelColumnInfo>
    {
        public ICollection<string> Keys
        {
            get
            {
                if (this.Dictionary != null)
                {
                    return this.Dictionary.Keys;
                }
                else
                {
                    return new Collection<string>(this.Select(this.GetKeyForItem).ToArray());
                }
            }
        }

        protected override string GetKeyForItem(ExcelColumnInfo item)
        {
            return item.PropertyName;
        }
    }
}
