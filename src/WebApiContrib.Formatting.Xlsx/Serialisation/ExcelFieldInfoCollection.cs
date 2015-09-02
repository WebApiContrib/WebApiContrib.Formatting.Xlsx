using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelFieldInfoCollection : KeyedCollection<string, ExcelFieldInfo>
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

        protected override string GetKeyForItem(ExcelFieldInfo item)
        {
            return item.PropertyName;
        }
    }
}
