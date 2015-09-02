using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class PerRequestContractResolver : DefaultXlsxContractResolver
    {
        public const string DEFAULT_KEY = "XlsxSerialisableProperties";

        public string HttpContextItemKey { get; set; }

        public PerRequestContractResolver(string httpContextItemKey = DEFAULT_KEY)
        {
            HttpContextItemKey = httpContextItemKey;
        }

        public override IEnumerable<string> GetSerialisableMemberNames(Type itemType, IEnumerable<object> data)
        {
            var defaultMemberNames = base.GetSerialisableMemberNames(itemType, data);
            var requestProperties = (IEnumerable<string>) System.Web.HttpContext.Current.Items[HttpContextItemKey];

            return requestProperties.Where(name => defaultMemberNames.Contains(name));
        }
    }
}
