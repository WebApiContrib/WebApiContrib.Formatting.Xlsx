using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Utils;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Resolves the properties whitelisted by name in an item (default <c>XlsxSerialisableProperties</c>) of the
    /// current request's <c>HttpContext</c>, optionally respecting the whitelist order.
    /// </summary>
    public class PerRequestColumnResolver : DefaultColumnResolver
    {
        public const string DEFAULT_KEY = "XlsxSerialisableProperties";

        /// <summary>
        /// The key to look up in the <c>HttpContext.Current.Items</c> collection.
        /// </summary>
        public string HttpContextItemKey { get; set; }

        /// <summary>
        /// Override output order with the order that properties are defined in.
        /// </summary>
        public bool UseCustomOrder { get; set; }

        public PerRequestColumnResolver(string httpContextItemKey = DEFAULT_KEY, bool useCustomOrder = false)
        {
            HttpContextItemKey = httpContextItemKey;
            UseCustomOrder = useCustomOrder;
        }

        /// <summary>
        /// Get member names from <c>System.Web.HttpContext.Current.Items[HttpContextItemKey]</c> if key was defined,
        /// or default member names from base class implementation if not.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        /// <remarks>Any names specified in the per-request dictionary that aren't serialisable will be
        /// discarded.</remarks>
        public override IEnumerable<string> GetSerialisableMemberNames(Type itemType, IEnumerable<object> data)
        {
            var defaultMemberNames = base.GetSerialisableMemberNames(itemType, data);
            var httpContextItems = HttpContextFactory.Current.Items;

            if (!httpContextItems.Contains(HttpContextItemKey)) return defaultMemberNames;

            var itemValue = httpContextItems[HttpContextItemKey];

            if (!(itemValue is IEnumerable<string>)) return defaultMemberNames;

            var requestProperties = (IEnumerable<string>)itemValue;

            return UseCustomOrder
                ? requestProperties.Where(name => defaultMemberNames.Contains(name))
                : defaultMemberNames.Where(name => requestProperties.Contains(name));
        }
    }
}
