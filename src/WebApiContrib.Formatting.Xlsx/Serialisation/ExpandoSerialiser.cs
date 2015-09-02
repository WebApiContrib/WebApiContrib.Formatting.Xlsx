using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Custom serialiser for <c>ExpandoObject</c>.
    /// </summary>
    public class ExpandoSerialiser : IXlsxSerialiser
    {
        public bool IgnoreFormatting
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Returns true if value or item type are, or inherit from, <c>ExpandoObject</c>.
        /// </summary>
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType.IsAssignableFrom(typeof(ExpandoObject)) || itemType.IsAssignableFrom(typeof(ExpandoObject));
        }

        public void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            if (value.GetType().IsAssignableFrom(typeof(ExpandoObject))) {
                value = new[] { value };
            }
 
            var data = value as IEnumerable<object>;
            var first = data.FirstOrDefault();

            if (first == null) return;

            var members = GetDynamicMembers(first);

            if (members.Count() == 0) return;

            // Add member names as headers.
            document.AppendRow(members);

            foreach (var item in data)
            {
                var propertyValues = GetDynamicPropertyValues(item);
                var row = new List<object>();

                foreach (var member in members)
                {
                    row.Add(propertyValues[member]);
                }

                document.AppendRow(row);
            }
        }

        public IEnumerable<string> GetDynamicMembers(object item)
        {
            var provider = item as IDynamicMetaObjectProvider;
            var meta = provider.GetMetaObject(Expression.Constant(provider));

            return meta.GetDynamicMemberNames();
        }

        public IDictionary<string, object> GetDynamicPropertyValues(object item)
        {
            return (IDictionary<string, object>)item;
        }
    }
}
