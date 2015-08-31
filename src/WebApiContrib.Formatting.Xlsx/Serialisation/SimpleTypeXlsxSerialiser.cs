using System;
using System.Collections;
using util = WebApiContrib.Formatting.Xlsx.FormatterUtils;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SimpleTypeXlsxSerialiser : IXlsxSerialiser
    {
        public bool IgnoreHeadersAndAttributes
        {
            get { return true; }
        }

        public bool CanSerialiseType(Type itemType)
        {
            return util.IsSimpleType(itemType);
        }

        public void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            // Can't convert IEnumerable<primitive> to IEnumerable<object>
            var values = (IEnumerable)value;

            foreach (var val in values)
            {
                document.AppendRow(new object[] { val });
            }
        }
    }
}
