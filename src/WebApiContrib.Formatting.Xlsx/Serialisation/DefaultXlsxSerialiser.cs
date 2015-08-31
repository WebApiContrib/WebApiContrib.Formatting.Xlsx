using System;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    class DefaultXlsxSerialiser : IXlsxSerialiser
    {
        public bool IgnoreHeadersAndAttributes
        {
            get { return true; }
        }

        public bool CanSerialiseType(Type itemType)
        {
            throw new NotImplementedException();
        }

        public void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            throw new NotImplementedException();
        }
    }
}
