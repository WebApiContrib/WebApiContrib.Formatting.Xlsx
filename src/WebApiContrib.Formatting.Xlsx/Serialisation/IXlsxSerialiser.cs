using System;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public interface IXlsxSerialiser
    {
        bool IgnoreHeadersAndAttributes { get; }

        bool CanSerialiseType(Type itemType);
        void Serialise(Type itemType, object value, XlsxDocumentBuilder document);
    }
}
