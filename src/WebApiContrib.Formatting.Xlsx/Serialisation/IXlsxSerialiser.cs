using System;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public interface IXlsxSerialiser
    {
        bool IgnoreHeadersAndAttributes { get; }

        bool CanSerialiseType(Type valueType, Type itemType);

        void Serialise(Type itemType, object value, XlsxDocumentBuilder document);
    }
}
