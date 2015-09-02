using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.ModelBinding;
using WebApiContrib.Formatting.Xlsx.Attributes;
using util = WebApiContrib.Formatting.Xlsx.FormatterUtils;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Serialises public, parameterless properties of a class, taking account of any custom attributes.
    /// </summary>
    public class CustomisableSerialiser : DefaultXlsxSerialiser
    {
        public override bool CanSerialiseType(Type valueType, Type itemType)
        {
            return true;
        }
    }
}
