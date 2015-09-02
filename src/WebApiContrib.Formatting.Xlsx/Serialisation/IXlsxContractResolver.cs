using System;
using System.Collections.Generic;
using System.Reflection;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// 
    /// </summary>
    public interface IXlsxContractResolver
    {
        /// <summary>
        /// Get the <c>ExcelFieldInfo</c> for all serialisable members of a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised.</param>
        ExcelFieldInfoCollection GetExcelFieldInfoCollection(Type itemType, IEnumerable<object> data);

        /// <summary>
        /// Get a list of all serialisable members for a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised.</param>
        IEnumerable<string> GetSerialisableMemberNames(Type itemType, IEnumerable<object> data);

        /// <summary>
        /// Get <c>PropertyInfo</c> for any serialisable properties in a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised.</param>
        IEnumerable<PropertyInfo> GetSerialisablePropertyInfo(Type itemType, IEnumerable<object> data);
    }
}
