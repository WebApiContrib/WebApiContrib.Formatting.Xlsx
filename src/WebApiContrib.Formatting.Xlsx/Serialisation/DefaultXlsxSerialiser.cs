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
    public class DefaultXlsxSerialiser : IXlsxSerialiser
    {
        public virtual bool IgnoreFormatting
        {
            get { return false; }
        }

        public virtual bool CanSerialiseType(Type valueType, Type itemType)
        {
            return true;
        }

        public virtual void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            var data = value as IEnumerable<object>;
            
            var fieldInfo = GetFieldInfo(itemType, data);
            var fields = fieldInfo.Keys.ToList();

            if (fields.Count == 0) return;


            // Add header row
            document.AppendRow((from f in fieldInfo select f.Header).ToList());

            // Output each row of data
            if (data != null && data.FirstOrDefault() != null)
            {
                foreach (var dataObject in data)
                {
                    var row = new List<object>();

                    for (int i = 0; i <= fields.Count - 1; i++)
                    {
                        var cellValue = GetFieldOrPropertyValue(dataObject, fields[i]);
                        var info = fieldInfo[i];

                        row.Add(FormatCellValue(cellValue, info));
                    }

                    document.AppendRow(row.ToArray());
                }
            }
            

            // Enforce any attributes on columns.
            for (int i = 1; i <= fields.Count; i++)
            {
                if (!string.IsNullOrEmpty(fieldInfo[i - 1].ExcelNumberFormat))
                {
                    document.FormatColumn(i, fieldInfo[i - 1].ExcelNumberFormat);
                }
            }
        }

        /// <summary>
        /// Get the <c>ExcelFieldInfo</c> for all members of a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        protected virtual ExcelFieldInfoCollection GetFieldInfo(Type itemType, IEnumerable<object> data)
        {
            var fields = GetSerialisableMemberNames(itemType, data);
            var properties = GetSerialisablePropertyInfo(itemType, data);

            var fieldInfo = new ExcelFieldInfoCollection();

            // Instantiate field names and fieldInfo lists with serialisable members.
            foreach (var field in fields)
            {
                var propName = field;
                var prop = properties.FirstOrDefault(p => p.Name == propName);

                if (prop == null) continue;

                fieldInfo.Add(new ExcelFieldInfo(field, util.GetAttribute<ExcelColumnAttribute>(prop)));
            }

            PopulateFieldInfoFromMetadata(fieldInfo, itemType, data);

            return fieldInfo;
        }

        /// <summary>
        /// Get a list of all serialisable members for a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        protected virtual IEnumerable<string> GetSerialisableMemberNames(Type itemType, IEnumerable<object> data)
        {
            return util.GetMemberNames(itemType);
        }

        /// <summary>
        /// Get <c>PropertyInfo</c> for all public properties with parameterless get methods in a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        protected virtual IEnumerable<PropertyInfo> GetSerialisablePropertyInfo(Type itemType, IEnumerable<object> data)
        {
            return (from p in itemType.GetProperties()
                    where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                    select p).ToList();
        }

        /// <summary>
        /// Populate missing or incomplete properties from model metadata.
        /// </summary>
        /// <param name="fieldInfo">The <c>ExcelFieldInfoCollection</c> to populate.</param>
        /// <param name="itemType">The type of item whose metadata this is being populated from.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        protected virtual void PopulateFieldInfoFromMetadata(ExcelFieldInfoCollection fieldInfo,
                                                          Type itemType,
                                                          IEnumerable<object> data)
        {
            // Populate missing attribute information from metadata.
            var metadata = ModelMetadataProviders.Current.GetMetadataForType(null, itemType);

            if (metadata != null && metadata.Properties != null)
            {
                foreach (var modelProp in metadata.Properties)
                {
                    var propertyName = modelProp.PropertyName;

                    if (!fieldInfo.Contains(propertyName)) continue;

                    var field = fieldInfo[propertyName];
                    var attribute = field.ExcelAttribute;

                    if (!field.IsExcelHeaderDefined)
                        field.Header = modelProp.DisplayName ?? propertyName;

                    if (attribute != null && attribute.UseDisplayFormatString)
                        field.FormatString = modelProp.DisplayFormatString;
                }
            }
        }

        /// <summary>
        /// Format a value before serialisation based on its attributes.
        /// </summary>
        /// <param name="cellValue">The value about to be serialised.</param>
        /// <param name="info">Formatting information for this cell based on attributes.</param>
        protected virtual object FormatCellValue(object cellValue, ExcelFieldInfo info)
        {
            // Boolean transformations.
            if (info.ExcelAttribute != null && info.ExcelAttribute.TrueValue != null && cellValue.Equals("True"))
                return info.ExcelAttribute.TrueValue;

            else if (info.ExcelAttribute != null && info.ExcelAttribute.FalseValue != null && cellValue.Equals("False"))
                return info.ExcelAttribute.FalseValue;

            else if (!string.IsNullOrWhiteSpace(info.FormatString) & string.IsNullOrEmpty(info.ExcelNumberFormat))
                return string.Format(info.FormatString, cellValue);

            else
                return cellValue;
        }

        /// <summary>
        /// Get a property value from an object.
        /// </summary>
        /// <param name="rowObject">The object whose property we want.</param>
        /// <param name="name">The name of the property we want.</param>
        protected virtual object GetFieldOrPropertyValue(object rowObject, string name)
        {
            var rowValue = util.GetFieldOrPropertyValue(rowObject, name);

            if (IsExcelSupportedType(rowValue)) return rowValue;

            return rowValue == null || DBNull.Value.Equals(rowValue)
                ? string.Empty
                : rowValue.ToString();
        }

        /// <summary>
        /// Determines if a particular value can be represented natively in Excel without being cast to a string.
        /// </summary>
        /// <param name="expression">The value to test.</param>
        protected static bool IsExcelSupportedType(object expression)
        {
            return expression is String 
                || expression is Int16
                || expression is Int32
                || expression is Int64
                || expression is Decimal
                || expression is Single
                || expression is Double
                || expression is DateTime;
        }

    }
}
