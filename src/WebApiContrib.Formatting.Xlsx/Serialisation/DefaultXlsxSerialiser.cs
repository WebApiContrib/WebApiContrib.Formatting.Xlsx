using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.ModelBinding;
using WebApiContrib.Formatting.Xlsx.Attributes;
using util = WebApiContrib.Formatting.Xlsx.FormatterUtils;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class DefaultXlsxSerialiser : IXlsxSerialiser
    {
        public bool IgnoreHeadersAndAttributes
        {
            get { return false; }
        }

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return true;
        }

        public void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            var data = value as IEnumerable<object>;
            
            var serialisableMembers = util.GetMemberNames(itemType);

            var properties = GetProperties(itemType);


            var fields = new List<string>();
            var fieldInfo = new ExcelFieldInfoCollection();

            // Instantiate field names and fieldInfo lists with serialisable members.
            foreach (var field in serialisableMembers)
            {
                var propName = field;
                var prop = properties.FirstOrDefault(p => p.Name == propName);

                if (prop == null) continue;

                fields.Add(field);
                fieldInfo.Add(new ExcelFieldInfo(field, util.GetAttribute<ExcelColumnAttribute>(prop)));
            }

            if (fields.Count == 0) return;


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

        public static IEnumerable<PropertyInfo> GetProperties(Type itemType)
        {
            return (from p in itemType.GetProperties()
                    where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                    select p).ToList();
        }

        public static object FormatCellValue(object cellValue, ExcelFieldInfo info)
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
        private static object GetFieldOrPropertyValue(object rowObject, string name)
        {
            var rowValue = util.GetFieldOrPropertyValue(rowObject, name);

            if (IsExcelSupportedType(rowValue)) return rowValue;

            return rowValue == null || DBNull.Value.Equals(rowValue)
                ? string.Empty
                : rowValue.ToString();
        }

        public static Boolean IsExcelSupportedType(object expression)
        {
            return expression is String || expression is Int16 || expression is Int32 || expression is Int64 || expression is Decimal || expression is Single || expression is Double || expression is DateTime;
        }

    }
}
