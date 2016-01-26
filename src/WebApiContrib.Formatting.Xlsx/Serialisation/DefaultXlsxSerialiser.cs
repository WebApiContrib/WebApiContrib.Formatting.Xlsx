using System;
using System.Collections.Generic;
using System.Linq;
using util = WebApiContrib.Formatting.Xlsx.FormatterUtils;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Serialises public, parameterless properties of a class, taking account of any custom attributes.
    /// </summary>
    public class DefaultXlsxSerialiser : IXlsxSerialiser
    {
        /// <summary>
        /// Default resolver determining which columns are generated from a type and how they are formatted.
        /// </summary>
        public IColumnResolver Resolver { get; set; }

        public DefaultXlsxSerialiser() : this(new DefaultColumnResolver()) { }

        public DefaultXlsxSerialiser(IColumnResolver resolver)
        {
            Resolver = resolver;
        }

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
            var columnInfo = Resolver.GetExcelColumnInfo(itemType, data);

            if (columnInfo.Count == 0) return;

            var columns = columnInfo.Keys.ToList();

            // Add header row
            document.AppendRow((from col in columnInfo select col.Header).ToList());

            // Output each row of data
            if (data != null)
            {
                foreach (var dataObject in data)
                {
                    var row = new List<object>();

                    for (int i = 0; i <= columns.Count - 1; i++)
                    {
                        var cellValue = GetFieldOrPropertyValue(dataObject, columns[i]);
                        var info = columnInfo[i];

                        row.Add(FormatCellValue(cellValue, info));
                    }

                    document.AppendRow(row.ToArray());
                }
            }

            // Enforce any attributes on columns.
            for (int i = 1; i <= columns.Count; i++)
            {
                if (!string.IsNullOrEmpty(columnInfo[i - 1].ExcelNumberFormat))
                {
                    document.FormatColumn(i, columnInfo[i - 1].ExcelNumberFormat);
                }
            }
        }

        /// <summary>
        /// Format a value before serialisation based on its attributes.
        /// </summary>
        /// <param name="cellValue">The value about to be serialised.</param>
        /// <param name="info">Formatting information for this cell based on attributes.</param>
        protected virtual object FormatCellValue(object cellValue, ExcelColumnInfo info)
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

            if (rowValue is DateTimeOffset)
                return ConvertFromDateTimeOffset((DateTimeOffset)rowValue);

            else if (IsExcelSupportedType(rowValue))
                return rowValue;

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

        // Taken from http://msdn.microsoft.com/en-us/library/bb546101.aspx
        private static DateTime ConvertFromDateTimeOffset(DateTimeOffset dateTime)
        {
            if (dateTime.Offset.Equals(TimeSpan.Zero))
                return dateTime.UtcDateTime;
            else if (dateTime.Offset.Equals(TimeZoneInfo.Local.GetUtcOffset(dateTime.DateTime)))
                return DateTime.SpecifyKind(dateTime.DateTime, DateTimeKind.Local);
            else
                return dateTime.DateTime;
        }
    }
}
