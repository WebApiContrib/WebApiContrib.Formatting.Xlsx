using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Permissions;
using System.Threading.Tasks;
using System.Web.ModelBinding;
using util = WebApiContrib.Formatting.Xlsx.FormatterUtils;

namespace WebApiContrib.Formatting.Xlsx
{

    /// <summary>
    /// Class used to send an Excel file to the response.
    /// </summary>
    public class XlsxMediaTypeFormatter : MediaTypeFormatter
    {

        #region Properties

        /// <summary>
        /// An action method that can be used to set the default cell style.
        /// </summary>
        protected Action<ExcelStyle> CellStyle { get; set; }

        /// <summary>
        /// An action method that can be used to set the default header row style.
        /// </summary>
        protected Action<ExcelStyle> HeaderStyle { get; set; }

        /// <summary>
        /// True if columns should be auto-fit to the cell contents after writing.
        /// </summary>
        protected bool AutoFit { get; set; }

        /// <summary>
        /// True if an auto-filter should be enabled for the data.
        /// </summary>
        protected bool AutoFilter { get; set; }

        /// <summary>
        /// True if the header row should be frozen.
        /// </summary>
        protected bool FreezeHeader { get; set; }

        /// <summary>
        /// Height for header row. (Default if null.)
        /// </summary>
        protected double? HeaderHeight { get; set; }
        
        #endregion

        #region Constructor

        /// <summary>
        /// Create a new ExcelMediaTypeFormatter.
        /// </summary>
        /// <param name="autoFit">True if the formatter should autofit columns after writing the data. (Default true.)</param>
        /// <param name="autoFilter">True if an autofilter should be applied to the worksheet. (Default false.)</param>
        /// <param name="freezeHeader">True if the header row should be frozen. (Default false.)</param>
        /// <param name="headerHeight">Height of the header row.</param>
        /// <param name="cellHeight">Height of each row of data.</param>
        /// <param name="cellStyle">An action method that modifies the worksheet cell style.</param>
        /// <param name="headerStyle">An action method that modifies the cell style of the first (header) row in the worksheet.</param>
        public XlsxMediaTypeFormatter(bool autoFit = true, bool autoFilter = false, bool freezeHeader = false, double headerHeight = 0, Action<ExcelStyle> cellStyle = null, Action<ExcelStyle> headerStyle = null)
        {
            SupportedMediaTypes.Clear();
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.ms-excel"));

            AutoFit = autoFit;
            AutoFilter = autoFilter;
            FreezeHeader = freezeHeader;
            HeaderHeight = headerHeight;
            CellStyle = cellStyle;
            HeaderStyle = headerStyle;
        }

        #endregion

        #region Methods

        public override void SetDefaultContentHeaders(Type type, HttpContentHeaders headers, MediaTypeHeaderValue mediaType)
        {
            // Get the raw request URI.
            string rawUri = System.Web.HttpContext.Current.Request.RawUrl;

            // Remove query string if present.
            int queryStringIndex = rawUri.IndexOf('?');
            if (queryStringIndex > -1)
            {
                rawUri = rawUri.Substring(0, queryStringIndex);
            }

            string fileName;

            // Look for ExcelDocumentAttribute on class.
            var itemType = util.GetEnumerableItemType(type);
            var excelDocumentAttribute = util.GetAttribute<ExcelDocumentAttribute>(itemType ?? type);

            if (excelDocumentAttribute != null && !string.IsNullOrEmpty(excelDocumentAttribute.FileName))
            {
                // If attribute exists with file name defined, use that.
                fileName = excelDocumentAttribute.FileName;
            }
            else
            {
                // Otherwise, use either the URL file name component or just "data".
                fileName = System.Web.VirtualPathUtility.GetFileName(rawUri) ?? "data";
            }

            // Add XLSX extension if not present.
            if (!fileName.EndsWith("xlsx", StringComparison.CurrentCultureIgnoreCase)) fileName += ".xlsx";

            // Set content disposition to use this file name.
            headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
                                             { FileName = fileName };

            base.SetDefaultContentHeaders(type, headers, mediaType);
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override Task WriteToStreamAsync(Type type, object value, System.IO.Stream writeStream, System.Net.Http.HttpContent content, System.Net.TransportContext transportContext)
        {
            // Create a worksheet
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Data");
            var worksheet = package.Workbook.Worksheets[1];

            int rowCount = 0;
            var valueType = value.GetType();

            // Apply cell styles.
            if (CellStyle != null) CellStyle(worksheet.Cells.Style);

            // Get the item type.
            var itemType = (util.IsSimpleType(valueType))
                ? null
                : util.GetEnumerableItemType(valueType);

            // If a single object was passed, treat it like a list with one value.
            if (itemType == null)
            {
                itemType = valueType;
                value = new object[] { value };
            }

            // Enumerations of primitive types are also handled separately, since they have
            // no properties to serialise (and thus, no headers or attributes to consider).
            if (util.IsSimpleType(itemType))
            {
                // Can't convert IEnumerable<primitive> to IEnumerable<object>
                var values = (IEnumerable) value;

                foreach (var val in values)
                {
                    AppendRow(new object[] { val }, worksheet, ref rowCount);
                }

                // Autofit cells if specified.
                if (AutoFit) worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Save and done.
                return Task.Factory.StartNew(() => package.SaveAs(writeStream));
            }

            var data = value as IEnumerable<object>;

            // What remains is an enumeration of object types.
            var serialisableMembers = util.GetMemberNames(itemType);

            var metadata = ModelMetadataProviders.Current.GetMetadataForType(null, itemType);
            
            var properties = (from p in itemType.GetProperties()
                              where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                              select p).ToList();


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
            
            if (fields.Count == 0) return Task.Factory.StartNew(() => package.SaveAs(writeStream));

            // Add header row
            AppendRow((from f in fieldInfo select f.Header).ToList(), worksheet, ref rowCount);

            // Output each row of data
            if (data != null && data.FirstOrDefault() != null) {
                foreach (var dataObject in data)
                {
                    var row = new List<object>();

                    for (int i = 0; i <= fields.Count - 1; i++) {
                        var cellValue = GetFieldOrPropertyValue(dataObject, fields[i]);
                        var info = fieldInfo[i];

                        // Boolean transformations.
                        if (info.ExcelAttribute != null && info.ExcelAttribute.TrueValue != null && cellValue.Equals("True"))
                            row.Add(info.ExcelAttribute.TrueValue);

                        else if (info.ExcelAttribute != null && info.ExcelAttribute.FalseValue != null && cellValue.Equals("False"))
                            row.Add(info.ExcelAttribute.FalseValue);

                        else if (!string.IsNullOrWhiteSpace(info.FormatString) & string.IsNullOrEmpty(info.ExcelNumberFormat))
                            row.Add(string.Format(info.FormatString, cellValue));

                        else
                            row.Add(cellValue);
                    }

                    AppendRow(row.ToArray(), worksheet, ref rowCount);
                }
            }

            // Enforce any attributes on columns.
            for (int i = 1; i <= fields.Count; i++) {
                if (!string.IsNullOrEmpty(fieldInfo[i - 1].ExcelNumberFormat)) {
                    worksheet.Cells[2, i, rowCount, i].Style.Numberformat.Format = fieldInfo[i - 1].ExcelNumberFormat;
                }
            }

            // Header cell styles
            if (HeaderStyle != null) HeaderStyle(worksheet.Row(1).Style);
            if (FreezeHeader) worksheet.View.FreezePanes(2, 1);

            var cells = worksheet.Cells[worksheet.Dimension.Address];

            // Add autofilter and fit to max column width (if requested).
            cells.AutoFilter = AutoFilter;
            if (AutoFit) cells.AutoFitColumns();

            // Set header row where specified.
            if (HeaderHeight.HasValue)
            {
                worksheet.Row(1).Height = HeaderHeight.Value;
                worksheet.Row(1).CustomHeight = true;
            }

            return Task.Factory.StartNew(() => package.SaveAs(writeStream));
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

        /// <summary>
        /// Append a row to the <c>StringBuilder</c> containing the CSV data.
        /// </summary>
        /// <param name="row">The row to append to this instance.</param>
        /// <param name="worksheet">The worksheet to append this row to.</param>
        /// <param name="rowCount">The number of rows appended so far.</param>
        private void AppendRow(IEnumerable<object> row, ExcelWorksheet worksheet, ref int rowCount)
        {
            rowCount++;

            var enumerable = row as IList<object> ?? row.ToList();
            for (var i = 1; i <= enumerable.Count(); i++)
            {
                // Unary-based indexes should not mix with zero-based. :(
                worksheet.Cells[rowCount, i].Value = enumerable.ElementAt(i - 1);
            }
        }

        public override bool CanWriteType(Type type)
        {
            // Should be able to serialise any type.
            return true;
        }

        public override bool CanReadType(Type type)
        {
            // Not yet possible; see issue page to track progress:
            // https://github.com/jordangray/xlsx-for-web-api/issues/5
            return false;
        }

        #endregion

    }
}
