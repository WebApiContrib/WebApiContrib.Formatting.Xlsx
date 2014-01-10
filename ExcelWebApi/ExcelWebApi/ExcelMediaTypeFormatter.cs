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

namespace ExcelWebApi
{

	/// <summary>
	/// Class used to send an Excel file to the response.
	/// </summary>
	/// <remarks>Relies upon <c>DataMember</c> attributes to determine serialization.</remarks>
    public class ExcelMediaTypeFormatter : MediaTypeFormatter
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

		/// <summary>
		/// Height for cells. (Default if null.)
		/// </summary>
		protected double? CellHeight { get; set; }

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
		public ExcelMediaTypeFormatter(bool autoFit = true, bool autoFilter = false, bool freezeHeader = false, float headerHeight = 0, float cellHeight = 0, Action<ExcelStyle> cellStyle = null, Action<ExcelStyle> headerStyle = null)
		{
			SupportedMediaTypes.Clear();
			SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
			SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.ms-excel"));

			AutoFit = autoFit;
			AutoFilter = autoFilter;
			FreezeHeader = freezeHeader;
			HeaderHeight = headerHeight;
			CellHeight = cellHeight;
			CellStyle = cellStyle;
			HeaderStyle = headerStyle;
		}

		#endregion

		#region Methods

		public override void SetDefaultContentHeaders(Type type, HttpContentHeaders headers, MediaTypeHeaderValue mediaType)
		{
			// Get the raw URI and strip out query string.
			string rawUri = System.Web.HttpContext.Current.Request.RawUrl;

			int queryStringIndex = rawUri.IndexOf('?');
			if (queryStringIndex > -1)
			{
				rawUri = rawUri.Substring(0, queryStringIndex);
			}

			// Get filename and add extension if none provided.
			string fileName = System.Web.VirtualPathUtility.GetFileName(rawUri);
            // ReSharper disable once PossibleNullReferenceException
			if (fileName.IndexOf(".") == -1)
			{
				fileName += ".xlsx";
			}

			// Set content disposition with a suggested filename.
			//headers.ContentDisposition = new ContentDispositionHeaderValue("inline") { FileName = fileName };

			base.SetDefaultContentHeaders(type, headers, mediaType);
		}

		[SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
		public override Task WriteToStreamAsync(Type type, object value, System.IO.Stream writeStream, System.Net.Http.HttpContent content, System.Net.TransportContext transportContext)
		{
			var data = (IEnumerable<object>)value;

			// Create a worksheet
			var package = new ExcelPackage();
			package.Workbook.Worksheets.Add("Data");
			var worksheet = package.Workbook.Worksheets[1];
			worksheet.Name = "Data";

			// Default cell styles
			if (CellStyle != null)
				CellStyle(worksheet.Cells.Style);
			if (HeaderStyle != null)
				HeaderStyle(worksheet.Row(1).Style);
			if (CellHeight.HasValue)
				worksheet.DefaultRowHeight = CellHeight.Value;
			if (HeaderHeight.HasValue)
				worksheet.Row(1).Height = HeaderHeight.Value;
			if (FreezeHeader)
				worksheet.View.FreezePanes(2, 1);

			int rowCount = 0;

			var fields = new List<string>();
			var fieldInfo = new ExcelFieldInfoCollection();

			// Use all public, parameterless, readable properties of inner type.
            var itemType = FormatterUtils.GetEnumerableItemType(value);
			if (itemType == null) throw new ArgumentException("Only IEnumerable<T> values can be deserialised using the Excel formatter.");

			var serializableMembers = FormatterUtils.GetDataMemberNames(itemType);

			var metadata = ModelMetadataProviders.Current.GetMetadataForType(null, itemType);
            
			var properties = (from p in itemType.GetProperties()
                              where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                              select p).ToList();

            foreach (var field in serializableMembers)
			{
				var propName = field;
                var prop = properties.FirstOrDefault(p => p.Name == propName);

			    if (prop == null) continue;

			    fields.Add(field);
			    fieldInfo.Add(new ExcelFieldInfo(field, FormatterUtils.GetAttribute<ExcelAttribute>(prop)));
			}

			if (metadata != null && metadata.Properties != null)
			{
				foreach (var modelProp in metadata.Properties)
				{
					var propertyName = modelProp.PropertyName;

					if (!fieldInfo.Contains(propertyName)) continue;

					fieldInfo[propertyName].Header = modelProp.DisplayName ?? propertyName;
					fieldInfo[propertyName].FormatString = modelProp.DisplayFormatString;
				}
			}

			if (fields.Count <= 0) return Task.Factory.StartNew(() => package.SaveAs(writeStream));

			AppendRow((from f in fieldInfo select f.Header).ToList(), worksheet, ref rowCount);

			// Output each row of data
			if (data != null && data.FirstOrDefault() != null) {
				foreach (var dataObject in data)
				{
					var row = new List<object>();

					for (int i = 0; i <= fields.Count - 1; i++) {
						var cellValue = GetFieldOrPropertyValue(dataObject, fields[i]);
						if (!string.IsNullOrWhiteSpace(fieldInfo[i].FormatString) & string.IsNullOrEmpty(fieldInfo[i].ExcelNumberFormat)) {
							row.Add(string.Format(fieldInfo[i].FormatString, cellValue));
						} else {
							row.Add(cellValue);
						}
					}

					//For Each field In Fields
					//	row.Add(GetPropertyValue(dataObject, field))
					//Next
					AppendRow(row.ToArray(), worksheet, ref rowCount);
				}
			}

			// Enforce any attributes on columns.
			for (int i = 1; i <= fields.Count; i++) {
				if (!string.IsNullOrEmpty(fieldInfo[i - 1].ExcelNumberFormat)) {
					worksheet.Cells[2, i, rowCount, i].Style.Numberformat.Format = fieldInfo[i - 1].ExcelNumberFormat;
				}
			}

			dynamic cells = worksheet.Cells[worksheet.Dimension.Address];

			cells.AutoFilter = AutoFilter;
			if (AutoFit)
				cells.AutoFitColumns();

			return Task.Factory.StartNew(() => package.SaveAs(writeStream));
		}

		/// <summary>
		/// Get a property value from an object.
		/// </summary>
		/// <param name="rowObject">The object whose property we want.</param>
		/// <param name="name">The name of the property we want.</param>
		private static object GetFieldOrPropertyValue(object rowObject, string name)
		{
            var rowValue = FormatterUtils.GetFieldOrPropertyValue(rowObject, name);

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
				// God, unary-based indexes should not mix with zero-based. :(
				worksheet.Cells[rowCount, i].Value = enumerable.ElementAt(i - 1);
			}
        }

        public override bool CanWriteType(Type type)
        {
            return type.GetInterface(typeof(IEnumerable).FullName) != null && typeof(IEnumerable).IsAssignableFrom(type);
        }

        public override bool CanReadType(Type type)
        {
            return type.GetInterface(typeof(IEnumerable).FullName) != null && typeof(IEnumerable).IsAssignableFrom(type);
        }

		#endregion

    }
}
