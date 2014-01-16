namespace XlsxForWebApi
{
    public class ExcelFieldInfo
    {
        public string PropertyName { get; set; }
        public ExcelAttribute ExcelAttribute { get; set; }
        public string FormatString { get; set; }
        public string Header { get; set; }

        public string ExcelNumberFormat
        {
            get { return ExcelAttribute != null ? ExcelAttribute.NumberFormat : null; }
        }

        public bool IsExcelHeaderDefined
        {
            get { return ExcelAttribute != null && ExcelAttribute.Header != null; }
        }

        public ExcelFieldInfo(string propertyName, ExcelAttribute excelAttribute = null, string formatString = null)
        {
            PropertyName = propertyName;
            ExcelAttribute = excelAttribute;
            FormatString = formatString;
            Header = IsExcelHeaderDefined ? ExcelAttribute.Header : propertyName;
        }
    }
}
