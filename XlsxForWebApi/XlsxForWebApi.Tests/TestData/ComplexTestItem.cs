using System.Runtime.Serialization;
using XlsxForWebApi;

namespace XlsxForWebApi.Tests.TestData
{
    [ExcelDocument("Complex test item")]
    public class ComplexTestItem
    {
        public string Value1 { get; set; }

        [ExcelColumn(Order = 2)]
        public string Value2 { get; set; }

        [ExcelColumn(Order = 1)]
        public string Value3 { get; set; }

        [ExcelColumn(Order = -2)]
        public string Value4 { get; set; }

        public string Value5 { get; set; }

        [ExcelColumn(Ignore = true)]
        public string Value6 { get; set; }
    }
}
