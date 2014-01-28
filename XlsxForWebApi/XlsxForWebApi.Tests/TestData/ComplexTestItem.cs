using System;
using System.Runtime.Serialization;
using XlsxForWebApi;

namespace XlsxForWebApi.Tests.TestData
{
    [ExcelDocument("Complex test item")]
    public class ComplexTestItem
    {
        public enum TestEnum
	    {
	        First,
            Second
	    }

        public string Value1 { get; set; }

        [ExcelColumn(Order = 2)]
        public DateTime Value2 { get; set; }

        [ExcelColumn(Header = "Header 3", Order = 1)]
        public bool Value3 { get; set; }

        [ExcelColumn(Order = -2, NumberFormat="???.???")]
        public double Value4 { get; set; }

        public TestEnum Value5 { get; set; }

        [ExcelColumn(Ignore = true)]
        public string Value6 { get; set; }
    }
}
