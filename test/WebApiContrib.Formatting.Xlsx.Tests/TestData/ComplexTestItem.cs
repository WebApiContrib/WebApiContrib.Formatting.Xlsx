using System;
using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Tests.TestData
{
    [ExcelDocument("Complex test item")]
    public class ComplexTestItem
    {
        public string Value1 { get; set; }

        [ExcelColumn(Order = 2)]
        public DateTime Value2 { get; set; }

        [ExcelColumn(Header = "Header 3", Order = 1)]
        public bool Value3 { get; set; }

        [ExcelColumn(Header = "Header 4", Order = -2, NumberFormat = "???.???")]
        public double Value4 { get; set; }

        [ExcelColumn(Header = "Header 5")]
        public TestEnum Value5 { get; set; }

        [ExcelColumn(Ignore = true)]
        public string Value6 { get; set; }
    }

    public enum TestEnum
    {
        First,
        Second
    }
}
