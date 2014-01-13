using System.Runtime.Serialization;
using ExcelWebApi;

namespace ExcelWebApi.Tests.TestData
{
    public class ComplexTestItem
    {
        public string Value1 { get; set; }

        [Excel(Order = 2)]
        public string Value2 { get; set; }

        [Excel(Order = 1)]
        public string Value3 { get; set; }

        [Excel(Order = -2)]
        public string Value4 { get; set; }

        public string Value5 { get; set; }

        [Excel(Ignore = true)]
        public string Value6 { get; set; }
    }
}
