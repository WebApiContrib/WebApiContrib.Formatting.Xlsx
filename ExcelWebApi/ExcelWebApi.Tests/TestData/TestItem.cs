using System.Runtime.Serialization;

namespace ExcelWebApi.Tests.TestData
{
    [DataContract]
    public class TestItem
    {
        [Excel(Order = 1)]
        public string Value1 { get; set; }

        [Excel(Order = 2)]
        public string Value2 { get; set; }
    }
}
