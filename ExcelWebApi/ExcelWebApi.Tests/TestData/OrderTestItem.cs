using System.Runtime.Serialization;
using ExcelWebApi;

namespace ExcelWebApi.Tests.TestData
{
    [DataContract]
    public class OrderTestItem
    {
        [Excel]
        public string Value1 { get; set; }

        [Excel]
        public string Value2 { get; set; }
    }
}
