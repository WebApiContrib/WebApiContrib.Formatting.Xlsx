using System.Runtime.Serialization;

namespace ExcelWebApi.Tests
{
    [DataContract]
    public class TestItem
    {
        [DataMember(Order = 1)]
        [Excel]
        public string Value1 { get; set; }

        [DataMember(Order = 2)]
        [Excel]
        public string Value2 { get; set; }
    }
}
