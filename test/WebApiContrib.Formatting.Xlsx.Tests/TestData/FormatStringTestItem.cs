using System;
using System.ComponentModel.DataAnnotations;
using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Tests.TestData
{
    public class FormatStringTestItem
    {
        [DisplayFormat(DataFormatString = "{0:D}")]
        public DateTime Value1 { get; set; }

        [DisplayFormat(DataFormatString = "{0:D}")]
        [ExcelColumn(UseDisplayFormatString = true)]
        public DateTime? Value2 { get; set; }

        [DisplayFormat(DataFormatString = "{0:D}")]
        [ExcelColumn(UseDisplayFormatString = false)]
        public DateTime? Value3 { get; set; }
    
        [ExcelColumn(UseDisplayFormatString = true)]
        public DateTime Value4 { get; set; }
    }
}
