using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Tests.TestData
{
    public class BooleanValueTestItem
    {
        public bool Value1 { get; set; }

        [ExcelColumn(TrueValue="Yes", FalseValue="No")]
        public bool Value2 { get; set; }
    }
}
