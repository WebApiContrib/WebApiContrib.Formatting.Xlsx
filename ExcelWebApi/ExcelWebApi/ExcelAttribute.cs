using System;

namespace ExcelWebApi
{
    public class ExcelAttribute : Attribute
    {
        public ExcelAttribute() { }

        public string NumberFormat { get; set; }
    }
}
