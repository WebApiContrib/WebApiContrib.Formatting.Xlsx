using System;

namespace XlsxForWebApi
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        // Nullable parameters not allowed on attributes. :(
        internal int? _order;

        /// <summary>
        /// Control the output of this property when serialized to Excel.
        /// </summary>
        public ExcelColumnAttribute() { }

        /// <summary>
        /// Control the output of this property when serialized to Excel.
        /// </summary>
        public ExcelColumnAttribute(string header) {
            Header = header;
        }

        /// <summary>
        /// Column header to use for this property.
        /// </summary>
        public string Header { get; set; }

        /// <summary>
        /// Ignore this property when serializing to Excel.
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// Override the serialized order of this property in the generated Excel document.
        /// </summary>public int Order
        public int Order
        {
            get { return _order ?? default(int); }
            set { _order = value; }
        }

        /// <summary>
        /// Apply the specified Excel number format string to this property in the generated Excel output.
        /// </summary>
        public string NumberFormat { get; set; }
    }
}
