using System;

namespace WebApiContrib.Formatting.Xlsx.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelDocumentAttribute : Attribute
    {
        
        /// <summary>
        /// Set properties of Excel documents generated from this type.
        /// </summary>
        public ExcelDocumentAttribute() { }

        /// <summary>
        /// Set properties of Excel documents generated from this type.
        /// </summary>
        /// <param name="fileName">The preferred file name for an Excel document generated from this type.</param>
        public ExcelDocumentAttribute(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// The preferred file name for an Excel document generated from this type.
        /// </summary>
        public string FileName { get; set; }
    }
}
