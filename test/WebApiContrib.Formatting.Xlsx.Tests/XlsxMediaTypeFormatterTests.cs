using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication.ExtendedProtection;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Tests.TestData;

namespace WebApiContrib.Formatting.Xlsx.Tests
{
    [TestClass]
    public class XlsxMediaTypeFormatterTests
    {
        const string XlsMimeType = "application/vnd.ms-excel";
        const string XlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        [TestMethod]
        public void SupportedMediaTypes_SupportsExcelMediaTypes()
        {
            var formatter = new XlsxMediaTypeFormatter();

            Assert.IsTrue(formatter.SupportedMediaTypes.Any(s => s.MediaType == XlsMimeType),
                          "XLS media type not supported.");

            Assert.IsTrue(formatter.SupportedMediaTypes.Any(s => s.MediaType == XlsxMimeType),
                          "XLSX media type not supported.");
        }

        [TestMethod]
        public void CanWriteType_AnyType_ReturnsTrue()
        {
            var formatter = new XlsxMediaTypeFormatter();

            // Simple types
            Assert.IsTrue(formatter.CanWriteType(typeof(bool)));
            Assert.IsTrue(formatter.CanWriteType(typeof(byte)));
            Assert.IsTrue(formatter.CanWriteType(typeof(sbyte)));
            Assert.IsTrue(formatter.CanWriteType(typeof(char)));
            Assert.IsTrue(formatter.CanWriteType(typeof(DateTime)));
            Assert.IsTrue(formatter.CanWriteType(typeof(DateTimeOffset)));
            Assert.IsTrue(formatter.CanWriteType(typeof(decimal)));
            Assert.IsTrue(formatter.CanWriteType(typeof(double)));
            Assert.IsTrue(formatter.CanWriteType(typeof(float)));
            Assert.IsTrue(formatter.CanWriteType(typeof(Guid)));
            Assert.IsTrue(formatter.CanWriteType(typeof(int)));
            Assert.IsTrue(formatter.CanWriteType(typeof(uint)));
            Assert.IsTrue(formatter.CanWriteType(typeof(long)));
            Assert.IsTrue(formatter.CanWriteType(typeof(ulong)));
            Assert.IsTrue(formatter.CanWriteType(typeof(short)));
            Assert.IsTrue(formatter.CanWriteType(typeof(TimeSpan)));
            Assert.IsTrue(formatter.CanWriteType(typeof(ushort)));
            Assert.IsTrue(formatter.CanWriteType(typeof(string)));
            Assert.IsTrue(formatter.CanWriteType(typeof(TestEnum)));

            // Complex types
            var anonymous = new { prop = "val" };
            Assert.IsTrue(formatter.CanWriteType(anonymous.GetType()));
            Assert.IsTrue(formatter.CanWriteType(typeof(Array)));
            Assert.IsTrue(formatter.CanWriteType(typeof(IEnumerable<>)));
            Assert.IsTrue(formatter.CanWriteType(typeof(object)));
            Assert.IsTrue(formatter.CanWriteType(typeof(SimpleTestItem)));
        }

        [TestMethod]
        public void CanReadType_TypeObject_ReturnsFalse()
        {
            var formatter = new XlsxMediaTypeFormatter();

            Assert.IsFalse (formatter.CanReadType(typeof(object)));
        }

        [TestMethod]
        public void WriteToStreamAsync_WithListOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" } }.ToList();

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);
            
            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual(data[0].Value1, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(data[0].Value2, sheet.GetValue<string>(2, 2), "Value in B2 is incorrect.");
            Assert.AreEqual(data[1].Value1, sheet.GetValue<string>(3, 1), "Value in A3 is incorrect.");
            Assert.AreEqual(data[1].Value2, sheet.GetValue<string>(3, 2), "Value in B3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual(data[0].Value1, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(data[0].Value2, sheet.GetValue<string>(2, 2), "Value in B2 is incorrect.");
            Assert.AreEqual(data[1].Value1, sheet.GetValue<string>(3, 1), "Value in A3 is incorrect.");
            Assert.AreEqual(data[1].Value2, sheet.GetValue<string>(3, 2), "Value in B3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfFormatStringTestItem_ValuesFormattedAppropriately()
        {
            var tomorrow = DateTime.Today.AddDays(1);
            var formattedDate = tomorrow.ToString("D");
            
            var data = new[] { new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = tomorrow,
                                                          Value3 = tomorrow,
                                                          Value4 = tomorrow },

                               new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = null,
                                                          Value3 = null,
                                                          Value4 = tomorrow } };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(4.0, sheet.Dimension.End.Column, "Worksheet should have four columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("Value3", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual("Value4", sheet.GetValue<string>(1, 4), "Header in D1 is incorrect.");
            Assert.AreNotEqual(formattedDate, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(formattedDate, sheet.GetValue<string>(2, 2), "Value in B2 is incorrect.");
            Assert.AreNotEqual(formattedDate, sheet.GetValue<string>(2, 3), "Value in C2 is incorrect.");
            Assert.AreNotEqual(formattedDate, sheet.GetValue<string>(2, 4), "Value in D2 is incorrect.");
            Assert.AreNotEqual(formattedDate, sheet.GetValue<string>(3, 1), "Value in A3 is incorrect.");
            Assert.AreEqual(string.Empty, sheet.GetValue<string>(3, 2), "Value in B3 is incorrect.");
            Assert.AreEqual(string.Empty, sheet.GetValue<string>(3, 3), "Value in C3 is incorrect.");
            Assert.AreNotEqual(formattedDate, sheet.GetValue<string>(3, 4), "Value in D3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfBooleanTestItem_TrueOrFalseValueUsedAsAppropriate()
        {
            var data = new[] { new BooleanTestItem { Value1 = true,
                                                     Value2 = true,
                                                     Value3 = true,
                                                     Value4 = true },

                               new BooleanTestItem { Value1 = false,
                                                     Value2 = false,
                                                     Value3 = false,
                                                     Value4 = false },
            
                               new BooleanTestItem { Value1 = true,
                                                     Value2 = true,
                                                     Value3 = null,
                                                     Value4 = null } };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(4.0, sheet.Dimension.End.Row, "Worksheet should have four rows (including header column).");
            Assert.AreEqual(4.0, sheet.Dimension.End.Column, "Worksheet should have four columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("Value3", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual("Value4", sheet.GetValue<string>(1, 4), "Header in D1 is incorrect.");
            Assert.AreEqual("True", sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual("Yes", sheet.GetValue<string>(2, 2), "Value in B2 is incorrect.");
            Assert.AreEqual("True", sheet.GetValue<string>(2, 3), "Value in C2 is incorrect.");
            Assert.AreEqual("Yes", sheet.GetValue<string>(2, 4), "Value in D2 is incorrect.");
            Assert.AreEqual("False", sheet.GetValue<string>(3, 1), "Value in A3 is incorrect.");
            Assert.AreEqual("No", sheet.GetValue<string>(3, 2), "Value in B3 is incorrect.");
            Assert.AreEqual("False", sheet.GetValue<string>(3, 3), "Value in C3 is incorrect.");
            Assert.AreEqual("No", sheet.GetValue<string>(3, 4), "Value in D3 is incorrect.");
            Assert.AreEqual("True", sheet.GetValue<string>(4, 1), "Value in A4 is incorrect.");
            Assert.AreEqual("Yes", sheet.GetValue<string>(4, 2), "Value in B4 is incorrect.");
            Assert.AreEqual(string.Empty, sheet.GetValue<string>(4, 3), "Value in C4 is incorrect.");
            Assert.AreEqual(string.Empty, sheet.GetValue<string>(4, 4), "Value in D4 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual(data.Value1, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(data.Value2, sheet.GetValue<string>(2, 2), "Value in B2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithComplexTestItem_WritesExcelDocumentToStream()
        {
            var data = new ComplexTestItem { Value1 = "Item 1",
                                             Value2 = DateTime.Today,
                                             Value3 = true,
                                             Value4 = 100.1,
                                             Value5 = TestEnum.First,
                                             Value6 = "Ignored" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows (including header column).");
            Assert.AreEqual(5.0, sheet.Dimension.End.Column, "Worksheet should have five columns.");
            Assert.AreEqual("Header 4", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("Header 5", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual("Header 3", sheet.GetValue<string>(1, 4), "Header in D1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 5), "Header in E1 is incorrect.");
            Assert.AreEqual(data.Value4, sheet.GetValue<double>(2, 1), "Data in A2 is incorrect.");
            Assert.AreEqual("???.???", sheet.Cells[2, 1].Style.Numberformat.Format, "NumberFormat of A2 is incorrect.");
            Assert.AreEqual(data.Value1, sheet.GetValue<string>(2, 2), "Data in B2 is incorrect.");
            Assert.AreEqual(data.Value5.ToString(), sheet.GetValue<string>(2, 3), "Data in C2 is incorrect.");
            Assert.AreEqual(data.Value3.ToString(), sheet.GetValue<string>(2, 4), "Data in D2 is incorrect.");
            Assert.AreEqual(data.Value2, sheet.GetValue<DateTime>(2, 5), "Data in E2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithListOfComplexTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new ComplexTestItem { Value1 = "Item 1",
                                                     Value2 = DateTime.Today,
                                                     Value3 = true,
                                                     Value4 = 100.1,
                                                     Value5 = TestEnum.First,
                                                     Value6 = "Ignored" },

                               new ComplexTestItem { Value1 = "Item 2",
                                                     Value2 = DateTime.Today.AddDays(1),
                                                     Value3 = false,
                                                     Value4 = 200.2,
                                                     Value5 = TestEnum.Second,
                                                     Value6 = "Also ignored" } }.ToList();

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(5.0, sheet.Dimension.End.Column, "Worksheet should have five columns.");
            Assert.AreEqual("Header 4", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("Header 5", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual("Header 3", sheet.GetValue<string>(1, 4), "Header in D1 is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 5), "Header in E1 is incorrect.");
            Assert.AreEqual(data[0].Value4, sheet.GetValue<double>(2, 1), "Data in A2 is incorrect.");
            Assert.AreEqual("???.???", sheet.Cells[2, 1].Style.Numberformat.Format, "NumberFormat of A2 is incorrect.");
            Assert.AreEqual(data[0].Value1, sheet.GetValue<string>(2, 2), "Data in B2 is incorrect.");
            Assert.AreEqual(data[0].Value5.ToString(), sheet.GetValue<string>(2, 3), "Data in C2 is incorrect.");
            Assert.AreEqual(data[0].Value3.ToString(), sheet.GetValue<string>(2, 4), "Data in D2 is incorrect.");
            Assert.AreEqual(data[0].Value2, sheet.GetValue<DateTime>(2, 5), "Data in E2 is incorrect.");
            Assert.AreEqual(data[1].Value4, sheet.GetValue<double>(3, 1), "Data in A3 is incorrect.");
            Assert.AreEqual("???.???", sheet.Cells[3, 1].Style.Numberformat.Format, "NumberFormat of A3 is incorrect.");
            Assert.AreEqual(data[1].Value1, sheet.GetValue<string>(3, 2), "Data in B3 is incorrect.");
            Assert.AreEqual(data[1].Value5.ToString(), sheet.GetValue<string>(3, 3), "Data in C3 is incorrect.");
            Assert.AreEqual(data[1].Value3.ToString(), sheet.GetValue<string>(3, 4), "Data in D3 is incorrect.");
            Assert.AreEqual(data[1].Value2, sheet.GetValue<DateTime>(3, 5), "Data in E3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Column, "Worksheet should have three columns.");
            Assert.AreEqual("prop1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("prop2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("prop3", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual(data.prop1, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(data.prop2, sheet.GetValue<double>(2, 2), "Value in B2 is incorrect.");
            Assert.AreEqual(data.prop3, sheet.GetValue<DateTime>(2, 3), "Value in C2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new[] {
                new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today },
                new { prop1 = "val2", prop2 = 2.0, prop3 = DateTime.Today.AddDays(1) }
            };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Column, "Worksheet should have three columns.");
            Assert.AreEqual("prop1", sheet.GetValue<string>(1, 1), "Header in A1 is incorrect.");
            Assert.AreEqual("prop2", sheet.GetValue<string>(1, 2), "Header in B1 is incorrect.");
            Assert.AreEqual("prop3", sheet.GetValue<string>(1, 3), "Header in C1 is incorrect.");
            Assert.AreEqual(data[0].prop1, sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
            Assert.AreEqual(data[0].prop2, sheet.GetValue<double>(2, 2), "Value in B2 is incorrect.");
            Assert.AreEqual(data[0].prop3, sheet.GetValue<DateTime>(2, 3), "Value in C2 is incorrect.");
            Assert.AreEqual(data[1].prop1, sheet.GetValue<string>(3, 1), "Value in A3 is incorrect.");
            Assert.AreEqual(data[1].prop2, sheet.GetValue<double>(3, 2), "Value in B3 is incorrect.");
            Assert.AreEqual(data[1].prop3, sheet.GetValue<DateTime>(3, 3), "Value in C3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithString_WritesExcelDocumentToStream()
        {
            var data = "Test";

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<string>(1, 1), "Value in A1 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfString_WritesExcelDocumentToStream()
        {
            var data = new[] { "1,1", "2,1" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<string>(1, 1), "Value in A1 is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<string>(2, 1), "Value in A2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithInt32_WritesExcelDocumentToStream()
        {
            var data = 100;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<int>(1, 1), "Value in A1 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfInt32_WritesExcelDocumentToStream()
        {
            var data = new[] { 100, 200 };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<int>(1, 1), "Value in A1 is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<int>(2, 1), "Value in A2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithDateTime_WritesExcelDocumentToStream()
        {
            var data = DateTime.Today;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<DateTime>(1, 1), "Value in A1 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfDateTime_WritesExcelDocumentToStream()
        {
            var data = new[] { DateTime.Today, DateTime.Today.AddDays(1) };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<DateTime>(1, 1), "Value in A1 is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<DateTime>(2, 1), "Value in A2 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithCellAndHeaderFormats_WritesFormattedExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var formatter = new XlsxMediaTypeFormatter(
                cellStyle: (ExcelStyle s) =>
                {
                    s.Font.Size = 15f;
                    s.Font.Bold = true;
                },
                headerStyle: (ExcelStyle s) =>
                {
                    s.Font.Size = 18f;
                    s.Border.Bottom.Style = ExcelBorderStyle.Thick;
                }
            );

            var sheet = GetWorksheetFromStream(formatter, data);

            Assert.IsTrue(sheet.Cells[1, 1].Style.Font.Bold, "Header in A1 should be bold.");
            Assert.IsTrue(sheet.Cells[1, 2].Style.Font.Bold, "Header in B1 should be bold.");
            Assert.IsTrue(sheet.Cells[1, 3].Style.Font.Bold, "Header in C1 should be bold.");
            Assert.IsTrue(sheet.Cells[2, 1].Style.Font.Bold, "Value in A2 should be bold.");
            Assert.IsTrue(sheet.Cells[2, 2].Style.Font.Bold, "Value in B2 should be bold.");
            Assert.IsTrue(sheet.Cells[2, 3].Style.Font.Bold, "Value in C2 should be bold.");
            Assert.IsTrue(sheet.Cells[3, 1].Style.Font.Bold, "Value in A3 should be bold.");
            Assert.IsTrue(sheet.Cells[3, 2].Style.Font.Bold, "Value in B3 should be bold.");
            Assert.IsTrue(sheet.Cells[3, 3].Style.Font.Bold, "Value in C3 should be bold.");
            Assert.AreEqual(18f, sheet.Cells[1, 1].Style.Font.Size, "Header in A1 should be in size 18 font.");
            Assert.AreEqual(18f, sheet.Cells[1, 2].Style.Font.Size, "Header in B1 should be in size 18 font.");
            Assert.AreEqual(18f, sheet.Cells[1, 3].Style.Font.Size, "Header in C1 should be in size 18 font.");
            Assert.AreEqual(15f, sheet.Cells[2, 1].Style.Font.Size, "Value in A2 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[2, 2].Style.Font.Size, "Value in B2 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[2, 3].Style.Font.Size, "Value in C2 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[3, 1].Style.Font.Size, "Value in A3 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[3, 2].Style.Font.Size, "Value in B3 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[3, 3].Style.Font.Size, "Value in C3 should be in size 15 font.");
            Assert.AreEqual(ExcelBorderStyle.Thick, sheet.Cells[1, 1].Style.Border.Bottom.Style, "Header in A1 should have a thick border.");
            Assert.AreEqual(ExcelBorderStyle.Thick, sheet.Cells[1, 2].Style.Border.Bottom.Style, "Header in B1 should have a thick border.");
            Assert.AreEqual(ExcelBorderStyle.Thick, sheet.Cells[1, 3].Style.Border.Bottom.Style, "Header in C1 should have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[2, 1].Style.Border.Bottom.Style, "Value in A2 should NOT have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[2, 2].Style.Border.Bottom.Style, "Value in B2 should NOT have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[2, 3].Style.Border.Bottom.Style, "Value in C2 should NOT have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[3, 1].Style.Border.Bottom.Style, "Value in A3 should NOT have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[3, 2].Style.Border.Bottom.Style, "Value in B3 should NOT have a thick border.");
            Assert.AreNotEqual(ExcelBorderStyle.Thick, sheet.Cells[3, 3].Style.Border.Bottom.Style, "Value in C3 should NOT have a thick border.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithHeaderRowHeight_WritesFormattedExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var formatter = new XlsxMediaTypeFormatter(headerHeight: 30f);

            var sheet = GetWorksheetFromStream(formatter, data);

            Assert.AreEqual(30f, sheet.Row(1).Height, "Row 1 should have height 30.");
        }
        
        #region Fakes and test-related classes
        public class FakeContent : HttpContent
        {
            public FakeContent() : base() { }

            protected override Task SerializeToStreamAsync(Stream stream, TransportContext context)
            {
                throw new NotImplementedException();
            }

            protected override bool TryComputeLength(out long length)
            {
                throw new NotImplementedException();
            }
        }

        public class FakeTransport : TransportContext
        {
            public override ChannelBinding GetChannelBinding(ChannelBindingKind kind)
            {
                throw new NotImplementedException();
            }
        }
        #endregion

        #region Utilities
        public ExcelWorksheet GetWorksheetFromStream<TItem>(XlsxMediaTypeFormatter formatter, TItem data)
        {
            var ms = new MemoryStream();

            var content = new FakeContent();
            content.Headers.ContentType = new MediaTypeHeaderValue("application/atom+xml");

            var task = formatter.WriteToStreamAsync(typeof(IEnumerable<TItem>),
                                                    data,
                                                    ms,
                                                    content,
                                                    new FakeTransport());

            task.Wait();

            ms.Seek(0, SeekOrigin.Begin);

            var package = new ExcelPackage(ms);
            return package.Workbook.Worksheets[1];

        }
        #endregion
    }
}
