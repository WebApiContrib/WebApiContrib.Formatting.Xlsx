using XlsxForWebApi.Tests.TestData;
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
using System.Runtime.Serialization;
using System.Security.Authentication.ExtendedProtection;
using System.Threading.Tasks;
using System.Xml;

namespace XlsxForWebApi.Tests
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
        public void CanWriteType_TypeEnumerable_CanWriteType()
        {
            var formatter = new XlsxMediaTypeFormatter();

            Assert.IsTrue(formatter.CanWriteType(typeof(IEnumerable<object>)));
        }

        [TestMethod]
        public void CanWriteType_TypeObject_CanWriteType()
        {
            var formatter = new XlsxMediaTypeFormatter();

            Assert.IsTrue(formatter.CanWriteType(typeof(object)));
        }

        [TestMethod]
        public void WriteToStreamAsync_WithListOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new List<SimpleTestItem> { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                                                  new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);
            
            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Value in first cell is incorrect.");
            Assert.AreEqual(data[1].Value2, sheet.GetValue<string>(3, 2), "Value in last cell is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithListOfComplexTestItem_WritesExcelDocumentToStream()
        {
            var data = new List<ComplexTestItem> { new ComplexTestItem { Value1 = "Item 1",
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
                                                                         Value6 = "Also ignored" } };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(5.0, sheet.Dimension.End.Column, "Worksheet should have five columns.");
            Assert.AreEqual("Header 4", sheet.GetValue<string>(1, 1), "First column header is incorrect.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 2), "Second column header is incorrect.");
            Assert.AreEqual("Header 5", sheet.GetValue<string>(1, 3), "Third column header is incorrect.");
            Assert.AreEqual("Header 3", sheet.GetValue<string>(1, 4), "Fourth column header is incorrect.");
            Assert.AreEqual("Value2", sheet.GetValue<string>(1, 5), "Fifth column header is incorrect.");
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
        public void WriteToStreamAsync_WithArrayOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new SimpleTestItem[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                                              new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Value in first cell is incorrect.");
            Assert.AreEqual(data[1].Value2, sheet.GetValue<string>(3, 2), "Value in last cell is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Value in first cell is incorrect.");
            Assert.AreEqual(data.Value2, sheet.GetValue<string>(2, 2), "Value in last cell is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithString_WritesExcelDocumentToStream()
        {
            var data = "Test";

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<string>(1, 1), "Value is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfString_WritesExcelDocumentToStream()
        {
            var data = new string[] { "1,1", "2,1" };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have two rows.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<string>(1, 1), "Value in first cell is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<string>(2, 1), "Value in last cell is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithInt32_WritesExcelDocumentToStream()
        {
            var data = 100;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<int>(1, 1), "Value is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfInt32_WritesExcelDocumentToStream()
        {
            var data = new int[] { 100, 200 };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<int>(1, 1), "Value is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<int>(2, 1), "Value is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithDateTime_WritesExcelDocumentToStream()
        {
            var data = DateTime.Today;

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data, sheet.GetValue<DateTime>(1, 1), "Value is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfDateTime_WritesExcelDocumentToStream()
        {
            var data = new DateTime[] { DateTime.Today, DateTime.Today.AddDays(1) };

            var sheet = GetWorksheetFromStream(new XlsxMediaTypeFormatter(), data);

            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(2.0, sheet.Dimension.End.Row, "Worksheet should have one row.");
            Assert.AreEqual(1.0, sheet.Dimension.End.Column, "Worksheet should have one column.");
            Assert.AreEqual(data[0], sheet.GetValue<DateTime>(1, 1), "Value is incorrect.");
            Assert.AreEqual(data[1], sheet.GetValue<DateTime>(2, 1), "Value is incorrect.");
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
