using ExcelWebApi.Tests.TestData;
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

namespace ExcelWebApi.Tests
{
    [TestClass]
    public class ExcelMediaTypeFormatterTests
    {
        const string XlsMimeType = "application/vnd.ms-excel";
        const string XlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        [TestMethod]
        public void SupportedMediaTypes_SupportsExcelMediaTypes()
        {
            var formatter = new ExcelMediaTypeFormatter();

            Assert.IsTrue(formatter.SupportedMediaTypes.Any(s => s.MediaType == XlsMimeType),
                          "XLS media type not supported.");

            Assert.IsTrue(formatter.SupportedMediaTypes.Any(s => s.MediaType == XlsxMimeType),
                          "XLSX media type not supported.");
        }

        [TestMethod]
        public void CanWriteType_TypeEnumerable_CanWriteType()
        {
            var formatter = new ExcelMediaTypeFormatter();

            Assert.IsTrue(formatter.CanWriteType(typeof(IEnumerable<object>)),
                          "Cannot write enumerable types.");
        }

        [TestMethod]
        public void CanWriteType_TypeObject_CannotWriteType()
        {
            var formatter = new ExcelMediaTypeFormatter();

            Assert.IsFalse(formatter.CanWriteType(typeof(object)),
                           "Can write any type.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithGenericCollection_WritesExcelDocumentToStream()
        {
            var formatter = new ExcelMediaTypeFormatter();

            var data = new List<SimpleTestItem> { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                                            new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var sheet = GetWorksheetFromStream(formatter, data);
            
            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");
            Assert.AreEqual(3.0, sheet.Dimension.End.Row, "Worksheet should have three rows (including header column).");
            Assert.AreEqual(2.0, sheet.Dimension.End.Column, "Worksheet should have two columns.");
            Assert.AreEqual("Value1", sheet.GetValue<string>(1, 1), "Value in first cell is incorrect.");
            Assert.AreEqual("3,2", sheet.GetValue<string>(3, 2), "Value in last cell is incorrect.");
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
        public ExcelWorksheet GetWorksheetFromStream<TItem>(ExcelMediaTypeFormatter formatter, IEnumerable<TItem> data)
        {
            var ms = new MemoryStream();

            formatter = new ExcelMediaTypeFormatter(autoFit: true,
                                                    autoFilter: true,
                                                    freezeHeader: true,
                                                    headerHeight: 20.0f,
                                                    cellHeight: 18f,
                                                    cellStyle: (ExcelStyle s) => s.WrapText = true,
                                                    headerStyle: (ExcelStyle s) => s.Border.Bottom.Style = ExcelBorderStyle.Double);

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
