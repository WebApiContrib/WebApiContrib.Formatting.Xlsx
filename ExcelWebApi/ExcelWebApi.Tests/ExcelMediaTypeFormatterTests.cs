using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication.ExtendedProtection;
using System.Threading.Tasks;

namespace ExcelWebApi.Tests
{
    [TestClass]
    public class ExcelMediaTypeFormatterTests
    {
        [TestMethod]
        public void SupportedMediaTypes_SupportsExcelMediaTypes()
        {
            var formatter = new ExcelMediaTypeFormatter();

            Assert.IsTrue(
                formatter.SupportedMediaTypes.Any(s => s.MediaType == "application/vnd.ms-excel"),
                "XLS media type not supported."
            );

            Assert.IsTrue(
                formatter.SupportedMediaTypes.Any(s => s.MediaType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                "XLSX media type not supported."
            );
        }

        [TestMethod]
        public void CanWriteType_TypeEnumerable_CanWriteType()
        {
            var formatter = new ExcelMediaTypeFormatter();
            Assert.IsTrue(formatter.CanWriteType(typeof(IEnumerable<object>)), "Cannot write enumerable types.");
        }

        [TestMethod]
        public void CanWriteType_TypeObject_CannotWriteType()
        {
            var formatter = new ExcelMediaTypeFormatter();
            Assert.IsFalse(formatter.CanWriteType(typeof(object)), "Can write any type.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithGenericCollection_WritesExcelDocumentToStream()
        {
            var ms = new MemoryStream();

            var content = new FakeContent();
            content.Headers.ContentType = new MediaTypeHeaderValue("application/atom+xml");

            var formatter = new ExcelMediaTypeFormatter();

            var task = formatter.WriteToStreamAsync(typeof(List<TestItem>),
                new List<TestItem> { new TestItem { Value = "Row 1" }, new TestItem { Value = "Row 2" } },
                ms,
                content,
                new FakeTransport()
            );

            task.Wait();

            ms.Seek(0, SeekOrigin.Begin);

            try
            {
                var package = new ExcelPackage(ms);
                var sheet = package.Workbook.Worksheets[1];

                Assert.IsTrue(sheet.Dimension.End.Row == 2, "Worksheet should have two rows.");
                Assert.IsTrue(sheet.Dimension.End.Column == 1, "Worksheet should have one column.");
                Assert.IsTrue(sheet.GetValue<string>(1, 1) == "Row 1", "Value in first row is incorrect.");
                Assert.IsTrue(sheet.GetValue<string>(2, 1) == "Row 2", "Value in second row is incorrect.");
            }
            catch (Exception e)
            {
                Assert.Fail("Could not read stream as an Excel workbook. (Exception: {0})", e.Message);
            }
        }

        #region Fakes and test-related classes
        public class TestItem
        {
            public string Value { get; set; }
        }

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
    }
}
