using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication.ExtendedProtection;
using System.Threading.Tasks;
using System.Web;
using WebApiContrib.Formatting.Xlsx.Serialisation;
using WebApiContrib.Formatting.Xlsx.Tests.TestData;
using WebApiContrib.Formatting.Xlsx.Utils;
using System.Collections;

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
            var types = new[] { // Simple types
                                typeof(bool), typeof(byte), typeof(sbyte), typeof(char),
                                typeof(DateTime), typeof(DateTimeOffset), typeof(decimal),
                                typeof(float), typeof(Guid), typeof(int), typeof(uint),
                                typeof(long), typeof(ulong), typeof(short), typeof(ushort),
                                typeof(TimeSpan), typeof(string), typeof(TestEnum),

                                // Complex types
                                new { anonymous = true }.GetType(), typeof(Array),
                                typeof(IEnumerable<>), typeof(object), typeof(SimpleTestItem) };


            var formatter = new XlsxMediaTypeFormatter();

            foreach (var type in types)
            {
                Assert.IsTrue(formatter.CanWriteType(type));
            }
        }

        [TestMethod]
        public void CanReadType_TypeObject_ReturnsFalse()
        {
            var formatter = new XlsxMediaTypeFormatter();

            Assert.IsFalse(formatter.CanReadType(typeof(object)));
        }

        [TestMethod]
        public void WriteToStreamAsync_WithListOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" } }.ToList();

            var expected = new[] { new object[] { "Value1",       "Value2"       },
                                   new object[] { data[0].Value1, data[0].Value2 },
                                   new object[] { data[1].Value1, data[1].Value2 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new[] { new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" },
                               new SimpleTestItem { Value1 = "3,1", Value2 = "3,2" }  };

            var expected = new[] { new object[] { "Value1",       "Value2"       },
                                   new object[] { data[0].Value1, data[0].Value2 },
                                   new object[] { data[1].Value1, data[1].Value2 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfFormatStringTestItem_ValuesFormattedAppropriately()
        {
            var tomorrow = DateTime.Today.AddDays(1);
            var formattedDate = tomorrow.ToString("D");

            // Let 1 Jan 1990 = day 1 and add 1 for each day since, counting 1990 as a leap year due to an Excel bug.
            var excelDate = (tomorrow - new DateTime(1900, 1, 1)).TotalDays + 2;
            var excelDateStr = excelDate.ToString();


            var data = new[] { new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = tomorrow,
                                                          Value3 = tomorrow,
                                                          Value4 = tomorrow },

                               new FormatStringTestItem { Value1 = tomorrow,
                                                          Value2 = null,
                                                          Value3 = null,
                                                          Value4 = tomorrow } };

            var expected = new[] { new[] { "Value1",     "Value2",      "Value3",     "Value4"     },
                                   new[] { excelDateStr, formattedDate, excelDateStr, excelDateStr },
                                   new[] { excelDateStr, string.Empty,  string.Empty, excelDateStr }  };

            GenerateAndCompareWorksheet(data, expected);
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

            var expected = new[] { new[] { "Value1", "Value2", "Value3",     "Value4"     },
                                   new[] { "True",   "Yes",    "True",       "Yes"        },
                                   new[] { "False",  "No",     "False",      "No"         },
                                   new[] { "True",   "Yes",    string.Empty, string.Empty }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithSimpleTestItem_WritesExcelDocumentToStream()
        {
            var data = new SimpleTestItem { Value1 = "2,1", Value2 = "2,2" };

            var expected = new[] { new[] { "Value1",    "Value2"    },
                                   new[] { data.Value1, data.Value2 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithEmptyListOfComplexTestItem_DoesNotCrash()
        {
            var data = new ComplexTestItem[0];

            var expected = new[] { new object[] { "Header 4", "Value1", "Header 5", "Header 3", "Value2" } };

            GenerateAndCompareWorksheet(data, expected);
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

            var expected = new[] { new object[] { "Header 4",  "Value1",    "Header 5",             "Header 3",             "Value2"    },
                                   new object[] { data.Value4, data.Value1, data.Value5.ToString(), data.Value3.ToString(), data.Value2 }  };

            var sheet = GenerateAndCompareWorksheet(data, expected);

            Assert.AreEqual("???.???", sheet.Cells[2, 1].Style.Numberformat.Format, "NumberFormat of A2 is incorrect.");
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


            var expected = new[] { new object[] { "Header 4",     "Value1",       "Header 5",                "Header 3",                "Value2"       },
                                   new object[] { data[0].Value4, data[0].Value1, data[0].Value5.ToString(), data[0].Value3.ToString(), data[0].Value2 },
                                   new object[] { data[1].Value4, data[1].Value1, data[1].Value5.ToString(), data[1].Value3.ToString(), data[1].Value2 }  };

            var sheet = GenerateAndCompareWorksheet(data, expected);

            Assert.AreEqual("???.???", sheet.Cells[2, 1].Style.Numberformat.Format, "NumberFormat of A2 is incorrect.");
            Assert.AreEqual("???.???", sheet.Cells[3, 1].Style.Numberformat.Format, "NumberFormat of A3 is incorrect.");
        }

        [TestMethod]
        public void WriteToStreamAsync_WithAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today };

            var expected = new[] { new object[] { "prop1",    "prop2",    "prop3"    },
                                   new object[] { data.prop1, data.prop2, data.prop3 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfAnonymousObject_WritesExcelDocumentToStream()
        {
            var data = new[] {
                new { prop1 = "val1", prop2 = 1.0, prop3 = DateTime.Today },
                new { prop1 = "val2", prop2 = 2.0, prop3 = DateTime.Today.AddDays(1) }
            };

            var expected = new[] { new object[] { "prop1",       "prop2",       "prop3"       },
                                   new object[] { data[0].prop1, data[0].prop2, data[0].prop3 },
                                   new object[] { data[1].prop1, data[1].prop2, data[1].prop3 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithString_WritesExcelDocumentToStream()
        {
            var data = "Test";

            var expected = new[] { new[] { data } };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfString_WritesExcelDocumentToStream()
        {
            var data = new[] { "1,1", "2,1" };

            var expected = new[] { new[] { data[0] },
                                   new[] { data[1] }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithInt32_WritesExcelDocumentToStream()
        {
            var data = 100;

            var expected = new[] { new[] { data } };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfInt32_WritesExcelDocumentToStream()
        {
            var data = new[] { 100, 200 };

            var expected = new[] { new[] { data[0] },
                                   new[] { data[1] }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithDateTime_WritesExcelDocumentToStream()
        {
            var data = DateTime.Today;

            var expected = new[] { new object[] { data } };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfDateTime_WritesExcelDocumentToStream()
        {
            var data = new[] { DateTime.Today, DateTime.Today.AddDays(1) };

            var expected = new[] { new[] { data[0] },
                                   new[] { data[1] }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithExpandoObject_WritesExcelDocumentToStream()
        {
            dynamic data = new ExpandoObject();

            data.Value1 = "Test";
            data.Value2 = 1;

            var expected = new[] { new object[] { "Value1",    "Value2"    },
                                   new object[] { data.Value1, data.Value2 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void WriteToStreamAsync_WithArrayOfExpandoObject_WritesExcelDocumentToStream()
        {
            dynamic row1 = new ExpandoObject();
            dynamic row2 = new ExpandoObject();

            row1.Value1 = "Test";
            row1.Value2 = 1;
            row2.Value1 = true;
            row2.Value2 = DateTime.Today;

            var data = new[] { row1, row2 };

            var expected = new[] { new object[] { "Value1",       "Value2"       },
                                   new object[] { data[0].Value1, data[0].Value2 },
                                   new object[] { data[1].Value1, data[1].Value2 }  };

            GenerateAndCompareWorksheet(data, expected);
        }

        [TestMethod]
        public void XlsxMediaTypeFormatter_WithDefaultHeaderHeight_DefaultsToSameHeightForAllCells()
        {
            var data = new[] { new SimpleTestItem { Value1 = "A1", Value2 = "B1" },
                               new SimpleTestItem { Value1 = "A1", Value2 = "B2" }  };

            var formatter = new XlsxMediaTypeFormatter();

            var sheet = GetWorksheetFromStream(formatter, data);

            Assert.AreNotEqual(sheet.Row(1).Height, 0d, "HeaderHeight should not be zero");
            Assert.AreEqual(sheet.Row(1).Height, sheet.Row(2).Height, "HeaderHeight should be the same as other rows");
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
            Assert.IsTrue(sheet.Cells[3, 3].Style.Font.Bold, "Value in C3 should be bold.");
            Assert.AreEqual(18f, sheet.Cells[1, 1].Style.Font.Size, "Header in A1 should be in size 18 font.");
            Assert.AreEqual(18f, sheet.Cells[1, 3].Style.Font.Size, "Header in C1 should be in size 18 font.");
            Assert.AreEqual(15f, sheet.Cells[2, 1].Style.Font.Size, "Value in A2 should be in size 15 font.");
            Assert.AreEqual(15f, sheet.Cells[3, 3].Style.Font.Size, "Value in C3 should be in size 15 font.");
            Assert.AreEqual(ExcelBorderStyle.Thick, sheet.Cells[1, 1].Style.Border.Bottom.Style, "Header in A1 should have a thick border.");
            Assert.AreEqual(ExcelBorderStyle.Thick, sheet.Cells[1, 3].Style.Border.Bottom.Style, "Header in C1 should have a thick border.");
            Assert.AreEqual(ExcelBorderStyle.None, sheet.Cells[2, 1].Style.Border.Bottom.Style, "Value in A2 should have no border.");
            Assert.AreEqual(ExcelBorderStyle.None, sheet.Cells[3, 3].Style.Border.Bottom.Style, "Value in C3 should have no border.");
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

        [TestMethod]
        public void XlsxMediaTypeFormatter_WithPerRequestColumnResolver_ReturnsSpecifiedProperties()
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


            var expected = new[] { new object[] { "Header 4",     "Value1",       "Header 5"                },
                                   new object[] { data[0].Value4, data[0].Value1, data[0].Value5.ToString() },
                                   new object[] { data[1].Value4, data[1].Value1, data[1].Value5.ToString() }  };

            var serialiseValues = new[] { "Value1", "Value4", "Value5" };

            var formatter = new XlsxMediaTypeFormatter();
            formatter.DefaultSerializer.Resolver = new PerRequestColumnResolver();

            HttpContextFactory.SetCurrentContext(new FakeHttpContext());
            HttpContextFactory.Current.Items[PerRequestColumnResolver.DEFAULT_KEY] = serialiseValues;

            var sheet = GenerateAndCompareWorksheet(data, expected, formatter);
        }

        [TestMethod]
        public void XlsxMediaTypeFormatter_WithPerRequestColumnResolverCustomOrder_ReturnsSpecifiedProperties()
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


            var expected = new[] { new object[] { "Value1",       "Header 4",     "Header 5"                },
                                   new object[] { data[0].Value1, data[0].Value4, data[0].Value5.ToString() },
                                   new object[] { data[1].Value1, data[1].Value4, data[1].Value5.ToString() }  };

            var serialiseValues = new[] { "Value1", "Value4", "Value5" };

            var formatter = new XlsxMediaTypeFormatter();
            formatter.DefaultSerializer.Resolver = new PerRequestColumnResolver(useCustomOrder: true);

            HttpContextFactory.SetCurrentContext(new FakeHttpContext());
            HttpContextFactory.Current.Items[PerRequestColumnResolver.DEFAULT_KEY] = serialiseValues;

            var sheet = GenerateAndCompareWorksheet(data, expected, formatter);
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

        public class FakeHttpContext : HttpContextBase
        {
            private IDictionary _items = new Dictionary<string, object>();

            public override IDictionary Items
            {
                get
                {
                    return _items;
                }
            }
        }
        #endregion

        #region Utilities
        /// <summary>
        /// Generate the serialised worksheet and ensure that it is formatted as expected.
        /// </summary>
        /// <typeparam name="TItem">Type of items to be serialised.</typeparam>
        /// <typeparam name="TExpected">Type of items in expected results array, usually <c>object</c>.</typeparam>
        /// <param name="data">Data to be serialised.</param>
        /// <param name="expected">Expected format of the generated worksheet.</param>
        /// <param name="formatter">Optional custom formatter instance to use for serialisation.</param>
        /// <returns>The generated <c>ExcelWorksheet</c> containing the serialised data.</returns>
        public ExcelWorksheet GenerateAndCompareWorksheet<TItem, TExpected>(TItem data,
                                                                            TExpected[][] expected,
                                                                            XlsxMediaTypeFormatter formatter = null)
        {
            var sheet = GetWorksheetFromStream(formatter ?? new XlsxMediaTypeFormatter(), data);

            CompareWorksheet(sheet, expected);

            return sheet;
        }

        /// <summary>
        /// Generate a worksheet containing the specified data using the provided <c>XlsxMediaTypeFormatter</c>
        /// instance.
        /// </summary>
        /// <typeparam name="TItem">Type of items to be serialised.</typeparam>
        /// <param name="formatter">Formatter instance to use for serialisation.</param>
        /// <param name="data">Data to be serialised.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Ensure that the data in a generated worksheet is serialised as expected.
        /// </summary>
        /// <typeparam name="TExpected">Type of items in expected results array, usually <c>object</c>.</typeparam>
        /// <param name="sheet">The generated <c>ExcelWorksheet</c> containing the serialised data.</param>
        /// <param name="expected">Expected format of the generated worksheet.</param>
        public void CompareWorksheet<TExpected>(ExcelWorksheet sheet, TExpected[][] expected)
        {
            Assert.IsNotNull(sheet.Dimension, "Worksheet has no cells.");

            Assert.AreEqual(expected.Length, sheet.Dimension.End.Row, "Wrong number of rows.");
            Assert.AreEqual(expected[0].Length, sheet.Dimension.End.Column, "Wrong number of columns.");

            for (var i = 0; i < expected.Length; i++)
            {
                for (var j = 0; j < expected[i].Length; j++)
                {
                    var value = expected[i][j];

                    var method = typeof(ExcelWorksheet).GetMethods()
                                                       .Where(m => m.Name == "GetValue")
                                                       .First(m => m.ContainsGenericParameters);

                    var type = typeof(TExpected) == typeof(object) ? value.GetType() : typeof(TExpected);

                    var cellValue = method.MakeGenericMethod(type)
                                          .Invoke(sheet, new object[] { i + 1, j + 1 });

                    var column = (char)('A' + j);
                    var row = i + 1;
                    var message = String.Format("Value in {0}{1} is incorrect.", column, row);

                    Assert.AreEqual(value, cellValue, message);
                }
            }
        }
        #endregion
    }
}
