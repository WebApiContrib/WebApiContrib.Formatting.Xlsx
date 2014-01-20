using XlsxForWebApi.Tests.TestData;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace XlsxForWebApi.Tests
{
    [TestClass]
    public class FormatterUtilsTests
    {
        [TestMethod]
        public void GetAttribute_ExcelColumnAttributeOfComplexTestItemValue2_ExcelColumnAttribute()
        {
            var value2 = typeof(ComplexTestItem).GetMember("Value2")[0];
            var excelAttribute = FormatterUtils.GetAttribute<ExcelColumnAttribute>(value2);

            Assert.IsNotNull(excelAttribute);
            Assert.AreEqual(2, excelAttribute.Order);
        }

        [TestMethod]
        public void GetAttribute_ExcelDocumentAttributeOfComplexTestItem_ExcelDocumentAttribute()
        {
            var complexTestItem = typeof(ComplexTestItem);
            var excelAttribute = FormatterUtils.GetAttribute<ExcelDocumentAttribute>(complexTestItem);

            Assert.IsNotNull(excelAttribute);
            Assert.AreEqual("Complex test item", excelAttribute.FileName);
        }

        [TestMethod]
        public void MemberOrder_SimpleTestItem_ReturnsMemberOrder()
        {
            var testItemType = typeof(SimpleTestItem);
            var value1 = testItemType.GetMember("Value1")[0];
            var value2 = testItemType.GetMember("Value2")[0];

            Assert.AreEqual(-1, FormatterUtils.MemberOrder(value1), "Value1 should have order -1.");
            Assert.AreEqual(-1, FormatterUtils.MemberOrder(value2), "Value2 should have order -1.");
        }

        [TestMethod]
        public void MemberOrder_ComplexTestItem_ReturnsMemberOrder()
        {
            var testItemType = typeof(ComplexTestItem);
            var value1 = testItemType.GetMember("Value1")[0];
            var value2 = testItemType.GetMember("Value2")[0];
            var value3 = testItemType.GetMember("Value3")[0];
            var value4 = testItemType.GetMember("Value4")[0];
            var value5 = testItemType.GetMember("Value5")[0];
            var value6 = testItemType.GetMember("Value6")[0];

            Assert.AreEqual(-1, FormatterUtils.MemberOrder(value1), "Value1 should have order -1.");
            Assert.AreEqual( 2, FormatterUtils.MemberOrder(value2), "Value2 should have order 2." );
            Assert.AreEqual( 1, FormatterUtils.MemberOrder(value3), "Value3 should have order 1." );
            Assert.AreEqual(-2, FormatterUtils.MemberOrder(value4), "Value4 should have order -2.");
            Assert.AreEqual(-1, FormatterUtils.MemberOrder(value5), "Value5 should have order -1.");
            Assert.AreEqual(-1, FormatterUtils.MemberOrder(value6), "Value6 should have order -1.");
        }

        [TestMethod]
        public void GetMemberNames_SimpleTestItem_ReturnsMemberNamesInOrder()
        {
            var memberNames = FormatterUtils.GetMemberNames(typeof(SimpleTestItem));

            Assert.IsNotNull(memberNames);
            Assert.AreEqual(2, memberNames.Count);
            Assert.AreEqual("Value1", memberNames[0]);
            Assert.AreEqual("Value2", memberNames[1]);
        }

        [TestMethod]
        public void GetMemberNames_ComplexTestItem_ReturnsMemberNamesInOrder()
        {
            var memberNames = FormatterUtils.GetMemberNames(typeof(ComplexTestItem));

            Assert.IsNotNull(memberNames);
            Assert.AreEqual(5, memberNames.Count);
            Assert.AreEqual("Value4", memberNames[0]);
            Assert.AreEqual("Value1", memberNames[1]);
            Assert.AreEqual("Value5", memberNames[2]);
            Assert.AreEqual("Value3", memberNames[3]);
            Assert.AreEqual("Value2", memberNames[4]);
        }

        [TestMethod]
        public void GetMemberInfo_SimpleTestItem_ReturnsMemberInfoList()
        {
            var memberInfo = FormatterUtils.GetMemberInfo(typeof(SimpleTestItem));

            Assert.IsNotNull(memberInfo);
            Assert.AreEqual(2, memberInfo.Count);
        }

        [TestMethod]
        public void GetEnumerableItemType_ListOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemList = typeof(List<SimpleTestItem>);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemList);

            Assert.IsNotNull(itemType);
            Assert.AreEqual(typeof(SimpleTestItem), itemType);
        }

        [TestMethod]
        public void GetEnumerableItemType_IEnumerableOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemList = typeof(IEnumerable<SimpleTestItem>);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemList);

            Assert.IsNotNull(itemType);
            Assert.AreEqual(typeof(SimpleTestItem), itemType);
        }
    }
}
