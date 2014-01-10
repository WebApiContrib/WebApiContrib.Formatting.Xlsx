using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ExcelWebApi.Tests
{
    [TestClass]
    public class FormatterUtilsTests
    {
        [TestMethod]
        public void GetAttribute_DataMemberOfTestItem_ReturnsDataMemberAttribute()
        {
            var value1 = typeof(TestItem).GetMember("Value1")[0];
            var dataMemberAttributes = FormatterUtils.GetAttribute<DataMemberAttribute>(value1);

            Assert.IsNotNull(dataMemberAttributes);
            Assert.AreEqual(1, dataMemberAttributes.Order);
        }

        [TestMethod]
        public void MemberOrder_TestItem_ReturnsMemberOrder()
        {
            var testItemType = typeof(TestItem);
            var value1 = testItemType.GetMember("Value1")[0];
            var value2 = testItemType.GetMember("Value2")[0];

            Assert.AreEqual(1, FormatterUtils.MemberOrder(value1), "Value1 should have order 1.");
            Assert.AreEqual(2, FormatterUtils.MemberOrder(value2), "Value2 should have order 2.");
        }

        [TestMethod]
        public void GetDataMemberNames_TestItem_ReturnsMemberNamesInOrder()
        {
            var dataMemberNames = FormatterUtils.GetDataMemberNames(typeof(TestItem));

            Assert.IsNotNull(dataMemberNames);
            Assert.AreEqual(2, dataMemberNames.Count);
            Assert.AreEqual("Value1", dataMemberNames[0]);
            Assert.AreEqual("Value2", dataMemberNames[1]);
        }

        [TestMethod]
        public void GetDataMemberInfo_TestItem_ReturnsMemberInfoList()
        {
            var memberInfo = FormatterUtils.GetDataMemberInfo(typeof(TestItem));

            Assert.IsNotNull(memberInfo);
            Assert.AreEqual(2, memberInfo.Count);
        }

        [TestMethod]
        public void GetEnumerableItemType_ListOfTestItem_ReturnsTestItemType()
        {
            var testItemList = new List<TestItem>();
            var itemType = FormatterUtils.GetEnumerableItemType(testItemList);

            Assert.IsNotNull(itemType);
            Assert.AreEqual(typeof(TestItem), itemType);
        }
    }
}
