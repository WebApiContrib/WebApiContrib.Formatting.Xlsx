using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using WebApiContrib.Formatting.Xlsx.Tests.TestData;
using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Tests
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
        public void GetMemberNames_AnonymousType_ReturnsMemberNamesInOrderDefined()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var memberNames = FormatterUtils.GetMemberNames(anonymous.GetType());

            Assert.IsNotNull(memberNames);
            Assert.AreEqual(2, memberNames.Count);
            Assert.AreEqual("prop1", memberNames[0]);
            Assert.AreEqual("prop2", memberNames[1]);
        }

        [TestMethod]
        public void GetMemberInfo_SimpleTestItem_ReturnsMemberInfoList()
        {
            var memberInfo = FormatterUtils.GetMemberInfo(typeof(SimpleTestItem));

            Assert.IsNotNull(memberInfo);
            Assert.AreEqual(2, memberInfo.Count);
        }

        [TestMethod]
        public void GetMemberInfo_AnonymousType_ReturnsMemberInfoList()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var memberInfo = FormatterUtils.GetMemberInfo(anonymous.GetType());

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

        [TestMethod]
        public void GetEnumerableItemType_ArrayOfSimpleTestItem_ReturnsTestItemType()
        {
            var testItemArray = typeof(SimpleTestItem[]);
            var itemType = FormatterUtils.GetEnumerableItemType(testItemArray);

            Assert.IsNotNull(itemType);
            Assert.AreEqual(typeof(SimpleTestItem), itemType);
        }

        [TestMethod]
        public void GetEnumerableItemType_ArrayOfAnonymousObject_ReturnsTestItemType()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var anonymousArray = new[] { anonymous };

            var itemType = FormatterUtils.GetEnumerableItemType(anonymousArray.GetType());

            Assert.IsNotNull(itemType);
            Assert.AreEqual(anonymous.GetType(), itemType);
        }

        [TestMethod]
        public void GetEnumerableItemType_ListOfAnonymousObject_ReturnsTestItemType()
        {
            var anonymous = new { prop1 = "value1", prop2 = "value2" };
            var anonymousList = new[] { anonymous }.ToList();

            var itemType = FormatterUtils.GetEnumerableItemType(anonymousList.GetType());

            Assert.IsNotNull(itemType);
            Assert.AreEqual(anonymous.GetType(), itemType);
        }

        [TestMethod]
        public void GetFieldOrPropertyValue_ComplexTestItem_ReturnsPropertyValues()
        {
            var obj = new ComplexTestItem() {
                Value1 = "Value 1",
                Value2 = DateTime.Today,
                Value3 = true,
                Value4 = 100.1,
                Value5 = TestEnum.Second,
                Value6 = "Value 6"
            };

            Assert.AreEqual(obj.Value1, FormatterUtils.GetFieldOrPropertyValue(obj, "Value1"));
            Assert.AreEqual(obj.Value2, FormatterUtils.GetFieldOrPropertyValue(obj, "Value2"));
            Assert.AreEqual(obj.Value3, FormatterUtils.GetFieldOrPropertyValue(obj, "Value3"));
            Assert.AreEqual(obj.Value4, FormatterUtils.GetFieldOrPropertyValue(obj, "Value4"));
            Assert.AreEqual(obj.Value5, FormatterUtils.GetFieldOrPropertyValue(obj, "Value5"));
            Assert.AreEqual(obj.Value6, FormatterUtils.GetFieldOrPropertyValue(obj, "Value6"));
        }

        [TestMethod]
        public void GetFieldOrPropertyValueT_ComplexTestItem_ReturnsPropertyValues()
        {
            var obj = new ComplexTestItem() {
                Value1 = "Value 1",
                Value2 = DateTime.Today,
                Value3 = true,
                Value4 = 100.1,
                Value5 = TestEnum.Second,
                Value6 = "Value 6"
            };

            Assert.AreEqual(obj.Value1, FormatterUtils.GetFieldOrPropertyValue<string>(obj, "Value1"));
            Assert.AreEqual(obj.Value2, FormatterUtils.GetFieldOrPropertyValue<DateTime>(obj, "Value2"));
            Assert.AreEqual(obj.Value3, FormatterUtils.GetFieldOrPropertyValue<bool>(obj, "Value3"));
            Assert.AreEqual(obj.Value4, FormatterUtils.GetFieldOrPropertyValue<double>(obj, "Value4"));
            Assert.AreEqual(obj.Value5, FormatterUtils.GetFieldOrPropertyValue<TestEnum>(obj, "Value5"));
            Assert.AreEqual(obj.Value6, FormatterUtils.GetFieldOrPropertyValue<string>(obj, "Value6"));
        }

        [TestMethod]
        public void GetFieldOrPropertyValue_AnonymousObject_ReturnsPropertyValues()
        {
            var obj = new { prop1 = "test", prop2 = 2.0, prop3 = DateTime.Today };

            Assert.AreEqual(obj.prop1, FormatterUtils.GetFieldOrPropertyValue(obj, "prop1"));
            Assert.AreEqual(obj.prop2, FormatterUtils.GetFieldOrPropertyValue(obj, "prop2"));
            Assert.AreEqual(obj.prop3, FormatterUtils.GetFieldOrPropertyValue(obj, "prop3"));
        }

        [TestMethod]
        public void GetFieldOrPropertyValueT_AnonymousObject_ReturnsPropertyValues()
        {
            var obj = new { prop1 = "test", prop2 = 2.0, prop3 = DateTime.Today };

            Assert.AreEqual(obj.prop1, FormatterUtils.GetFieldOrPropertyValue<string>(obj, "prop1"));
            Assert.AreEqual(obj.prop2, FormatterUtils.GetFieldOrPropertyValue<double>(obj, "prop2"));
            Assert.AreEqual(obj.prop3, FormatterUtils.GetFieldOrPropertyValue<DateTime>(obj, "prop3"));
        }

        [TestMethod]
        public void IsSimpleType_SimpleTypes_ReturnsTrue()
        {
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(bool)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(byte)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(sbyte)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(char)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(DateTime)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(DateTimeOffset)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(decimal)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(double)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(float)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(Guid)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(int)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(uint)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(long)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(ulong)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(short)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(TimeSpan)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(ushort)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(string)));
            Assert.IsTrue(FormatterUtils.IsSimpleType(typeof(TestEnum)));
        }

        [TestMethod]
        public void IsSimpleType_ComplexTypes_ReturnsFalse()
        {
            var anonymous = new { prop = "val" };

            Assert.IsFalse(FormatterUtils.IsSimpleType(anonymous.GetType()));
            Assert.IsFalse(FormatterUtils.IsSimpleType(typeof(Array)));
            Assert.IsFalse(FormatterUtils.IsSimpleType(typeof(IEnumerable<>)));
            Assert.IsFalse(FormatterUtils.IsSimpleType(typeof(object)));
            Assert.IsFalse(FormatterUtils.IsSimpleType(typeof(SimpleTestItem)));
        }
    }
}
