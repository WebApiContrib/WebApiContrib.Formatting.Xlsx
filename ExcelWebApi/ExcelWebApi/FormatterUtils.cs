using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace ExcelWebApi
{
    public class FormatterUtils
    {
        protected const BindingFlags PublicInstanceBindingFlags = BindingFlags.Instance | BindingFlags.Public;

        /// <summary>
        /// Get the `Attribute` object of the specified type associated with a member. 
        /// </summary>
        /// <typeparam name="TAttribute">Type of attribute to get.</typeparam>
        /// <param name="memberInfo">The member to look for the attribute on.</param>
        public static TAttribute GetAttribute<TAttribute>(MemberInfo memberInfo)
        {
            var attributes = from a in memberInfo.GetCustomAttributes(true)
                             where a is TAttribute
                             select a;

            return (TAttribute) attributes.FirstOrDefault();
        }

        /// <summary>
        /// Get the value of the `ExcelAttribute.Order` attribute associated with a given
        /// member. If not found, will default to the `DataMember.Order` value.
        /// </summary>
        /// <param name="member">The member for which to find the `ExcelAttribute.Order` value.</param>
        public static Int32 MemberOrder(MemberInfo member)
        {
            var excelProperty = GetAttribute<ExcelAttribute>(member);
            if (excelProperty != null && excelProperty._order.HasValue)
                return excelProperty.Order;

            var dataMember = GetAttribute<DataMemberAttribute>(member);
            if (dataMember != null)
                return dataMember.Order;

            return -1;
        }

        /// <summary>
        /// Get the value of the `ExcelAttribute.Ignore` attribute associated with a given
        /// member. If not found, will default to the `DataMember.Ignore` value.
        /// </summary>
        /// <param name="member">The member for which to find the `ExcelAttribute.Ignore` value.</param>
        public static bool IsMemberIgnored(MemberInfo member)
        {
            var excelProperty = GetAttribute<ExcelAttribute>(member);
            if (excelProperty != null)
                return excelProperty.Ignore;

            return false;
        }

        /// <summary>
        /// Get an ordered list of non-ignored public instance property names of a type.
        /// </summary>
        /// <param name="type">The type on which to look for members.</param>
        public static List<string> GetMemberNames(Type type)
        {
            var memberInfo =  type.GetProperties(PublicInstanceBindingFlags)
                                  .OfType<MemberInfo>()
                                  .Union(type.GetFields(PublicInstanceBindingFlags));

            var memberNames = from p in memberInfo
                              where !IsMemberIgnored(p)
                              orderby MemberOrder(p)
                              select p.Name;

            return memberNames.ToList();
        }

        /// <summary>
        /// Get an ordered list of <c>MemberInfo</c> for non-ignored public instance
        /// properties on the specified type.
        /// </summary>
        /// <param name="type">The type on which to look for members.</param>
        public static List<MemberInfo> GetMemberInfo(Type type)
        {
            var memberInfo = type.GetProperties(PublicInstanceBindingFlags)
                                 .OfType<MemberInfo>()
                                 .Union(type.GetFields(PublicInstanceBindingFlags));

            var orderedMemberInfo = from p in memberInfo
                                    where !IsMemberIgnored(p)
                                    orderby MemberOrder(p)
                                    select p;

            return orderedMemberInfo.ToList();
        }

        /// <summary>
        /// Get the item type of an object that implements `IEnumerable`.
        /// </summary>
        /// <param name="value">An instance whose underlying type to check.</param>
        /// <returns></returns>
        public static Type GetEnumerableItemType(object value)
        {
            Type[] interfaces = value.GetType().GetInterfaces();
            foreach (Type i in interfaces)
            {
                if (i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {
                    return i.GetGenericArguments()[0];
                }
            }

            return null;
        }

        /// <summary>
        /// Get a field or property value from an object.
        /// </summary>
        /// <param name="obj">The object whose property we want.</param>
        /// <param name="name">The name of the field or property we want.</param>
        public static object GetFieldOrPropertyValue(object obj, string name)
        {
            var type = obj.GetType();

            var member = type.GetField(name) ?? type.GetProperty(name) as MemberInfo;

            if (member == null) return null;

            object value;

            switch (member.MemberType)
            {
                case MemberTypes.Property:
                    value = ((PropertyInfo)member).GetValue(obj, null);
                    break;
                case MemberTypes.Field:
                    value = ((FieldInfo)member).GetValue(obj);
                    break;
                default:
                    value = null;
                    break;
            }

            return value;
        }

        /// <summary>
        /// Get a field or property value from an object.
        /// </summary>
        /// <param name="obj">The object whose property we want.</param>
        /// <param name="name">The name of the field or property we want.</param>
        public static T GetFieldOrPropertyValue<T>(object obj, string name)
        {
            var type = obj.GetType();

            var member = type.GetField(name) ?? type.GetProperty(name) as MemberInfo;

            if (member == null) return default(T);

            object value;

            switch (member.MemberType)
            {
                case MemberTypes.Property:
                    value = ((PropertyInfo)member).GetValue(obj, null);
                    break;
                case MemberTypes.Field:
                    value = ((FieldInfo)member).GetValue(obj);
                    break;
                default:
                    value = null;
                    break;
            }

            return (T)value;
        }

    }
}
