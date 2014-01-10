using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace ExcelWebApi
{
    public class FormatterUtils
    {
        protected static BindingFlags InstanceBindingFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

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
        /// Get the value of the `Order` property specified on a given member.
        /// </summary>
        /// <param name="member">The member for which to find the `DataMember.Order` value.</param>
        public static Int32 MemberOrder(MemberInfo member)
        {
            var dataMember = GetAttribute<DataMemberAttribute>(member);
            return dataMember.Order;
        }

        /// <summary>
        /// Get a list of all members of a type on which a `DataMemberAttribute` has been
        /// specified, ordered by the value of the `DataMemberAttribute.Order` value on
        /// each member.
        /// </summary>
        /// <param name="type">The data contract type on which to look for members.</param>
        public static List<string> GetDataMemberNames(Type type)
        {
            var memberInfo =  type.GetProperties(InstanceBindingFlags)
                                  .OfType<MemberInfo>()
                                  .Union(type.GetFields(InstanceBindingFlags))
                                  .ToList();

            var memberNames = from p in memberInfo
                              where Attribute.IsDefined(p, typeof(DataMemberAttribute))
                              orderby MemberOrder(p)
                              select p.Name;

            return memberNames.ToList();
        }

        /// <summary>
        /// Get a list of members on the specified type that have an associated
        /// `DataMemberAttribute`.
        /// </summary>
        /// <param name="type">The data contract type on which to look for members.</param>
        public static List<MemberInfo> GetDataMemberInfo(Type type)
        {
            var memberInfo = type.GetProperties(InstanceBindingFlags)
                                 .OfType<MemberInfo>()
                                 .Union(type.GetFields(InstanceBindingFlags));

            var dataMemberInfo = from p in memberInfo
                                 where Attribute.IsDefined(p, typeof(DataMemberAttribute))
                                 orderby MemberOrder(p)
                                 select p;

            return dataMemberInfo.ToList();
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
