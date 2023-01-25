using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using AplusCore.Runtime;
using AplusCore.Types;

using DLR = System.Linq.Expressions;

namespace AplusCore
{
    public static class ExtensionMethods
    {
        public static IEnumerable<string> ToStringArray<T>(this IEnumerable<T> list)
        {
            return list.Select<T, string>(item => item.ToString());
        }

        /// <summary>
        /// Creates an AArray with type AInteger from the input list of integers
        /// </summary>
        /// <param name="list">List of Integers</param>
        /// <returns></returns>
        public static AType ToAArray(this IEnumerable<int> list)
        {
            AType array = AArray.Create(ATypes.AInteger);

            foreach (int item in list)
            {
                array.AddWithNoUpdate(AInteger.Create(item));
            }
            array.UpdateInfo();

            return array;
        }

        /// <summary>
        /// Returns the indexer property for the given <paramref name="target"/> type
        /// with one <paramref name="indexerType"/> parameter
        /// </summary>
        /// <param name="target">The Type where to look the indexer properties</param>
        /// <param name="indexerType">The Type which the indexer property should contain</param>
        /// <returns>an indexer PropertyInfo or null if such property info does not exists</returns>
        public static PropertyInfo GetIndexerProperty(this Type target, Type indexerType)
        {
            PropertyInfo property = null;

            foreach (PropertyInfo info in target.GetProperties())
            {
                ParameterInfo[] parms = info.GetIndexParameters();
                if (parms.Length > 0 && parms[0].ParameterType == indexerType)
                {
                    property = info;
                    break;
                }
            }

            return property;
        }

        public static string ToTypeString(this ATypes type)
        {
            switch (type)
            {
                case ATypes.AInteger:
                    return "int";
                case ATypes.AFloat:
                    return "float";
                case ATypes.ASymbol:
                    return "sym";
                case ATypes.AChar:
                    return "char";
                case ATypes.ANull:
                    return "null";
                case ATypes.ABox:
                    return "box";
                case ATypes.AFunc:
                    return "func";
                default:
                    return "unknown";
            }
        }

        /// <summary>
        /// Sort argument Array, with stable sort.
        /// </summary>
        /// <param name="list"></param>
        public static AType InsertionSortIndex(this AType argument,Func<AType,AType, int> method)
        {
            int[] index = Enumerable.Range(0, argument.Length).ToArray();

            List<AType> list = new List<AType>(argument);

            for (int j = 1; j < list.Count; j++)
            {
                AType key = list[j];

                int i = j - 1;
                for (; i >= 0 && method(list[i], key) > 0; i--)
                {
                    list[i + 1] = list[i];
                    index[i + 1] = index[i];
                }
                list[i + 1] = key;
                index[i + 1] = j;
            }

            return index.ToAArray();
        }

        /// <summary>
        /// Calculates the product of the items.
        /// </summary>
        /// <param name="items">Series of integers.</param>
        /// <returns>Product of the items or 1 if there is no element.</returns>
        public static int Product(this IList<int> items)
        {
            if (items.Count == 0)
            {
                return 1;
            }

            return items.Aggregate((actualProduct, nextFactor) => actualProduct * nextFactor);
        }
    }
}
