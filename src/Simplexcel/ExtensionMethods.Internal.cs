using System;
using System.Collections.Generic;
using System.Reflection;

namespace Simplexcel
{
    internal static class ExtensionMethods
    {
        internal static IEnumerable<PropertyInfo> GetAllProperties(this TypeInfo typeInfo) => GetAllForType(typeInfo, ti => ti.DeclaredProperties);

        private static IEnumerable<T> GetAllForType<T>(TypeInfo typeInfo, Func<TypeInfo, IEnumerable<T>> accessor)
        {
            if (typeInfo == null)
            {
                yield break;
            }

            // The Stack is to make sure that we fetch Base Type Properties first
            var baseTypes = new Stack<TypeInfo>();
            while (typeInfo != null)
            {
                baseTypes.Push(typeInfo);
                typeInfo = typeInfo.BaseType?.GetTypeInfo();
            }

            while (baseTypes.Count > 0)
            {
                var ti = baseTypes.Pop();
                foreach (var t in accessor(ti))
                {
                    yield return t;
                }
            }
        }

        internal static int GetCollectionHashCode<T>(this IEnumerable<T> input)
        {
            var hashCode = 0;
            if (input != null)
            {
                foreach (var item in input)
                {
                    var itemHashCode = item == null ? 0 : item.GetHashCode();
                    hashCode = (hashCode * 397) ^ itemHashCode;
                }
            }
            return hashCode;
        }

        private static readonly char[] HexEncodingTable = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

        /// <remarks>
        /// Taken from Bouncy Castle and slightly changed to create a string rather than write to a stream.
        /// https://www.bouncycastle.org/
        ///
        /// Licensed under MIT License
        /// https://www.bouncycastle.org/csharp/licence.html
        /// </remarks>
        internal static string ToHexString(this IReadOnlyList<byte> bytes, int offset = 0, int length = -1)
        {
            if (bytes == null || bytes.Count == 0) { return string.Empty; }
            if (length == -1) { length = bytes.Count; }

            var result = new char[2 * length];
            var index = offset;
            for (int i = offset; i < (offset + length); i++)
            {
                var v = bytes[i];
                result[index++] = HexEncodingTable[v >> 4];
                result[index++] = HexEncodingTable[v & 0xf];
            }
            return new string(result);
        }
    }
}
