using System;
using System.Runtime.Serialization;

namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// The Style Information about a cell. Isolated from the Cell object because we need to compare styles when building the styles.xml file
    /// </summary>
    [DataContract]
    internal class XlsxCellStyle : IEquatable<XlsxCellStyle>
    {
        internal XlsxCellStyle()
        {
            Font = new XlsxFont();
        }

        [DataMember]
        internal XlsxFont Font { get; set; }
        
        [DataMember]
        internal CellBorder Border { get; set; }
        
        [DataMember]
        internal string Format { get; set; }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(XlsxCellStyle)) return false;
            return Equals((XlsxCellStyle)obj);
        }

        public bool Equals(XlsxCellStyle other)
        {
            if (ReferenceEquals(null, other))
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Equals(other.Border, Border)
                && other.Font.Equals(Font)
                && other.Format.Equals(Format)
                ;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = Border.GetHashCode();
                result = (result * 397) ^ Font.GetHashCode();
                result = (result * 397) ^ Format.GetHashCode();
                return result;
            }
        }

        public static bool operator ==(XlsxCellStyle left, XlsxCellStyle right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxCellStyle left, XlsxCellStyle right)
        {
            return !Equals(left, right);
        }
    }
}