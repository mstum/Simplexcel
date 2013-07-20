using System;
using System.Drawing;
using System.Runtime.Serialization;

namespace Simplexcel.XlsxInternal
{
    [DataContract]
    internal class XlsxFont : IEquatable<XlsxFont>
    {
        [DataMember]
        internal string Name { get; set; }
        [DataMember]
        internal int Size { get; set; }
        [DataMember]
        internal bool Bold { get; set; }
        [DataMember]
        internal bool Italic { get; set; }
        [DataMember]
        internal bool Underline { get; set; }
        [DataMember]
        internal Color TextColor { get; set; }

        internal XlsxFont()
        {
            Name = "Calibri";
            Size = 11;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(XlsxFont)) return false;
            return Equals((XlsxFont)obj);
        }

        public bool Equals(XlsxFont other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other.Name, Name) 
                && other.Size == Size 
                && other.Bold.Equals(Bold) 
                && other.Italic.Equals(Italic) 
                && other.Underline.Equals(Underline) 
                && other.TextColor.Equals(TextColor);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = (Name != null ? Name.GetHashCode() : 0);
                result = (result*397) ^ Size;
                result = (result*397) ^ Bold.GetHashCode();
                result = (result*397) ^ Italic.GetHashCode();
                result = (result*397) ^ Underline.GetHashCode();
                result = (result*397) ^ TextColor.GetHashCode();
                return result;
            }
        }

        public static bool operator ==(XlsxFont left, XlsxFont right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxFont left, XlsxFont right)
        {
            return !Equals(left, right);
        }
    }
}
