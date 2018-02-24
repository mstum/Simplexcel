using System;
using System.Collections.Generic;
using System.Linq;

namespace Simplexcel
{
    // TODO: Not Yet Implemented
    internal class GradientFill : IEquatable<GradientFill>
    {
        public IList<GradientStop> Stops { get; }

        public GradientFill()
        {
            Stops = new List<GradientStop>();
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(GradientFill)) return false;
            return Equals((GradientFill)obj);
        }

        public bool Equals(GradientFill other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return other.Stops.SequenceEqual(Stops);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = Stops.GetCollectionHashCode();
                return result;
            }
        }

        public static bool operator ==(GradientFill left, GradientFill right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(GradientFill left, GradientFill right)
        {
            return !Equals(left, right);
        }
    }
}