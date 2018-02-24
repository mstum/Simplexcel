using System;

namespace Simplexcel
{
    internal class GradientStop : IEquatable<GradientStop>
    {
        /// <summary>
        /// The position indicated here indicates the point where the color is pure.
        /// </summary>
        public double Position { get; set; }

        /// <summary>
        /// The pure color of this Gradient Stop
        /// </summary>
        public Color Color { get; set; }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(GradientStop)) return false;
            return Equals((GradientStop)obj);
        }

        public bool Equals(GradientStop other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other.Position, Position)
                && Equals(other.Color, Color);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = Position.GetHashCode();
                result = (result * 397) ^ Color.GetHashCode();
                return result;
            }
        }

        public static bool operator ==(GradientStop left, GradientStop right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(GradientStop left, GradientStop right)
        {
            return !Equals(left, right);
        }
    }
}
