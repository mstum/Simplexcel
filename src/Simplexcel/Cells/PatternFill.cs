using System;

namespace Simplexcel
{
    /// <summary>
    /// A Pattern Fill of a Cell.
    /// </summary>
    public class PatternFill : IEquatable<PatternFill>
    {
        // TODO: Add a Gradient Fill to this.
        // This isn't the structure in the XML, but that's how Excel Presents it, as a "Fill Effect"

        bool firstTimeSet = false;
        private Color? _bgColor;

        /// <summary>
        /// The type of Fill Pattern to use.
        /// Refer to the documentation for a list of each pattern.
        /// </summary>
        public PatternType PatternType { get; set; }

        /// <summary>
        /// The Background Color of the cell
        /// </summary>
        public Color? PatternColor { get; set; }

        /// <summary>
        /// The Background Color of the Fill.
        /// (No effect if <see cref="PatternType"/> is <see cref="PatternType.None"/>)
        /// </summary>
        public Color? BackgroundColor
        {
            get => _bgColor;
            set
            {
                _bgColor = value;

                // PatternType defaults to None, but the first time we set a background color,
                // set it to solid as the user likely wants the background color to show.
                // Further modifications won't change the PatternType, as this is now a delibrate setting
                if (_bgColor.HasValue && PatternType == PatternType.None && !firstTimeSet)
                {
                    PatternType = PatternType.Solid;
                    firstTimeSet = true;
                }
            }
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(PatternFill)) return false;
            return Equals((PatternFill)obj);
        }

        public bool Equals(PatternFill other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other.PatternType, PatternType)
                && Equals(other.PatternColor, PatternColor)
                && Equals(other.PatternColor, PatternColor);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = PatternType.GetHashCode();
                if (PatternColor.HasValue)
                {
                    result = (result * 397) ^ PatternColor.GetHashCode();
                }
                if (PatternColor.HasValue)
                {
                    result = (result * 397) ^ PatternColor.GetHashCode();
                }
                return result;
            }
        }

        public static bool operator ==(PatternFill left, PatternFill right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(PatternFill left, PatternFill right)
        {
            return !Equals(left, right);
        }
    }
}