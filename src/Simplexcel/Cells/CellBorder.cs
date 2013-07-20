using System;
using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// Border around a cell
    /// </summary>
    [Flags]
    [DataContract]
    public enum CellBorder
    {
        /// <summary>
        /// No Border
        /// </summary>
        [EnumMember]
        None = 0,

        /// <summary>
        /// Border at the top
        /// </summary>
        [EnumMember]
        Top = 1,

        /// <summary>
        /// Border on the Right side
        /// </summary>
        [EnumMember]
        Right = 2,

        /// <summary>
        /// Border at the bottom
        /// </summary>
        [EnumMember]
        Bottom = 4,

        /// <summary>
        /// Border on the Left Side
        /// </summary>
        [EnumMember]
        Left = 8,

        /// <summary>
        /// Borders on all four sides
        /// </summary>
        All = Top + Right + Bottom + Left
    }
}
