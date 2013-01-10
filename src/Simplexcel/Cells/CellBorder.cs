using System;

namespace Simplexcel
{
    /// <summary>
    /// Border around a cell
    /// </summary>
    [Flags]
    public enum CellBorder
    {
        /// <summary>
        /// No Border
        /// </summary>
        None = 0,

        /// <summary>
        /// Border at the top
        /// </summary>
        Top = 1,

        /// <summary>
        /// Border on the Right side
        /// </summary>
        Right = 2,

        /// <summary>
        /// Border at the bottom
        /// </summary>
        Bottom = 4,

        /// <summary>
        /// Border on the Left Side
        /// </summary>
        Left = 8,

        /// <summary>
        /// Borders on all four sides
        /// </summary>
        All = Top + Right + Bottom + Left
    }
}
