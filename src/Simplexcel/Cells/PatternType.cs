namespace Simplexcel
{
    /// <summary>
    /// The Type of the pattern in a <see cref="PatternFill"/>.
    /// </summary>
    public enum PatternType
    {
        /// <summary>
        /// none
        /// </summary>
        None = 0,

        /// <summary>
        /// solid
        /// </summary>
        Solid,

        /// <summary>
        /// 75% Gray
        /// darkGray
        /// </summary>
        Gray750,

        /// <summary>
        /// 50% Gray
        /// mediumGray
        /// </summary>
        Gray500,

        /// <summary>
        /// 25% Gray
        /// lightGray
        /// </summary>
        Gray250,

        /// <summary>
        /// 12.5% Gray
        /// gray125
        /// </summary>
        Gray125,

        /// <summary>
        /// 6.25% Gray
        /// gray0625
        /// </summary>
        Gray0625,
        
        /// <summary>
        /// Horizontal Stripe
        /// darkHorizontal
        /// </summary>
        HorizontalStripe,

        /// <summary>
        /// Vertical Stripe
        /// darkVertical
        /// </summary>
        VerticalStripe,

        /// <summary>
        /// Reverse Diagonal Stripe
        /// darkDown
        /// </summary>
        ReverseDiagonalStripe,

        /// <summary>
        /// Diagonal Stripe
        /// darkUp
        /// </summary>
        DiagonalStripe,

        /// <summary>
        /// Diagonal Crosshatch
        /// darkGrid
        /// </summary>
        DiagonalCrosshatch,

        /// <summary>
        /// Thick Diagonal Crosshatch
        /// darkTrellis
        /// </summary>
        ThickDiagonalCrosshatch,

        /// <summary>
        /// Thin Horizontal Stripe
        /// lightHorizontal
        /// </summary>
        ThinHorizontalStripe,

        /// <summary>
        /// Thin Vertical Stripe
        /// lightVertical
        /// </summary>
        ThinVerticalStripe,

        /// <summary>
        /// Thin Reverse Diagonal Stripe
        /// lightDown
        /// </summary>
        ThinReverseDiagonalStripe,

        /// <summary>
        /// Thin Diagonal Stripe
        /// lightUp
        /// </summary>
        ThinDiagonalStripe,

        /// <summary>
        /// Thin Horizontal Crosshatch
        /// lightGrid
        /// </summary>
        ThinHorizontalCrosshatch,

        /// <summary>
        /// Thin Diagonal Crosshatch
        /// lightTrellis
        /// </summary>
        ThinDiagonalCrosshatch
    }
}
