using System;
namespace Simplexcel.XlsxInternal
{
    internal static class XlsxEnumFormatter
    {
        internal static string GetXmlValue(this Panes pane)
        {
            switch (pane)
            {
                case Panes.BottomLeft:
                    return "bottomLeft";
                case Panes.BottomRight:
                    return "bottomRight";
                case Panes.TopLeft:
                    return "topLeft";
                case Panes.TopRight:
                    return "topRight";
                default:
                    throw new ArgumentOutOfRangeException(nameof(pane), "Invalid Pane: " + pane);
            }
        }

        internal static string GetXmlValue(this PaneState state)
        {
            switch (state)
            {
                case PaneState.Split:
                    return "split";
                case PaneState.Frozen:
                    return "frozen";
                case PaneState.FrozenSplit:
                    return "frozenSplit";
                default:
                    throw new ArgumentOutOfRangeException(nameof(state), "Invalid Pane State: " + state);
            }
        }
    }
}