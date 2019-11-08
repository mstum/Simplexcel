using System;
namespace Simplexcel.XlsxInternal
{
    internal static class XlsxEnumFormatter
    {
        internal static string GetXmlValue(this Panes pane) => pane switch
        {
            Panes.BottomLeft => "bottomLeft",
            Panes.BottomRight => "bottomRight",
            Panes.TopLeft => "topLeft",
            Panes.TopRight => "topRight",
            _ => throw new ArgumentOutOfRangeException(nameof(pane), "Invalid Pane: " + pane),
        };

        internal static string GetXmlValue(this PaneState state) => state switch
        {
            PaneState.Split => "split",
            PaneState.Frozen => "frozen",
            PaneState.FrozenSplit => "frozenSplit",
            _ => throw new ArgumentOutOfRangeException(nameof(state), "Invalid Pane State: " + state),
        };
    }
}