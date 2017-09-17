namespace Simplexcel
{
    public class Pane
    {
        internal const int DefaultXSplit = 0;
        internal const int DefaultYSplit = 0;
        internal const Panes DefaultActivePane = Panes.TopLeft;
        internal const PaneState DefaultState = PaneState.Split;

        public int? XSplit { get; set; }

        public int? YSplit { get; set; }

        public Panes? ActivePane { get; set; }

        public PaneState? State { get; set; }

        public string TopLeftCell { get; set; }
    }
}