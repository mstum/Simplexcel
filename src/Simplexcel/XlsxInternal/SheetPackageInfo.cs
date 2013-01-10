namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// Information about a Worksheet in the Package
    /// </summary>
    internal class SheetPackageInfo
    {
        internal int SheetId { get; set; }
        internal string RelationshipId { get; set; }
        internal string SheetName { get; set; }
        internal string RepeatRows { get; set; }
        internal string RepeatCols { get; set; }

        internal SheetPackageInfo()
        {
            RelationshipId = string.Empty;
            SheetName = string.Empty;
            RepeatRows = string.Empty;
            RepeatCols = string.Empty;
        }
    }
}
