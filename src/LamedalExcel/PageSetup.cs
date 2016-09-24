using System.Runtime.Serialization;
using LamedalExcel.enums;

namespace LamedalExcel
{
    /// <summary>
    /// Page Setup for a Sheet
    /// </summary>
    [DataContract]
    public sealed class PageSetup
    {
        /// <summary>
        /// How many rows to repeat on each page when printing?
        /// </summary>
        [DataMember]
        public int PrintRepeatRows { get; set; }

        /// <summary>
        /// How many columns to repeat on each page when printing?
        /// </summary>
        [DataMember]
        public int PrintRepeatColumns { get; set; }

        /// <summary>
        /// The Orientation of the page, Portrait (default) or Landscape
        /// </summary>
        [DataMember]
        public enOrientation Orientation { get; set; }
    }
}
