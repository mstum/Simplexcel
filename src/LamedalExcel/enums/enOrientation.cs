using System.Runtime.Serialization;

namespace LamedalExcel.enums
{
    /// <summary>
    /// The Orientation of a page
    /// </summary>
    [DataContract]
    public enum enOrientation
    {
        /// <summary>
        /// Portrait (default)
        /// </summary>
        [EnumMember]
        Portrait = 0,
        /// <summary>
        /// Landscape
        /// </summary>
        [EnumMember]
        Landscape
    }
}
