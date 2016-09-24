using System.Runtime.Serialization;

namespace LamedalExcel.enums
{
    /// <summary>
    /// The Type of a Cell, important for handling the values corrently (e.g., Excel not converting long numbers into scientific notation or removing leading zeroes)
    /// </summary>
    [DataContract]
    public enum enCellType
    {
        /// <summary>
        /// Cell contains text, so use it exactly as specified
        /// </summary>
        [EnumMember]
        Text,

        /// <summary>
        /// Cell contains a number (may strip leading zeroes but sorts properly)
        /// </summary>
        [EnumMember]
        Number
        // TODO: Date
    }
}
