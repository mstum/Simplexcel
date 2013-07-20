using System.Runtime.Serialization;

namespace Simplexcel
{
    /// <summary>
    /// The Orientation of a page
    /// </summary>
    [DataContract]
    public enum Orientation
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
