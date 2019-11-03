using System;
using System.Diagnostics;

namespace Simplexcel
{
    /// <summary>
    /// The version of the Simplexcel Library.
    /// </summary>
    public static class SimplexcelVersion
    {
        /// <summary>
        /// The version of the Simplexcel Library.
        /// This might include a suffix for development versions, e.g., "2.3.0.177-v3-dev"
        /// </summary>
        public static string VersionString { get; }

        /// <summary>
        /// The version of the Simplexcel Library.
        /// This does not indicate if this is a development version.
        /// </summary>
        public static Version Version { get; }

        /// <summary>
        /// The Public Key that was used when signing Simplexcel
        /// </summary>
        public static string PublicKey { get; }

        /// <summary>
        /// The Public Key Token that was used when signing Simplexcel
        /// </summary>
        public static string PublicKeyToken { get; }

        static SimplexcelVersion()
        {
            // AssemblyVersion is always 3.0.0.0 due to strong naming
            // [assembly: AssemblyVersion("3.0.0.0")]
            // [assembly: AssemblyFileVersion("3.0.0.177")]
            // [assembly: AssemblyInformationalVersion("3.0.0.177-v3-dev")]
            var assembly = typeof(SimplexcelVersion).Assembly;
            VersionString = FileVersionInfo.GetVersionInfo(assembly.Location).ProductVersion;
            Version = new Version(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);

            var asmName = assembly.GetName();
            PublicKey = asmName.GetPublicKey().ToHexString();
            PublicKeyToken = asmName.GetPublicKeyToken().ToHexString();
        }
    }
}
