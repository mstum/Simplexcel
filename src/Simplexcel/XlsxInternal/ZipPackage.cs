using System;
using System.IO;

namespace Simplexcel.XlsxInternal 
{
    internal class ZipPackage
    {
        private readonly Stream _stream;


        internal ZipPackage(Stream underlyingStream)
        {
            if(underlyingStream == null)
            {
                throw new ArgumentNullException(nameof(underlyingStream));
            }

            _stream = underlyingStream;
        } 

        internal ZipPackagePart CreatePart(string targetUri, string contentType){
            throw new Exception(targetUri);
        }

        internal void Flush()
        {
            _stream.Flush();
        }

        internal void Close()
        {

        }

        internal void CreateInternalRelationship(string targetUri, string relationshipType, String id){
            throw new Exception(targetUri);
        }
    }

    internal class ZipPackagePart
    {
        internal Stream GetStream(FileMode mode, FileAccess access){
            throw new Exception();
        }
    }
}