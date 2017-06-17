namespace Simplexcel.XlsxInternal
{
    internal class RelationshipCounter
    {
        private int _count;

        internal int GetNextId()
        {
            _count++;
            return _count;
        }
    }
}
