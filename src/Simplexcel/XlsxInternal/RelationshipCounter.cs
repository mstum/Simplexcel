namespace Simplexcel.XlsxInternal
{
    // ToDo: This is hack-ish, make this better.
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
