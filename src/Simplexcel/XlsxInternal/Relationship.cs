namespace Simplexcel.XlsxInternal
{
    /// <summary>
    /// A Relationship inside the Package
    /// </summary>
    internal class Relationship
    {
        public string Id { get; set; }
        public XmlFile Target { get; set; }
        public string Type { get; set; }

        public Relationship(RelationshipCounter counter)
        {
            Id = "r" + counter.GetNextId();
        }
    }
}
