using Xunit;

namespace Simplexcel.Tests
{
    public class CellAddressTests
    {
        [Fact]
        public void CellAddress_ToString_ProperlyConverts()
        {
            Assert.Equal("A1", new CellAddress(0, 0).ToString());
            Assert.Equal("B1", new CellAddress(0, 1).ToString());
            Assert.Equal("A2", new CellAddress(1, 0).ToString());
            Assert.Equal("B2", new CellAddress(1, 1).ToString());
            Assert.Equal("Z1", new CellAddress(0, 25).ToString());
            Assert.Equal("AA1", new CellAddress(0, 26).ToString());
            Assert.Equal("BA1", new CellAddress(0, 52).ToString());
            Assert.Equal("BZ1", new CellAddress(0, 77).ToString());
            Assert.Equal("CA1", new CellAddress(0, 78).ToString());
            Assert.Equal("IV1", new CellAddress(0, 255).ToString());
            Assert.Equal("CA1", new CellAddress(0, 78).ToString());
            Assert.Equal("ZZ1", new CellAddress(0, 701).ToString());
            Assert.Equal("AAA1", new CellAddress(0, 702).ToString());
            Assert.Equal("AAZ1", new CellAddress(0, 727).ToString());
            Assert.Equal("ABA1", new CellAddress(0, 728).ToString());
            Assert.Equal("XFD1", new CellAddress(0, 16383).ToString());
            Assert.Equal("ABA1000", new CellAddress(999, 728).ToString());            
        }

        [Fact]
        public void CellAdress_FromString_ProperlyParses()
        {
            Assert.Equal(0, new CellAddress("A1").Column);
            Assert.Equal(0, new CellAddress("A1").Row);

            Assert.Equal(1, new CellAddress("B1").Column);
            Assert.Equal(0, new CellAddress("B1").Row);

            Assert.Equal(0, new CellAddress("A2").Column);
            Assert.Equal(1, new CellAddress("A2").Row);

            Assert.Equal(1, new CellAddress("B2").Column);
            Assert.Equal(1, new CellAddress("B2").Row);

            Assert.Equal(25, new CellAddress("Z1").Column);
            Assert.Equal(0, new CellAddress("Z1").Row);

            Assert.Equal(26, new CellAddress("AA1").Column);
            Assert.Equal(0, new CellAddress("AA1").Row);

            Assert.Equal(52, new CellAddress("BA1").Column);
            Assert.Equal(0, new CellAddress("BA1").Row);

            Assert.Equal(77, new CellAddress("BZ1").Column);
            Assert.Equal(0, new CellAddress("BZ1").Row);

            Assert.Equal(78, new CellAddress("CA1").Column);
            Assert.Equal(0, new CellAddress("CA1").Row);

            Assert.Equal(701, new CellAddress("ZZ1").Column);
            Assert.Equal(0, new CellAddress("ZZ1").Row);

            Assert.Equal(702, new CellAddress("AAA1").Column);
            Assert.Equal(0, new CellAddress("AAA1").Row);

            Assert.Equal(727, new CellAddress("AAZ1").Column);
            Assert.Equal(0, new CellAddress("AAZ1").Row);

            Assert.Equal(728, new CellAddress("ABA1").Column);
            Assert.Equal(0, new CellAddress("ABA1").Row);

            Assert.Equal(16383, new CellAddress("XFD1").Column);
            Assert.Equal(0, new CellAddress("XFD1").Row);

            Assert.Equal(728, new CellAddress("ABA1000").Column);
            Assert.Equal(999, new CellAddress("ABA1000").Row);
        }
    }
}
