using System;
using Xunit;

namespace Simplexcel.Tests
{
    public class WorksheetTests
    {
        [Fact]
        public void Worksheet_NameTooLong_ThrowsArgumentException()
        {
            var name = new string('x', Worksheet.MaxSheetNameLength + 1);

            Assert.Throws<ArgumentException>(() => new Worksheet(name));
        }

        [Fact]
        public void Worksheet_NameInvalidChars_ThrowsArgumentException()
        {
            foreach (var invalid in Worksheet.InvalidSheetNameChars)
            {
                var name = "x" + invalid + "y";
                Assert.Throws<ArgumentException>(() => new Worksheet(name));
            }
        }

        [Fact]
        public void Worksheet_EmptyName_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => new Worksheet(string.Empty));
        }

        [Fact]
        public void Worksheet_NullName_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => new Worksheet(null));
        }

        [Fact]
        public void FreezeTopLeft_NegativeRows_Throws()
        {
            var ws = new Worksheet("Foo");
            Assert.Throws<ArgumentOutOfRangeException>("rows", () => ws.FreezeTopLeft(-1, 9));
        }

        [Fact]
        public void FreezeTopLeft_NegativeColumns_Throws()
        {
            var ws = new Worksheet("Foo");
            Assert.Throws<ArgumentOutOfRangeException>("columns", () => ws.FreezeTopLeft(9, -1));
        }
    }
}
