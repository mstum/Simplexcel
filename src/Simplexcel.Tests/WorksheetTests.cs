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

        [Fact]
        public void Worksheet_PopulatePersonClass_TwoColumns()
        {
            var ws = new Worksheet("Foo");
            ws.Populate(new[] { new PersonClass { FirstName = "Michael", LastName = "Stum" } });
            Assert.Equal(2, ws.Cells.ColumnCount);
            Assert.Equal("FirstName", ws["A1"].Value);
            Assert.Equal("LastName", ws["B1"].Value);
        }

        [Fact]
        public void Worksheet_PopulatePersonRecord_TwoColumns()
        {
            var ws = new Worksheet("Foo");
            ws.Populate(new[] { new PersonRecord(FirstName: "Michael", LastName: "Stum") });
            Assert.Equal(2, ws.Cells.ColumnCount);
            Assert.Equal("FirstName", ws["A1"].Value);
            Assert.Equal("LastName", ws["B1"].Value);
        }

        private class PersonClass
        {
            public string FirstName { get; init; }
            public string LastName { get; init; }
        }

        private record PersonRecord(string FirstName, string LastName);
    }
}
