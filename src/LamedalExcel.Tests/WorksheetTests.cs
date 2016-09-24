using System;
using Xunit;

namespace LamedalExcel.Tests
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
    }
}
