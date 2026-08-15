using System.IO;
using ClosedXML.Excel;

namespace UnitTests
{
    static class TestFiles
    {
        static readonly string directory = Path.Combine(Path.GetTempPath(), "CsExcelTests");

        static TestFiles() => Directory.CreateDirectory(directory);

        public static string PathFor(string filename) => Path.Combine(directory, filename);

        public static XLWorkbook Open(string filename) => new(PathFor(filename));
    }
}
