using System.Drawing;
using System.IO;
using NUnit.Framework;

namespace Agbm.NpoiExcel.Tests.IntegrationalTests
{
    [ TestFixture ]
    class ExcelExporterTests
    {
        [ Test ]
        public void ExportData_ColorCellWithString_CreatesFile ()
        {
            var file = @"d:\test.xlsx";
            var cells = new[] {
                new ExportingCell( "Aquamarine", 0, 0, Color.LightSkyBlue ),
                new ExportingCell( "LightGreen", 1, 0, Color.LightGreen ),
                new ExportingCell( "Custom", 2, 0, new byte[] { 0xc5, 0xe0, 0xb4 } )
            };

            ExcelExporter.ExportData( file, cells );

            Assert.That( File.Exists( file ) );
        }
    }
}
