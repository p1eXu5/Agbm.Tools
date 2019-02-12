using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using Agbm.NpoiExcel.Tests.Factory;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Agbm.NpoiExcel.Tests.IntegrationalTests
{
    [TestFixture]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Local" ) ]
    public class ExcelImporterIntegrationalTests
    {
        [SetUp]
        public void SetupCulture()
        {
            CultureInfo.CurrentUICulture = new CultureInfo( "en-us" );
        }



        [Test]
        public void ImportData_ByDefault_ReturnsNotNullSheetTable()
        {
            // Arrange:
            var stream = GetExcelMemoryStream();

            // Action:
            var sheetTable = ExcelImporter.GetSheetTable( stream );

            // Assert:
            Assert.That( sheetTable, Is.Not.Null );
        }



        [ Test ]
        public void GetDataFromTable_ComplexTInt_RetutnsExpected ()
        {
            var typeRepository = GetTypeRepository();
            var typeConverter = GetTypeConverter();

            var sheet = MockedSheetFactory.GetMockedSheet( typeof( TestType ), typeRepository, TestType.Data );
            var sheetTable = new SheetTable( sheet );

            var typeMap = typeRepository.GetTypeAndPropertyMap( sheetTable );

            var actual = ExcelImporter.GetDataFromTable( sheetTable, typeMap.propertyMap, typeConverter ).ToArray();

            Assert.That( actual[0], Is.EqualTo( TestType.Expected ) );
        }



        #region Factory

        private ITypeRepository GetTypeRepository ()
        {
            var repo = new TypeRepository();
            repo.RegisterType< TestType >();

            return repo;
        }

        private ITypeConverter< TestType, TestType > GetTypeConverter ()
        {
            var conv = new TestTypeConverter();
            return conv;
        }

        private Stream GetExcelMemoryStream()
        {
            var book = new XSSFWorkbook();

            var row = book.CreateSheet ("0").CreateRow (0);

            row.CreateCell(0).SetCellValue("Some data");
            row.CreateCell(1).SetCellValue("Some data");

            var stream = new MemoryStream();
            book.Write (stream, true);
            stream.Position = 0;

            return stream;
        }



        class TestType
        {
            public int IntType { get; set; }
            public int? NullableIntType { get; set; }

            public double DoubleType { get; set; }
            public double? NullableDoubleType { get; set; }

            public bool BooleanType { get; set; }
            public bool? NullableBooleanType { get; set; }

            public DateTime DateTime { get; set; }
            public DateTime? NullableDateTime { get; set; }

            public static object[] Data { get; } = new object[] { 0, "", 0.0, "", true, "", new DateTime( 2019, 2, 8, 12, 8, 0 ), "" };

            public static readonly TestType Expected= new TestType {
                IntType = 0,
                NullableIntType = null,
                DoubleType = 0.0,
                NullableDoubleType = null,
                BooleanType = true,
                NullableBooleanType = null,
                DateTime = new DateTime( 2019, 2, 8, 12, 8, 0 ),
                NullableDateTime = null,
            };

            public override bool Equals ( object obj )
            {
                if ( ReferenceEquals( obj, null ) ) return false;
                return obj is TestType other && Equals( other );
            }

            public override string ToString ()
            {
                return $"\n\t{nameof(IntType),-20} = {IntType,-6}\n" +
                       $"\t{nameof(NullableIntType),-20} = {NullableIntType?.ToString() ?? "null",-6}\n" +
                       $"\t{nameof(DoubleType),-20} = {DoubleType,-6}\n" +
                       $"\t{nameof(NullableDoubleType),-20} = {NullableDoubleType?.ToString() ?? "null",-6}\n" +
                       $"\t{nameof(BooleanType),-20} = {BooleanType,-6}\n" +
                       $"\t{nameof(NullableBooleanType),-20} = {NullableBooleanType?.ToString() ?? "null",-6}\n" +
                       $"\t{nameof(DateTime),-20} = {DateTime,-6}\n" +
                       $"\t{nameof(NullableDateTime),-20} = {NullableDateTime?.ToString() ?? "null",-6}\n";
            }

            bool Equals ( TestType obj )
            {
                return IntType == obj.IntType
                       && NullableIntType.Equals( obj.NullableIntType )
                       && DoubleType.Equals( obj.DoubleType )
                       && NullableDoubleType.Equals( obj.NullableDoubleType )
                       && BooleanType == obj.BooleanType
                       && NullableBooleanType.Equals( obj.NullableBooleanType )
                       && DateTime.Equals( obj.DateTime )
                       && NullableDateTime.Equals( obj.NullableDateTime );
            }

            public override int GetHashCode ()
            {
                throw new NotImplementedException();
            }
        }

        class TestTypeConverter : ITypeConverter< TestType, TestType >
        {
            public TestType Convert ( TestType obj )
            {
                return obj;
            }
        }

        #endregion
    }
}
