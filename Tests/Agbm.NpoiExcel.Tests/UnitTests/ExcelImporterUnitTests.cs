using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Agbm.Helpers.Extensions;
using Agbm.NpoiExcel.Tests.Factory;
using NUnit.Framework;

namespace Agbm.NpoiExcel.Tests.UnitTests
{
    [TestFixture]
    public class ExcelImporterUnitTests
    {
        [SetUp]
        public void SetupCulture()
        {
            CultureInfo.CurrentUICulture = new CultureInfo ("en-Us", false);
        }

        #region GetSheetTable With FileName

        [Test]
        public void ImportData_FileNameIsNull_Throws()
        {
            var type = TypeExcelFactory.EmptyClass;
            string path = null;
            var ex = Assert.Catch<ArgumentException> (() => ExcelImporter.ImportData(path, type, 0));

            StringAssert.Contains ("Path cannot be null.", ex.Message);
        }

        [Test]
        public void ImportData_FileNameIsEmptyString_Throws()
        {
            var type = TypeExcelFactory.EmptyClass;
            var ex = Assert.Catch<ArgumentException>(() => ExcelImporter.ImportData(string.Empty, type, 0));

            StringAssert.Contains("Empty path name is not legal.", ex.Message);
        }

        [Test]
        public void ImportData_FileNameIsWhitespaces_Throws()
        {
            var type = TypeExcelFactory.EmptyClass;
            var ex = Assert.Catch<ArgumentException>(() => ExcelImporter.ImportData("  ", type, 0));

            StringAssert.Contains("The path is not of a legal form.", ex.Message);
        }

        [Test]
        public void ImportData_FileNameConsistsOfExtensionOnly_ReturnsEmptyCollection()
        {
            var type = TypeExcelFactory.EmptyClass;
            Assert.Catch<FileNotFoundException>(  () => ExcelImporter.ImportData(".xlsx", type, 0) );
        }

        [Test]
        public void ImportData_FileDoesntExist_Throw()
        {
            var type = TypeExcelFactory.EmptyClass;
            Assert.Catch<FileNotFoundException>(  () => ExcelImporter.ImportData("notexistedfile.xlsx", type, 0));
        }

        [Test]
        public void ImportData_FileDoesNotContainExcelData_ReturnsEmptyCollection()
        {
            var file = "test.txt".AppendAssemblyPath();

            var stream = File.Create (file);
            stream.Close();

            var type = TypeExcelFactory.EmptyClass;
            var res = ExcelImporter.ImportData(file, type, 0).Cast< Empty >();

            Assert.That( res, Is.Empty );

            File.Delete (file);
        }

        #endregion


        #region GetSheetTable With Stream

        [Test]
        public void ImportData_TypeIsNull_Throws()
        {
            var stream = new MemoryStream(new byte[1]);
            var ex = Assert.Catch<ArgumentNullException>(() => ExcelImporter.ImportData(stream, null, 0));

            StringAssert.Contains("Type cannot be null.", ex.Message);
        }

        [Test]
        public void ImportData_TypeWithoutPublicParameterlessCtor_Throws()
        {
            var type = TypeExcelFactory.ClassWithParameterizedCtor;
            var stream = StreamFactory.GetExcelMemoryStream();
            var ex = Assert.Catch<TypeAccessException> (() => ExcelImporter.ImportData (stream, type, 0));

            StringAssert.Contains("has no public parameterless constructor!", ex.Message);
        }

        [Test]
        public void ImportData_SheetIndexIsNegative_Throws()
        {
            var type = TypeExcelFactory.EmptyClass;
            var stream = new MemoryStream(new byte[1]);
            var ex = Assert.Catch<ArgumentException>(() => ExcelImporter.ImportData(stream, type, -1));

            StringAssert.Contains("Index of sheet cannot be less than zero.", ex.Message);
        }

        [Test]
        public void ImportData_StreamDoesNotContainExcelData_ReturnsEmptyCollection()
        {
            var type = TypeExcelFactory.EmptyClass;
            var stream = new MemoryStream(new byte[1]);
            var res = ExcelImporter.ImportData(stream, type, 0).Cast< Empty >();

            stream.Close();
            Assert.That( res, Is.Empty );
        }

        #endregion
    }
}
