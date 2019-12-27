using AutoMapper;
using DotNETCoreAPIMapExcelToObject;
using DotNETCoreAPIMapExcelToObject.Models;
using ExcelToObject;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MapExcelToObjectUnitTests
{
    [TestClass]
    public class ExcelToObjectConfigFileMapperTests
    {
        private readonly IExcelToObject<MappedObject> Mapper = new ExcelToObjectConfigFileMapper<MappedObject>("ExcelToObjectMapperTestFiles");

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\goed.xlsx", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ReturnANonEmptyList_When_FileIsCorectAndHasRecords()
        {
            IEnumerable<MappedObject> result = Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\goed.xlsx"));

            Assert.IsTrue(result.Count() > 0);
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\datum is geen datumveld.xlsx", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_MappedValueIsNotADate()
        {
            Assert.ThrowsException<ExcelToObjectException>(() => Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\datum is geen datumveld.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\kolommen tekst en getal bestaan niet.xlsx", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_ColumnsAreMissing()
        {
            Assert.ThrowsException<ExcelToObjectException>(() => Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\kolommen tekst en getal bestaan niet.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\sheet Blad1 bestaat niet.xlsx", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_SheetIsMissing()
        {
            Assert.ThrowsException<ExcelToObjectException>(() => Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\sheet Blad1 bestaat niet.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\tekstbestand hernoemd naar.xlsx", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_FileIsNotXLSXFormat()
        {
            Assert.ThrowsException<ExcelToObjectException>(() => Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\tekstbestand hernoemd naar.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\tekstbestand.txt", "ExcelToObjectMapperTestFiles")]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\MappedObject.mapping", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_FileExtensionIsNotXLSX()
        {
            Assert.ThrowsException<ExcelToObjectException>(() => Mapper.MapToObjects(File.OpenRead(@"ExcelToObjectMapperTestFiles\tekstbestand.txt")));
        }
    }
}
