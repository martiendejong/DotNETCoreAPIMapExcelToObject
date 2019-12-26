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
    public class ExcelToObjectMapperTests
    {
        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\goed.xlsx", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ReturnANonEmptyList_When_FileIsCorectAndHasRecords()
        {
            ExcelToObjectMapper mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));
            IEnumerable<MappedObject> result = mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\goed.xlsx"));

            Assert.IsTrue(result.Count() > 0);
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\datum is geen datumveld.xlsx", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_MappedValueIsNotADate()
        {
            IExcelToObject mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));

            Assert.ThrowsException<ExcelToObjectException>(() => mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\datum is geen datumveld.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\kolommen tekst en getal bestaan niet.xlsx", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_ColumnsAreMissing()
        {
            ExcelToObjectMapper mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));

            Assert.ThrowsException<ExcelToObjectException>(() => mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\kolommen tekst en getal bestaan niet.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\sheet Blad1 bestaat niet.xlsx", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_SheetIsMissing()
        {
            ExcelToObjectMapper mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));

            Assert.ThrowsException<ExcelToObjectException>(() => mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\sheet Blad1 bestaat niet.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\tekstbestand hernoemd naar.xlsx", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_FileIsNotXLSXFormat()
        {
            ExcelToObjectMapper mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));

            Assert.ThrowsException<ExcelToObjectException>(() => mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\tekstbestand hernoemd naar.xlsx")));
        }

        [TestMethod]
        [DeploymentItem(@"ExcelToObjectMapperTestFiles\tekstbestand.txt", "ExcelToObjectMapperTestFiles")]
        public void MapToObjects_Should_ThrowAnExcelToObjectException_When_FileExtensionIsNotXLSX()
        {
            ExcelToObjectMapper mapper = new ExcelToObjectMapper(new Mapper(CreateConfiguration()));

            Assert.ThrowsException<ExcelToObjectException>(() => mapper.MapToObjects<MappedObject>(File.OpenRead(@"ExcelToObjectMapperTestFiles\tekstbestand.txt")));
        }

        // Automapper configuration
        private MapperConfiguration CreateConfiguration()
        {
            var config = new MapperConfiguration(cfg =>
            {
                // Add all profiles in current assembly
                cfg.AddProfiles(new Profile[]
                {
                    new DataRowToMappedObjectMapping()
                });
            });

            return config;
        }
    }
}
