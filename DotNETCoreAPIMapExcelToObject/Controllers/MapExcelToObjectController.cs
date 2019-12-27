using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DotNETCoreAPIMapExcelToObject.Models;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using ExcelToObject;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DotNETCoreAPIMapExcelToObject.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class MapExcelToObjectController : ControllerBase
    {
        private readonly IExcelToObject<MappedObject> ExcelToMappedObject;

        private readonly IExcelToObject<AnotherMappedObject> ExcelToAnoterMappedObject;

        public MapExcelToObjectController(IExcelToObject<MappedObject> excelToMappedObject, IExcelToObject<AnotherMappedObject> excelToAnotherMappedObject)
        {
            ExcelToMappedObject = excelToMappedObject;
            ExcelToAnoterMappedObject = excelToAnotherMappedObject;
        }

        [HttpPost]
        [Route("Object")]
        public IActionResult MapFileToObject(IFormFile file)
        {
            return HandleMapFileRequest(file, stream => ExcelToMappedObject.MapToObjects(stream));
        }        

        [HttpPost]
        [Route("AnotherObject")]
        public IActionResult MapFileToAnotherObject(IFormFile file)
        {
            return HandleMapFileRequest(file, stream => ExcelToAnoterMappedObject.MapToObjects(stream));
        }

        private IActionResult HandleMapFileRequest<MappedObjectType>(IFormFile file, Func<Stream, IEnumerable<MappedObjectType>> mapFunction)
        {
            if (file.Length == 0)
            {
                return BadRequest("File is empty");
            }
            if (!file.FileName.EndsWith(".xlsx"))
            {
                return BadRequest("File extension must be xlsx");
            }

            try
            {
                return Ok(DoMapFile(file, mapFunction));
            }
            catch (ExcelToObjectException exception)
            {
                return BadRequest(exception.Message);
            }
        }

        private IEnumerable<MappedObjectType> DoMapFile<MappedObjectType>(IFormFile file, Func<Stream, IEnumerable<MappedObjectType>> mapFunction)
        {
            IEnumerable<MappedObjectType> resultObjects;
            using (var stream = file.OpenReadStream())
            {
                resultObjects = mapFunction(stream);
            }

            return resultObjects;
        }
    }
}