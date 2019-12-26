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
        private readonly IExcelToObject ExcelToDataSet;

        public MapExcelToObjectController(IExcelToObject excelToDataSet)
        {
            ExcelToDataSet = excelToDataSet;
        }

        [HttpPost]
        public IActionResult MapFile(IFormFile file)
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
                return Ok(DoMapFile(file));
            }
            catch (ExcelToObjectException exception)
            {
                return BadRequest(exception.Message);
            }
        }

        private IEnumerable<MappedObject> DoMapFile(IFormFile file)
        {
            IEnumerable<MappedObject> resultObjects;
            using (var stream = file.OpenReadStream())
            {
                resultObjects = ExcelToDataSet.MapToObjects<MappedObject>(stream);
            }

            return resultObjects;
        }
    }
}