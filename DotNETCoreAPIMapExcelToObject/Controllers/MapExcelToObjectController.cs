using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DotNETCoreAPIMapExcelToObject.Models;
using ExcelDataReader;
using ExcelToObject;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DotNETCoreAPIMapExcelToObject.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class MapExcelToObjectController : ControllerBase
    {
        private readonly IMapper Mapper;

        private readonly IExcelToObject ExcelToDataSet;

        public MapExcelToObjectController(IMapper mapper, IExcelToObject excelToDataSet)
        {
            Mapper = mapper;
            ExcelToDataSet = excelToDataSet;
        }

        [HttpPost]
        public async Task<IActionResult> MapFileAsync(IFormFile file)
        {
            IActionResult result;
            
            if (file.Length == 0)
            {
                result = BadRequest("File is empty");
            }
            else if (!file.FileName.EndsWith(".xlsx"))
            {
                result = BadRequest("File extension must be xlsx");
            }
            else
            {
                try
                {
                    IEnumerable<MappedObject> resultObjects;
                    using (var stream = file.OpenReadStream())
                    {
                        resultObjects = ExcelToDataSet.MapToObjects<MappedObject>(stream);
                    }

                    /*DataSet dataSet;

                    using (var stream = file.OpenReadStream())
                    {
                        //dataSet = ExcelToDataSet.GetDataSet(stream);

                        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        ExcelDataSetConfiguration config = new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        };

                        DataSet dataSet = excelReader.AsDataSet(config);
                        IEnumerable<MappedObject> resultObjects = Mapper.Map<IEnumerable<MappedObject>>(dataSet);
                    }
                    IEnumerable<MappedObject> resultObjects = Mapper.Map<IEnumerable<MappedObject>>(dataSet);*/

                    result = Ok(resultObjects);
                }
                catch (Exception exception)
                {
                    throw;
                }
            }

            return result;
        }
    }
}