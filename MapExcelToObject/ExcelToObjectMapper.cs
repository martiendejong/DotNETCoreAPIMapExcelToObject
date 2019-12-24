using AutoMapper;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelToObject
{
    public class ExcelToObjectMapper : IExcelToObject
    {
        private readonly ExcelDataSetConfiguration Configuration = new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        };

        private readonly IMapper Mapper;

        public ExcelToObjectMapper(IMapper mapper)
        {
            Mapper = mapper;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public DataSet GetDataSet(Stream xlsxStream)
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(xlsxStream);
            return excelReader.AsDataSet(Configuration);
        }

        public IEnumerable<MappedObjectType> MapToObjects<MappedObjectType>(Stream xlsxStream)
        {
            DataSet dataSet;
            try
            {
                dataSet = GetDataSet(xlsxStream);
            }
            catch(Exception getDataSetException)
            {
                throw;
                //throw new GetDataSetException(getDataSetException);
            }

            IEnumerable<MappedObjectType> mappedObjects;
            try
            {
                DataTable dataTable = Mapper.Map<DataTable>(dataSet);
                mappedObjects = Mapper.Map<IEnumerable<MappedObjectType>>(dataTable.AsEnumerable());
            }
            catch(Exception mapException)
            {
                throw;// new 
            }

            return mappedObjects;
        }
    }
}
