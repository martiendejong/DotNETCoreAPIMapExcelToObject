using AutoMapper;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelToObject
{
    public class ExcelToObjectMapper<MappedObjectType> : IExcelToObject<MappedObjectType>
    {
        private static readonly ExcelDataSetConfiguration Configuration = new ExcelDataSetConfiguration()
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

        public IEnumerable<MappedObjectType> MapToObjects(Stream xlsxStream)
        {
            try
            {
                DataSet dataSet = GetDataSet(xlsxStream);
                DataTable dataTable = Mapper.Map<DataTable>(dataSet);
                IEnumerable<MappedObjectType> mappedObjects = Mapper.Map<IEnumerable<MappedObjectType>>(dataTable.AsEnumerable());
                return mappedObjects;
            }
            catch (AutoMapperMappingException mappingException)
            {
                throw TranslateException(mappingException);
            }
            catch (ExcelReaderException excelException)
            {
                throw new ExcelToObjectException(excelException.Message);
            }
        }

        #region exception handling

        /// <summary>
        /// Translates the AutoMapperMappingException to a ExcelToObjectMapperException
        /// </summary>
        /// <param name="mappingException">The exception to be translated</param>
        /// <returns>The translated exception.</returns>
        private ExcelToObjectException TranslateException(AutoMapperMappingException mappingException)
        {
            AutoMapperMappingException memberMapException = GetMemberMapException(mappingException);
            Exception regularException = GetRegularException(memberMapException);
            if (memberMapException.TypeMap.SourceType == typeof(DataSet)
                && memberMapException.TypeMap.CustomCtorExpression != null)
            {
                return new ExcelToObjectException($"Cannot map {memberMapException.TypeMap.CustomCtorExpression.Body} to a DataTable: {regularException.Message}");
            }
            else if (memberMapException.TypeMap.SourceType == typeof(DataRow)
                && memberMapException.MemberMap.CustomMapExpression.Body != null)
            {
                return new ExcelToObjectException($"Cannot map {memberMapException.MemberMap.CustomMapExpression.Body} to {memberMapException.MemberMap.DestinationName}: {regularException.Message}");
            }
            return new ExcelToObjectException(memberMapException.Message);
        }

        /// <summary>
        /// Loop through the innerexceptions and get the first one which is not a AutoMapperMappingException instance.
        /// </summary>
        /// <param name="memberMapException">The high-level exception.</param>
        /// <returns>The inner exception that is not a  instance if any.</returns>
        private static Exception GetRegularException(AutoMapperMappingException memberMapException)
        {
            Exception underlyingException = memberMapException;
            while (underlyingException is AutoMapperMappingException && (underlyingException.InnerException != null))
            {
                underlyingException = underlyingException.InnerException;
            }
            return underlyingException;
        }

        /// <summary>
        /// Loop through the innerexceptions and get the first one with a IMemberMap instance.
        /// </summary>
        /// <param name="mappingException">The high-level exception.</param>
        /// <returns>The inner exception with a IMemberMap instance if any.</returns>
        private static AutoMapperMappingException GetMemberMapException(AutoMapperMappingException mappingException)
        {
            AutoMapperMappingException memberMapException = mappingException;
            while (memberMapException.MemberMap == null && memberMapException.InnerException is AutoMapperMappingException)
            {
                memberMapException = memberMapException.InnerException as AutoMapperMappingException;
            }
            return memberMapException;
        }

        #endregion
    }
}
