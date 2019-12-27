using AutoMapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToObject
{

    public class ConfigurableMapping<T> : Profile
    {
        public ConfigurableMapping(string sheetName, IEnumerable<MappedField> fields)
        {
            CreateMap<DataSet, DataTable>().ConstructUsing(dataSet => dataSet.Tables[sheetName]);
            IMappingExpression<DataRow, T> mappingExpression = CreateMap<DataRow, T>();
            foreach(MappedField field in fields)
            {
                mappingExpression = mappingExpression.ForMember(field.DestinationName, config => config.MapFrom(row => row[field.SourceName]));
            }
        }
    }
}
