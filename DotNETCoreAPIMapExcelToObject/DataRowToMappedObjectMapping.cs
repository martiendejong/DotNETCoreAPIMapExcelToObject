using AutoMapper;
using DotNETCoreAPIMapExcelToObject.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DotNETCoreAPIMapExcelToObject
{
    public class DataRowToMappedObjectMapping : Profile
    {
        public DataRowToMappedObjectMapping()
        {
            CreateMap<DataSet, DataTable>().ConstructUsing(dataSet => dataSet.Tables["Blad1"]);
            CreateMap<DataRow, MappedObject>()
                .ForMember(mappedObject => mappedObject.Text, config => config.MapFrom(row => row["Tekst"]))
                .ForMember(mappedObject => mappedObject.Number, config => config.MapFrom(row => row["Getal"]))
                .ForMember(mappedObject => mappedObject.DecimalNumber, config => config.MapFrom(row => row["Kommagetal"]))
                .ForMember(mappedObject => mappedObject.Date, config => config.MapFrom(row => row["Datum"]))
                .ForMember(mappedObject => mappedObject.Id, config => config.MapFrom(row => row["Guid"]));
        }
    }
}
