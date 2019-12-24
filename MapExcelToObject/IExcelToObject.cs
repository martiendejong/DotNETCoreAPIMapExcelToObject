using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelToObject
{
    public interface IExcelToObject
    {
        DataSet GetDataSet(Stream xlsxStream);

        IEnumerable<MappedObjectType> MapToObjects<MappedObjectType>(Stream xlsxStream);
    }
}
