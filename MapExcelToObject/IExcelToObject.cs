using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelToObject
{
    public interface IExcelToObject<MappedObjectType>
    {
        DataSet GetDataSet(Stream xlsxStream);

        IEnumerable<MappedObjectType> MapToObjects(Stream xlsxStream);
    }
}
