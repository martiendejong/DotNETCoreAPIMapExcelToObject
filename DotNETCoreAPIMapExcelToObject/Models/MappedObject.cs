using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DotNETCoreAPIMapExcelToObject.Models
{
    public class MappedObject
    {
        public string Text { get; set; }

        public int Number { get; set; }

        public decimal DecimalNumber { get; set; }

        public DateTime Date { get; set; }

        public Guid Id { get; set; }
    }
}
