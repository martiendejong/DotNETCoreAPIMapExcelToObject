using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObject
{
    public class ExcelToObjectException : Exception
    {
        public ExcelToObjectException(string message) : base(message)
        {
        }
    }
}
