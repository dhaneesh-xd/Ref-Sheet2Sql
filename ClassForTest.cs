using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    public class ClassForTest
    {
        public string GetConnectionString(string filePath)
        {
            string extension = System.IO.Path.GetExtension(filePath);
            if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;\"";
            }
            else if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            }
            else
            {
                throw new NotSupportedException("File extension is not supported.");
            }
        }
    }
}
