using System.Data;
using System.IO;

namespace TestExcelDataReader
{
    class Table
    {
        public FileInfo FileInfo { get; set; }
        public DataTable DataTable { get; set; }
        public Table(FileInfo fileInfo, DataTable dataTable)
        {
            FileInfo = fileInfo;
            DataTable = dataTable;
        }
    }
}
