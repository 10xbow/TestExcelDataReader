using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Data;

namespace TestExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 0) return;
            using (var stream = File.Open(args[0], FileMode.Open, FileAccess.Read))
            {
                var fileInfo = new FileInfo(args[0]);
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    var tables = new List<DataTable>();
                    foreach (DataTable table in result.Tables)
                    {
                        tables.Add(table);
                    }

                    var dataTable = new DataTable("x");

                    var columnCount = tables.Select(t => t.Columns.Count).Max();

                    for (int i = 1; i <= columnCount; i++)
                    {
                        switch (i)
                        {
                            case 1:
                                dataTable.Columns.Add("FileName");
                                break;
                            case 2:
                                dataTable.Columns.Add("SheetName");
                                break;
                            case 3:
                                dataTable.Columns.Add("RowNumber");
                                break;
                            default:
                                dataTable.Columns.Add($"Columns{i}");
                                break;
                        }
                    }

                    foreach (var table in tables)
                    {
                        foreach (DataRow row in table.Rows)
                        {

                            var newRow = dataTable.NewRow();
                            object[] extend = { fileInfo.Name, table.TableName, table.Rows.IndexOf(row) + 1 };

                            Console.WriteLine($"{columnCount} {row.ItemArray.Length.ToString()}");

                            //dataTable.Rows.Add(newRow);
                            //var line = new List<string> {fileInfo.Name, table.TableName, rowIndex.ToString()};
                            //line.AddRange(row.ItemArray.Select(i => i.ToString()));
                            //Console.WriteLine(string.Join(",",line));
                        }
                    }
                }
            }

            Console.ReadLine();
        }
    }
}
