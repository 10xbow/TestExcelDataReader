using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace TestExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            var paths = new List<string>();
            if (args.Count() == 0) paths.Add(AppDomain.CurrentDomain.BaseDirectory);
            paths.AddRange(args);

            var queue = new List<FileInfo>();

            foreach (var path in paths)
            {
                if (File.Exists(path)) queue.Add(new FileInfo(path));
                if (Directory.Exists(path))
                {
                    Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)
                             .ToList()
                             .ForEach(f => queue.Add(new FileInfo(f)));
                }
            }
            var targetExtensions = new List<string> { ".xls", ".xlsx", ".xlsm" };
            var excelFiles = queue.Where(q => targetExtensions.Contains(q.Extension.ToLower())).ToList();

            excelFiles.ForEach(e => Console.WriteLine(e.FullName));

            MakeExcelFile(Concat(GetTables(excelFiles)));

            Console.WriteLine("Complete");
            Console.ReadLine();
        }

        static List<Table> GetTables(IReadOnlyCollection<FileInfo> files)
        {
            var tables = new List<Table>();
            foreach (var file in files)
            {
                using (var stream = File.Open(file.FullName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        reader.AsDataSet()
                              .Tables
                              .Cast<DataTable>()
                              .ToList()
                              .ForEach(dt => tables.Add(new Table(file, dt)));
                    }
                }
            }
            return tables;
        }

        static DataTable Concat(IReadOnlyCollection<Table> tables)
        {
            var resultTable = new DataTable("Sheet1");
            var extraColumns = new DataColumn[]
            {
                new DataColumn { ColumnName = "ファイル名", DataType = typeof(string) },
                new DataColumn { ColumnName = "シート名", DataType = typeof(string) },
                new DataColumn { ColumnName = "行番号", DataType = typeof(int) },
            };

            resultTable.Columns.AddRange(extraColumns);

            for (int i = 1; i <= tables.Max(m => m.DataTable.Columns.Count); i++)
            {
                resultTable.Columns.Add($"列{i}");
            }

            foreach (var table in tables)
            {
                var fileName = table.FileInfo.Name;
                var sheetName = table.DataTable.TableName;
                foreach (DataRow row in table.DataTable.Rows)
                {
                    var newRow = resultTable.NewRow();
                    var extraData = new List<object> { fileName, sheetName, table.DataTable.Rows.IndexOf(row) + 1 };
                    var origin = row.ItemArray.ToList();
                    newRow.ItemArray = extraData.Concat(origin).ToArray();
                    resultTable.Rows.Add(newRow);
                }
            }
            return resultTable;
        }

        static void MakeExcelFile(DataTable dataTable)
        {
            var file = new FileInfo($"{DateTime.Now:yyyy-MM-dd_hhmmss}.xlsx");

            using (ExcelPackage pck = new ExcelPackage(file))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                pck.Save();
            }
        }
    }
}
