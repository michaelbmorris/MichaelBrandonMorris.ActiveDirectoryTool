using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Extensions.PrimitiveExtensions;
using static System.Environment;
using static System.Environment.SpecialFolder;

namespace ActiveDirectoryTool
{
    public class DataFileWriter
    {
        private const string DateTimeFormat = "yyyyMMddTHHmmss";

        private readonly string _outputPath = Path.Combine(
            GetFolderPath(MyDocuments), "ActiveDirectoryTool");

        public DataView Data { get; set; }
        public QueryType QueryType { get; set; }
        public string Scope { get; set; }

        public string WriteToCsv()
        {
            var fileName = Scope.Remove("OU=").Remove("DC=").Replace(',', '-')
                           + "--" + QueryType + '-' +
                           DateTime.Now.ToString(DateTimeFormat) + ".csv";
            var fullFileName = Path.Combine(_outputPath, fileName);
            var dataTable = Data.ToTable();
            var stringBuilder = new StringBuilder();

            var columnNames = dataTable.Columns.Cast<DataColumn>().Select(
                column => column.ColumnName);
            stringBuilder.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dataTable.Rows)
            {
                var fields = row.ItemArray.Select(field => string.Concat(
                        "\"", field.ToString().Replace("\"", "\"\""), "\""));
                stringBuilder.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(fullFileName, stringBuilder.ToString());
            Process.Start(fullFileName);
            return fullFileName;
        }
    }
}