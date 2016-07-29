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
        private const char Comma = ',';
        private const string CsvExtension = ".csv";
        private const string DateTimeFormat = "yyyyMMddTHHmmss";
        private const string DcPrefix = "DC=";
        private const string DoubleHyphen = "--";
        private const char Hyphen = '-';
        private const string OuPrefix = "OU=";

        private static readonly string OutputDirectoryPath = Path.Combine(
            GetFolderPath(MyDocuments), "ActiveDirectoryTool");

        public DataView Data { get; set; }
        public QueryType QueryType { get; set; }
        public Scope Scope { get; set; }

        public string WriteToCsv()
        {
            string fileName;
            if (Scope == null)
            {
                fileName = QueryType + Hyphen +
                           DateTime.Now.ToString(DateTimeFormat) +
                           CsvExtension;
            }
            else
            {
                fileName = Scope.Context.Remove(OuPrefix).Remove(DcPrefix)
                    .Replace(Comma, Hyphen) + DoubleHyphen + QueryType +
                           Hyphen + DateTime.Now.ToString(DateTimeFormat) +
                           CsvExtension;
            }
            var fullFileName = Path.Combine(OutputDirectoryPath, fileName);
            var dataTable = Data.ToTable();
            var stringBuilder = new StringBuilder();

            var columnNames = dataTable.Columns.Cast<DataColumn>().Select(
                column => column.ColumnName);
            stringBuilder.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dataTable.Rows)
            {
                var fields = row.ItemArray.Select(
                    field => string.Concat(
                        "\"", field.ToString().Replace("\"", "\"\""), "\""));
                stringBuilder.AppendLine(string.Join(",", fields));
            }

            Directory.CreateDirectory(OutputDirectoryPath);
            File.WriteAllText(fullFileName, stringBuilder.ToString());
            Process.Start(fullFileName);
            return fullFileName;
        }
    }
}