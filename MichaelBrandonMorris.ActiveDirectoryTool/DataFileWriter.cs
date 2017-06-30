using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class DataFileWriter.
    /// </summary>
    /// TODO Edit XML Comment Template for DataFileWriter
    public class DataFileWriter
    {
        /// <summary>
        ///     The comma
        /// </summary>
        /// TODO Edit XML Comment Template for Comma
        private const char Comma = ',';

        /// <summary>
        ///     The CSV extension
        /// </summary>
        /// TODO Edit XML Comment Template for CsvExtension
        private const string CsvExtension = ".csv";

        /// <summary>
        ///     The date time format
        /// </summary>
        /// TODO Edit XML Comment Template for DateTimeFormat
        private const string DateTimeFormat = "yyyyMMddTHHmmss";

        /// <summary>
        ///     The dc prefix
        /// </summary>
        /// TODO Edit XML Comment Template for DcPrefix
        private const string DcPrefix = "DC=";

        /// <summary>
        ///     The double hyphen
        /// </summary>
        /// TODO Edit XML Comment Template for DoubleHyphen
        private const string DoubleHyphen = "--";

        /// <summary>
        ///     The hyphen
        /// </summary>
        /// TODO Edit XML Comment Template for Hyphen
        private const char Hyphen = '-';

        /// <summary>
        ///     The ou prefix
        /// </summary>
        /// TODO Edit XML Comment Template for OuPrefix
        private const string OuPrefix = "OU=";

        /// <summary>
        ///     The output directory path
        /// </summary>
        /// TODO Edit XML Comment Template for OutputDirectoryPath
        private static readonly string OutputDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ActiveDirectoryTool");

        /// <summary>
        ///     Gets or sets the data.
        /// </summary>
        /// <value>The data.</value>
        /// TODO Edit XML Comment Template for Data
        public DataView Data
        {
            get;
            set;
        }

        /// <summary>
        ///     Gets or sets the type of the query.
        /// </summary>
        /// <value>The type of the query.</value>
        /// TODO Edit XML Comment Template for QueryType
        public QueryType QueryType
        {
            get;
            set;
        }

        /// <summary>
        ///     Gets or sets the scope.
        /// </summary>
        /// <value>The scope.</value>
        /// TODO Edit XML Comment Template for Scope
        public Scope Scope
        {
            get;
            set;
        }

        /// <summary>
        ///     Writes to CSV.
        /// </summary>
        /// <returns>System.String.</returns>
        /// TODO Edit XML Comment Template for WriteToCsv
        public string WriteToCsv()
        {
            string fileName;
            if (Scope == null)
            {
                fileName = QueryType
                           + Hyphen
                           + DateTime.Now.ToString(DateTimeFormat)
                           + CsvExtension;
            }
            else
            {
                fileName =
                    Scope.Context.Remove(OuPrefix)
                        .Remove(DcPrefix)
                        .Replace(Comma, Hyphen)
                    + DoubleHyphen
                    + QueryType
                    + Hyphen
                    + DateTime.Now.ToString(DateTimeFormat)
                    + CsvExtension;
            }
            var fullFileName = Path.Combine(OutputDirectoryPath, fileName);
            var dataTable = Data.ToTable();
            var stringBuilder = new StringBuilder();

            var columnNames = dataTable.Columns.Cast<DataColumn>()
                .Select(column => column.ColumnName);
            stringBuilder.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dataTable.Rows)
            {
                var fields = row.ItemArray.Select(
                    field => string.Concat(
                        "\"",
                        field.ToString().Replace("\"", "\"\""),
                        "\""));
                stringBuilder.AppendLine(string.Join(",", fields));
            }

            Directory.CreateDirectory(OutputDirectoryPath);
            File.WriteAllText(fullFileName, stringBuilder.ToString());
            Process.Start(fullFileName);
            return fullFileName;
        }
    }
}