using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PrimitiveExtensions;
using static System.Environment;
using static System.Environment.SpecialFolder;

namespace ActiveDirectoryToolWpf
{
    public class DataFileWriter
    {
        private const string DateTimeFormat = "yyyyMMddTHHmmss";

        private readonly string _outputPath = Path.Combine(
            GetFolderPath(MyDocuments), "ActiveDirectoryTool");

        public DataGrid Data { get; set; }
        public QueryType QueryType { get; set; }
        public string Scope { get; set; }

        public string WriteToCsv()
        {
            Data.SelectAllCells();
            Data.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, Data);
            Data.UnselectAllCells();
            var result = (string) Clipboard.GetData(
                DataFormats.CommaSeparatedValue);
            Clipboard.Clear();
            var fileName = Scope.Remove("OU=").Remove("DC=").Replace(',','-')
                + "--" + QueryType + '-' +
                DateTime.Now.ToString(DateTimeFormat) + ".csv";
            using (var writer = new StreamWriter(
                Path.Combine(_outputPath, fileName)))
            {
                writer.WriteLine(result);
            }
            return fileName;
        }
    }
}