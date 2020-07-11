using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace CsvToExcel
{
    class Excel
    {
        private const int SamplingLines = 10;
        private ExcelApp excel = new ExcelApp();

        public Excel()
        {
            Workbook workbook = excel.Workbooks.Add();
            workbook.RemoveDocumentInformation(XlRemoveDocInfoType.xlRDIAll);
        }

        public void ImportCsv(string path, Encoding encoding, Delimiter delimiter, TextQualifier textQualifier)
        {
            excel.ActiveSheet.Name = Path.GetFileNameWithoutExtension(path);
            int[] columnDataTypes = GetColumnDataTypes(path, encoding, delimiter, textQualifier);
            if (columnDataTypes.Length > 0) {
                QueryTable queryTable = excel.ActiveSheet.QueryTables.Add("TEXT;" + path, (Range)(excel.ActiveSheet.Range("A1")), Type.Missing);
                queryTable.Name = Path.GetFileNameWithoutExtension(path);
                queryTable.FieldNames = true;
                queryTable.RowNumbers = false;
                queryTable.FillAdjacentFormulas = false;
                queryTable.PreserveFormatting = true;
                queryTable.RefreshOnFileOpen = false;
                queryTable.RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
                queryTable.SavePassword = false;
                queryTable.SaveData = true;
                queryTable.AdjustColumnWidth = false;
                queryTable.RefreshPeriod = 0;
                queryTable.TextFilePromptOnRefresh = false;
                try {
                    queryTable.TextFilePlatform = encoding.Value.CodePage;
                } catch {
                    throw new Exception("Microsoft Excel does not support the specified encoding.");
                }
                queryTable.TextFileStartRow = 1;
                queryTable.TextFileParseType = XlTextParsingType.xlDelimited;
                queryTable.TextFileTextQualifier = textQualifier.ExcelValue;
                queryTable.TextFileConsecutiveDelimiter = false;
                queryTable.TextFileTabDelimiter = (delimiter.Value == "\t");
                queryTable.TextFileCommaDelimiter = (delimiter.Value == ",");
                queryTable.TextFileSemicolonDelimiter = (delimiter.Value == ";");
                queryTable.TextFileSpaceDelimiter = (delimiter.Value == " ");
                queryTable.TextFileColumnDataTypes = columnDataTypes;
                queryTable.Refresh(false);
                queryTable.Delete();
                excel.Visible = true;
            }
        }

        public void Quit()
        {
            excel.ActiveWorkbook.Close(0);
            excel.Quit();
        }

        static int[] GetColumnDataTypes(string fileName, Encoding encoding, Delimiter delimiter, TextQualifier textQualifier)
        {
            int columnsCount = 0;
            int linesCount = 0;
            using (var parser = new TextFieldParser(fileName, encoding.Value.GetEncoding())) {
                parser.Delimiters = new string[] { delimiter.Value };
                parser.HasFieldsEnclosedInQuotes = (!string.IsNullOrEmpty(textQualifier.Value));
                while (!parser.EndOfData) {
                    string[] tokens = parser.ReadFields();
                    linesCount++;
                    if (tokens.Length > columnsCount) {
                        columnsCount = tokens.Length;
                    }
                    if (linesCount >= SamplingLines) {
                        break;
                    }
                }
            }
            return Enumerable.Repeat(2, columnsCount).ToArray();
        }
    }
}
