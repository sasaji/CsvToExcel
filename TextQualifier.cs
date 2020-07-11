using Microsoft.Office.Interop.Excel;

namespace CsvToExcel
{
    class TextQualifier
    {
        public string Name;
        public string Value;
        public XlTextQualifier ExcelValue;

        public override string ToString()
        {
            return Name;
        }
    }
}
