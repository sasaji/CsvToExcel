using System.Text;

namespace CsvToExcel
{
    class Encoding
    {
        public string Name;
        public EncodingInfo Value;

        public override string ToString()
        {
            return Name;
        }
    }
}
