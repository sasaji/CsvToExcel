using System;
using System.Windows.Forms;

namespace CsvToExcel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var excel = new Excel();
            var dialog = new Dialog(args.Length == 0 ? string.Empty : args[0]);
            while (true) {
                if (dialog.ShowDialog()) {
                    try {
                        excel.ImportCsv(dialog.FileName, dialog.Encoding, dialog.Delimiter, dialog.TextQualifier);
                        break;
                    } catch (Exception ex) {
                        MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                } else {
                    excel.Quit();
                    break;
                }
            }
        }

        static IntPtr MyHookProc(IntPtr hWnd, UInt16 msg, Int32 wParam, Int32 lParam)
        {
            switch (msg) {
                case WindowMessage.Size:
                    return IntPtr.Zero;
                default:
                    return IntPtr.Zero;
            }
        }
    }
}
