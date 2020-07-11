using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Point = System.Drawing.Point;

namespace CsvToExcel
{
    class Dialog
    {
        private OpenFileDialog dialog = null;
        private Panel p = new Panel();
        private ComboBox encodingComboBox = new ComboBox();
        private ComboBox delimiterComboBox = new ComboBox();
        private ComboBox textQualifierCombBox = new ComboBox();
        private System.Text.EncodingInfo[] encodings = System.Text.Encoding.GetEncodings();
        private List<Encoding> encs = new List<Encoding>();
        private List<Delimiter> delimiters = new List<Delimiter>() {
            new Delimiter() { Name = "タブ", Value = "\t" },
            new Delimiter() { Name = "カンマ", Value = "," },
            new Delimiter() { Name = "セミコロン", Value = ";" },
            new Delimiter() { Name = "スペース", Value = " " }
        };
        private List<TextQualifier> textQualifiers = new List<TextQualifier>() {
            new TextQualifier() { Name = "なし", Value = "", ExcelValue = XlTextQualifier.xlTextQualifierNone },
            new TextQualifier() { Name = "ダブルクォーテーション", Value = "\"", ExcelValue = XlTextQualifier.xlTextQualifierDoubleQuote },
            new TextQualifier() { Name = "シングルクォーテーション", Value = "'", ExcelValue = XlTextQualifier.xlTextQualifierSingleQuote }
        };

        public Dialog(string path)
        {
            // Create panel for the selection part of the dialog
            //Panel p = new Panel();
            //p.Size = new Size(0, 0);
            p.BorderStyle = BorderStyle.None;
            p.AutoSize = true;
            p.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var encodingLabel = new System.Windows.Forms.Label();
            encodingLabel.Text = "文字コード:";
            encodingLabel.Font = SystemFonts.MessageBoxFont;

            foreach (var item in System.Text.Encoding.GetEncodings()) {
                encodingComboBox.Items.Add(new Encoding() { Name = item.DisplayName, Value = item });
            }
            encodingComboBox.SelectedItem = encodingComboBox.Items.Cast<Encoding>().Single(x => x.Value.CodePage == System.Text.Encoding.Default.CodePage);
            encodingComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            encodingComboBox.Font = SystemFonts.MessageBoxFont;
            encodingComboBox.Width += 160;
            encodingComboBox.Location = new Point(encodingLabel.Location.X, encodingLabel.Location.Y + 20);

            var delimiterLabel = new System.Windows.Forms.Label();
            delimiterLabel.Text = "区切り文字:";
            delimiterLabel.Location = new Point(encodingLabel.Location.X, encodingComboBox.Location.Y + 26);
            delimiterLabel.Font = SystemFonts.MessageBoxFont;

            foreach (var item in delimiters) {
                delimiterComboBox.Items.Add(item);
            }
            delimiterComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            delimiterComboBox.Font = SystemFonts.MessageBoxFont;
            delimiterComboBox.Location = new Point(encodingLabel.Location.X, delimiterLabel.Location.Y + 20);
            delimiterComboBox.Width += 40;
            delimiterComboBox.SelectedIndex = 0;

            var textQualifierLabel = new System.Windows.Forms.Label();
            textQualifierLabel.Text = "囲み文字:";
            textQualifierLabel.Font = SystemFonts.MessageBoxFont;
            textQualifierLabel.Location = new Point(encodingLabel.Location.X, delimiterComboBox.Location.Y + 26);

            foreach (var item in textQualifiers) {
                textQualifierCombBox.Items.Add(item);
            }
            textQualifierCombBox.DropDownStyle = ComboBoxStyle.DropDownList;
            textQualifierCombBox.Font = SystemFonts.MessageBoxFont;
            textQualifierCombBox.Width += 40;
            textQualifierCombBox.Location = new Point(encodingLabel.Location.X, textQualifierLabel.Location.Y + 20);
            textQualifierCombBox.SelectedIndex = 0;

            p.Controls.Add(encodingComboBox);
            p.Controls.Add(delimiterComboBox);
            p.Controls.Add(textQualifierCombBox);
            p.Controls.Add(encodingLabel);
            p.Controls.Add(delimiterLabel);
            p.Controls.Add(textQualifierLabel);

            // Create and show the OpenFile Dialog
            dialog = new OpenFileDialog("", path, "すべてのファイル (*.*)\0*.*\0\0", p, "CsvToExcel");
        }

        public bool ShowDialog()
        {
            return dialog.Show();
        }

        public string FileName
        {
            get { return dialog.SelectedPath; }
        }

        public Encoding Encoding
        {
            get { return (Encoding)encodingComboBox.SelectedItem; }
        }

        public Delimiter Delimiter
        {
            get { return (Delimiter)delimiterComboBox.SelectedItem; }
        }

        public TextQualifier TextQualifier
        {
            get { return (TextQualifier)textQualifierCombBox.SelectedItem; }
        }
    }
}
