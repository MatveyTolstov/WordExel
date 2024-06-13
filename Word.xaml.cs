using Microsoft.Win32;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WordExceel
{
    /// <summary>
    /// Логика взаимодействия для Word.xaml
    /// </summary>
    public partial class Word : Window
    {
        public Word()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog dlg = new SaveFileDialog();


            dlg.Filter = "Документы Word (*.docx)|*.docx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {

                string filename = dlg.FileName;
                SaveFile(filename);
            }

            MessageBox.Show("Файл успешно сохранён");

        }

        private void SaveFile(string filename)
        {

            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            FileStream fileStream = new FileStream(filename, FileMode.Create);
            range.Save(fileStream, DataFormats.Rtf);
            fileStream.Close();
            Document doc = new Document();
            doc.LoadFromFile(filename);
            doc.SaveToFile(filename, FileFormat.Docx);
        }

        private void LoadFile(string filename)
        {
            if (File.Exists(filename))
            {
                Document doc = new Document();
                doc.LoadFromFile(filename);
                doc.SaveToFile(filename, FileFormat.Rtf);
                TextRange textRange = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                FileStream fileStream = new FileStream(filename, FileMode.OpenOrCreate);
                textRange.Load(fileStream, DataFormats.Rtf);
                fileStream.Close();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.Filter = "Документы Word (*.docx)|*.docx";

            if (dlg.ShowDialog() == true)
            {
                string filename = dlg.FileName;
                LoadFile(filename);
            }
        }
    }
}
