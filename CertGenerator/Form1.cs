using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Data.Analysis;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace CertGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string inputFile;
        static string fileName = "MASTERCert.docx";
        string MasterCert = Path.Combine(Environment.CurrentDirectory, fileName);
        string savePathPartOne;

        private void Btn_Generate_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel file to read from";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls;*.csv";

            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                inputFile = openFileDialog1.FileName;
                
                if (inputFile.Contains(".xl"))
                {
                    // Some way of converting .xl files to .csv
                }
                if (inputFile.Trim() != "")
                {
                    //readExcel(inputFile);
                    var GradFrame = makeDataFrame(inputFile);
                    ChooseFolder();
                    fillCert(GradFrame);
                }
            }
           
        }

        private DataFrame makeDataFrame(string InFile)
        {
            var dataFrame = DataFrame.LoadCsv(InFile, ',');
            var Quals = dataFrame.Rows;

            foreach (DataFrameRow r in Quals){  // To Test Dataframe outputs
                MessageBox.Show(r.ToString());
            }
            return dataFrame;
        }   
        
        private void fillCert(DataFrame Graduates)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            var sourceDoc = app.Documents.Open(MasterCert);

            sourceDoc.ActiveWindow.Selection.WholeStory();
            sourceDoc.ActiveWindow.Selection.Copy();



            for (int i = 0; i < Graduates.Rows.Count; i++)
            {
                var newDocument = new Microsoft.Office.Interop.Word.Document();
                newDocument.ActiveWindow.Selection.PasteAndFormat(default);
                ReplaceText(newDocument.Content, "$[NAME]", Graduates[i, 1].ToString());
                ReplaceText(newDocument.Content, "$[COURSE]", Graduates[i,2].ToString());
                ReplaceText(newDocument.Content, "$[COMPLETION]", Graduates[i,3].ToString());
                ReplaceText(newDocument.Content, "$[SIGNATORY]", Graduates[i,4].ToString());

                foreach(Section section in newDocument.Sections)
                {
                    foreach(HeaderFooter footer in section.Footers)
                    {
                       ReplaceText(footer.Range, "$[CERT#]", Graduates[i, 0].ToString());
                       ReplaceText(footer.Range, "$[VALIDEND]", Graduates[i,5].ToString());
                   }
                }
                string savePath = (Graduates[i, 1].ToString() + "-" + Graduates[i, 2].ToString());
                newDocument.SaveAs(Path.Combine(savePathPartOne,"%name%".Replace("%name%", savePath)));
                newDocument.Close();
            }
            

            sourceDoc.Close();
            

            app.Quit();
        }

        static void ReplaceText(Range range, string findText, string replaceText)
        {
            Find findObject = range.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceText;

            findObject.Execute(Replace: WdReplace.wdReplaceAll);
        }

        public void ChooseFolder()
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                savePathPartOne = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
