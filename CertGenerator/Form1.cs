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

namespace CertGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string inputFile;
        string MasterCert = "C://Users//bradj//OneDrive//Desktop//MASTERCert.docx";

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
            sourceDoc.ActiveWindow.Selection.CopyFormat();

            var newDocument = new Microsoft.Office.Interop.Word.Document();
            newDocument.ActiveWindow.Selection.PasteFormat();
            newDocument.SaveAs(@"D:\test1.docx");

            sourceDoc.Close();
            newDocument.Close();

            app.Quit();
        }

    }
}
