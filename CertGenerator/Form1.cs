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
                    
                }
                if (inputFile.Trim() != "")
                {
                    //readExcel(inputFile);
                    makeDataFrame(inputFile);
                    
                }
            }
        }

        private void makeDataFrame(string InFile)
        {
            var dataFrame = DataFrame.LoadCsv(InFile, ',');
            var Quals = dataFrame.Rows;

            foreach (DataFrameRow r in Quals){
                MessageBox.Show(r.ToString());
            }
        }       

    }
}
