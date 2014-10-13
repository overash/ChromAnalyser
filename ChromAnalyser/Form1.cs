using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox;
using org.apache.pdfbox.util;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

    
        Experiment exp;
        List<string> componentsNames;
        private void button5_Click(object sender, EventArgs e)
        {
            loadedChroms.Items.Clear();
            MessageBox.Show("Выберите папку с хроматограммами");
            FolderBrowserDialog browseDialog = new FolderBrowserDialog();
            browseDialog.ShowDialog();
            string path = browseDialog.SelectedPath;
            if (path == "") return;

            string[] files = Directory.GetFiles(path);
            exp = new Experiment("Text");
            
            foreach (string str in files)
                exp.AddChromFromFile(str, null);

            componentsNames = new List<string>();
            foreach (Chrom chrom in exp.Chroms)
            {
                var names = from n in chrom.components select n.name.ToLower();
                componentsNames.AddRange(names);
            }
            loadedChroms.Items.AddRange(componentsNames.Distinct().ToArray());

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> componentsToReport = new List<string>();
            foreach (object obj in loadedChroms.CheckedItems)
            {
                componentsToReport.Add(obj.ToString());
            }
            ExcelAdapter exAdapter = new ExcelAdapter();
            exAdapter.fillDataTable(exp.Chroms, componentsToReport);
            exAdapter.printDataInExcel();
        }
    }
}
