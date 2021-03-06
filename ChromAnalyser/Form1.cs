﻿using System;
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

            string[] files = Directory.GetFiles(path, "*.pdf");
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
            if (exp.Chroms != null && exp.Chroms.Count > 0)
            {
                foreach (object obj in loadedChroms.CheckedItems)
                {
                    componentsToReport.Add(obj.ToString());
                }
                ExcelAdapter exAdapter = new ExcelAdapter();
                exAdapter.fillDataTable(exp.Chroms, componentsToReport);
                exAdapter.printDataInExcel();
            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (loadedChroms.SelectedItem != null && loadedChroms.SelectedIndex > 0)
            {
                Object temp;
                int index = loadedChroms.SelectedIndex;
                temp = loadedChroms.SelectedItem;
                loadedChroms.Items.RemoveAt(index);
                loadedChroms.Items.Insert(index - 1, temp);
                loadedChroms.SetSelected(index - 1, true);
            }
        }

        private void loadedChroms_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (loadedChroms.SelectedItem != null && loadedChroms.SelectedIndex < loadedChroms.Items.Count-1)
            {
                Object temp;
                int index = loadedChroms.SelectedIndex;
                temp = loadedChroms.SelectedItem;
                loadedChroms.Items.RemoveAt(index);
                loadedChroms.Items.Insert(index + 1, temp);
                loadedChroms.SetSelected(index + 1, true);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            SerializationAdapter save = new SerializationAdapter(saveFileDialog1.FileName, loadedChroms);
            save.Serialize();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (File.Exists(openFileDialog1.FileName))
            {
                SerializationAdapter load = new SerializationAdapter(openFileDialog1.FileName, loadedChroms);
                load.Deserialize();
            }
        }
    }
}
