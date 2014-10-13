using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox;
using org.apache.pdfbox.util;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Data;

namespace WindowsFormsApplication1
{
    public class Component: IComparable<Component>
    {
        public string name { get; set; }
        public double concentration { get; set; }
        public double time { get; set; }
        public int CompareTo(Component comp)
        {
            if (this.concentration < comp.concentration) return -1;
            if (this.concentration > comp.concentration) return 1; else return 0;
        }
        public Component(string name, double concentration, double time)
        {
            this.name = name;
            this.concentration = concentration;
            this.time = time;
        }
        public Component()
        {
            name = string.Empty;
            concentration = time = 0;
        }

        public void Copy(Component comp)
        {
            this.name = comp.name;
            this.concentration = comp.concentration;
            this.time = comp.time;
        }

        public bool Equals(Component comp)
        {
            if ((name.ToUpper() == comp.name.ToUpper()) && (time == comp.time) && (concentration == comp.concentration))
                return true;
            return false;
        }
    }

    public class Chrom
    {
        public string name { get; set; }
        public List<Component> components;
        public event Action onComponentAdded;        
        public Chrom(string name, List<Component> components)
        {
            this.name = name;
            this.components = new List<Component>();
            this.components.AddRange(components);
        }
        public Chrom(string name)
        {
            this.name = name;
            this.components = new List<Component>();
        }
        public Chrom()
        {
            components = new List<Component>();
        }
        public bool addComponent(string name, double concentration, double time)
        {
            Component compToAdd = new Component(name, concentration, time);
            if (!components.Contains(compToAdd))
            {
                components.Add(compToAdd);
                if (onComponentAdded != null)onComponentAdded();
                return true;
            }
            return false;
        }
        public bool addComponent(Component comp)
        {
            return addComponent(comp.name, comp.concentration, comp.time);
        }
        public void Parse(string text)
        {
            List<string> lines = new List<string>();
            text = text.Remove(0, text.IndexOf('№'));
            lines.AddRange(Regex.Split(text, @"\r\n"));

            foreach (string str in lines)
            {
                string name_str;
                string concentration_str;
                string time_str;

                time_str = Regex.Match(str, @"\d+\.\d+(?=\s+\w+)").Value;
                name_str = Regex.Match(str, String.Format(@"(?<={0}\s+)\S+",time_str)).Value;
                concentration_str = Regex.Match(str, String.Format(@"(?<={0}\s+)\b\d+\.\d+\b",name_str)).Value;

                // Последняя проверка в условии - чтобы имя компонента не состояло из одних цифр 100.000 из конца хроматограммы
                if (time_str != "" && name_str != "" && concentration_str != "" && Regex.Matches(name_str,@"[\d,\.]").Count != name_str.Length)
                {
                    double concentration = double.Parse(concentration_str, CultureInfo.InvariantCulture);
                    double time = double.Parse(time_str, CultureInfo.InvariantCulture);

                    addComponent(new Component(name_str, concentration, time));
                }

               // NormalizeMe();
            }
        }

        private void NormalizeMe()
        {
            double summ=0;
            for (int i = 0; i < components.Count; i++)
            {
                summ += components[i].concentration;
            }
            for (int i = 0; i < components.Count; i++)
            {
                components[i].concentration = Math.Round(components[i].concentration / summ * 100, 3);
            }
        }
    }
    
    public class Experiment
    {
        public string Name { get; set; }
        public List<Chrom> Chroms;
        public event Action onChromAdded;
        
        public Chrom feedChrom;
        public void AddChromFromFile(string path, Action onComponentAdded)
        {
            PDDocument document = PDDocument.load(path);
            PDFTextStripper stripper = new PDFTextStripper();
            string text = stripper.getText(document);
            document.close();
            
            Chroms.Add(new Chrom(Path.GetFileNameWithoutExtension(path)));
            Chroms[Chroms.Count - 1].onComponentAdded += onComponentAdded;
            Chroms[Chroms.Count-1].Parse(text);
            if (onChromAdded!= null) onChromAdded();
        }
        public void AddFeedChromFromFile(string path)
        {
            PDDocument document = PDDocument.load(path);
            PDFTextStripper stripper = new PDFTextStripper();
            string text = stripper.getText(document);
            document.close();
            feedChrom.Parse(text);
        }
        public double CalculateConversion(Chrom chrom)
        {
            double conversion = 0;

            if (feedChrom != null)
            {
                foreach (Component component in chrom.components)
                {
                    foreach (Component feedcomponent in feedChrom.components)
                    {
                        if (feedcomponent.name.ToUpper() == component.name.ToUpper())
                        {
                            double delta = feedcomponent.concentration - component.concentration;
                            if (delta > 0) conversion += delta;
                        }
                    }
                }
            }

            return conversion;
        }
        public Experiment(string name)
        {
            this.Name = name;
            Chroms = new List<Chrom>();
            feedChrom = new Chrom();
        }
    }

    public class ExcelAdapter : IDisposable
    {
        private Excel.Application exApp;
        private Excel.Workbook exWb;
        private Excel.Worksheet exWs;
        private DataTable table;
        public ExcelAdapter()
        {
            exApp = new Excel.Application();
            exApp.SheetsInNewWorkbook = 1;
            exApp.Visible = false;
            exApp.DisplayAlerts = false;
            exApp.Workbooks.Add();          
            exWb = exApp.Workbooks.get_Item(1);
            exWs = exWb.Sheets.get_Item(1);
            table = new DataTable();
        }

        private void checkTableForNullReferences()
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if (table.Rows[i][j] == null) table.Rows[i][j] = 0;
                }
            }
        }
        public void fillDataTable(IEnumerable<Chrom> chroms, IEnumerable<string> componentsName)
        {
            table = new DataTable();
            table.Columns.Add("Хроматограмма");

           /* for (int i = 0; i < componentsName.Count(); i++)
                componentsName[i] = componentsName[i].ToLower();*/

            foreach (string name in componentsName)
                    table.Columns.Add(name);
            
            foreach (Chrom chrom in chroms)
            {
                DataRow rowToAdd = table.NewRow();
                rowToAdd["Хроматограмма"] = chrom.name;
                foreach (Component component in chrom.components)
                {
                    if (componentsName.Contains(component.name.ToLower()))
                        rowToAdd[component.name] = component.concentration;
                }
                table.Rows.Add(rowToAdd);
            }

            checkTableForNullReferences();
        }

        public void printDataInExcel()
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                exWs.Cells[1, i + 1] = table.Columns[i].ColumnName;
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                exWs.Cells[i + 2, 1] = table.Rows[i][0];
            }

            for (int i = 1; i < table.Columns.Count; i++)
            {
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    try
                    {
                        //table.Rows[j][i]
                        double value = double.Parse(table.Rows[j][i].ToString());
                        exWs.Cells[j + 2, i + 1].Value = value
 ;
                    }
                    catch { }
                }
            }
           exApp.Visible = true;
        }
        public void Close()
        {
            if (exWb != null) exWb.Close();
            if (exApp != null) exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
            exWs = null;
            exWb = null;
            exApp = null;
            GC.Collect();
        }
        public void Dispose()
        {
            Close();
        }
    }
}
