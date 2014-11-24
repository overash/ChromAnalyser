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
    public enum chromType { liquid, gas };

    public class fillChromAdapter
    {
        string text;
        public fillChromAdapter(string text)
        {
            this.text = text;
        }

        public Chrom fillChrom(Chrom toWhere)
        {   
            chromType type = chromType.gas;
            if (text.ToLower().Contains("расчет по компонентам") || text.ToLower().Contains("расчет хроматограммы")) type = chromType.gas; else type = chromType.liquid;
            bool memem = text.ToLower().Contains("расчет по компонентам");
            Chrom chr = new Chrom(toWhere.name, type);
            
            if (type == chromType.liquid)
            {
                #region liquid
                List<string> lines = new List<string>();
                if (text.IndexOf('№') < 0) return null;
                text = text.Remove(0, text.IndexOf('№'));
                lines.AddRange(Regex.Split(text, @"\r\n"));

                foreach (string str in lines)
                {
                    string name_str = "";
                    string concentration_str = "";
                    string time_str = "";

                    
                    try
                    {
                        time_str = Regex.Match(str, @"\d+\.\d+(?=\s+\w+)").Value;
                        name_str = Regex.Match(str, String.Format(@"(?<={0}\s+)\S+", time_str)).Value;
                        concentration_str = Regex.Match(str, String.Format(@"(?<={0}\s+)\b\d+\.\d+\b", name_str)).Value;
                    }
                    catch
                    {

                    }
                    // Последняя проверка в условии - чтобы имя компонента не состояло из одних цифр 100.000 из конца хроматограммы
                    if (time_str != "" && name_str != "" && concentration_str != "" && Regex.Matches(name_str, @"[\d,\.]").Count != name_str.Length)
                    {
                        double concentration = double.Parse(concentration_str, CultureInfo.InvariantCulture);
                        double time = double.Parse(time_str, CultureInfo.InvariantCulture);

                        chr.addComponent(new Component(name_str.ToLower(), concentration, time));
                    }
                }

                // NormalizeMe();
                #endregion
            }
            
            else
            {
                List<string> lines = new List<string>();
                int pos = text.ToLower().IndexOf("расчет хроматограммы");
                if (pos > 0)
                {
                    text = text.Remove(0, pos + "Расчет хроматограммы".Length);
                    pos = text.IndexOf("Всего по ДТП-2", StringComparison.OrdinalIgnoreCase);
                    text = text.Substring(0, pos - 1);
                    lines.AddRange(Regex.Split(text, @"\r\n").ToArray());
                    foreach (string str in lines)
                    {
                        //" 1,279 водород 94,42522"
                        string name_str = "";
                        string concentration_str = "";
                        string time_str = "";

                        name_str = Regex.Match(str, @"(?<=\d+\,\d+\s+)\S+").Value;
                        time_str = Regex.Match(str, @"\d+\,\d+(?=\s+\S)").Value;
                        concentration_str = Regex.Match(str, @"(?<=\w+\s+)\d+\,\d+").Value;

                        time_str = time_str.Replace(',', '.');
                        concentration_str = concentration_str.Replace(',', '.');

                        if (time_str != "" && name_str != "" && concentration_str != "" && Regex.Matches(name_str, @"[\d,\.]").Count != name_str.Length)
                        {
                            double concentration = double.Parse(concentration_str, CultureInfo.InvariantCulture);
                            double time = double.Parse(time_str, CultureInfo.InvariantCulture);

                            chr.addComponent(new Component(name_str.ToLower(), concentration, time));
                        }
                    }

                }
                else
                {
                    string search = " Время, мин Компонент Группа Площадь Высота Площадь, % Концентрация Ед. концентрации Детектор";
                    pos = text.ToLower().IndexOf(search.ToLower());
                    text = text.Remove(0, pos + search.Length);
                    lines.AddRange(Regex.Split(text, @"\r\n").ToArray());
                    // " 9.448 пропан 4014.656 185.462 14.516 0.549 ДТП-1"
                    foreach (string str in lines)
                    {
                        string name_str = "";
                        string concentration_str = "";
                        string time_str = "";

                        name_str = Regex.Match(str, @"(?<=\d+\.\d+\s+)\S+").Value;
                        time_str = Regex.Match(str, @"\d+\.\d+").Value;
                        concentration_str = Regex.Match(str, @"\d+\.\d+(?=\s+ДТП)").Value;

                        if (time_str != "" && name_str != "" && concentration_str != "" && Regex.Matches(name_str, @"[\d,\.]").Count != name_str.Length)
                        {
                            double concentration = double.Parse(concentration_str, CultureInfo.InvariantCulture);
                            double time = double.Parse(time_str, CultureInfo.InvariantCulture);

                            chr.addComponent(new Component(name_str.ToLower(), concentration, time));
                        }
                    }
                }

            }



            return chr;
        }
    }
    public class Component: IComparable<Component>
    {
        public override string ToString()
        {
            return name;
        }

        public string name { get; set; }
        public double concentration { get; set; }
        public double time { get; set; }

        public chromType type { get; set; }
        
        public int CompareTo(Component comp)
        {
            if (this.name == comp.name) return 0;
            if (this.name.Length > comp.name.Length) return 1; else return -1;
        }
        
        public Component(string name, double concentration, double time)
        {
            this.name = name;
            this.concentration = concentration;
            this.time = time;
            type = chromType.liquid;
        }
        public Component(string name, double concentration, double time, chromType type)
        {
            this.name = name;
            this.concentration = concentration;
            this.time = time;
            this.type = chromType.liquid;
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
            this.type = comp.type;
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
        public override string ToString()
        {
            return name;
        }
        public string name { get; set; }
        public List<Component> components;
        public event Action onComponentAdded;
        public chromType _chromType;
        public Chrom(string name, chromType type, List<Component> components)
        {
            this.name = name;
            this.components = new List<Component>();
            this.components.AddRange(components);
            this._chromType = type;
        }
        public Chrom(string name, chromType type)
        {
            this.name = name;
            this.components = new List<Component>();
            this._chromType = type;
        }

        public Chrom(string name)
        {
            this.name = name;
            this.components = new List<Component>();
            this._chromType = chromType.liquid;
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
      /*  public void Parse(string text)
        {
            fillChromAdapter filler = new fillChromAdapter(text);
            filler.fillChrom(this);
        }*/

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

            fillChromAdapter filler = new fillChromAdapter(text);
            
            Chroms.Add(new Chrom(Path.GetFileNameWithoutExtension(path)));
            Chroms[Chroms.Count - 1].onComponentAdded += onComponentAdded;
            Chroms[Chroms.Count - 1] = filler.fillChrom(Chroms[Chroms.Count - 1]);
            if (onChromAdded!= null) onChromAdded();
        }
        public void AddFeedChromFromFile(string path)
        {
            PDDocument document = PDDocument.load(path);
            PDFTextStripper stripper = new PDFTextStripper();
            string text = stripper.getText(document);
            document.close();
            fillChromAdapter filler = new fillChromAdapter(text);
            feedChrom = filler.fillChrom(feedChrom);
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

        public List<Chrom> getLiquidChroms()
        {
            var liquidChroms = from n in Chroms where n._chromType == chromType.liquid select n;
            return new List<Chrom>(liquidChroms.ToArray());
        }
        public List<Chrom> getGasChroms()
        {
            var gasChroms = from n in Chroms where n._chromType == chromType.gas select n;
            return new List<Chrom>(gasChroms.ToArray());
        }
    }

    public class ExcelAdapter : IDisposable
    {
        private Excel.Application exApp;
        private Excel.Workbook exWb;
        private Excel.Worksheet exWs;
        private DataTable tableLiq, tableGas;
        public ExcelAdapter()
        {
            exApp = new Excel.Application();
            exApp.SheetsInNewWorkbook = 1;
            exApp.Visible = false;
            exApp.DisplayAlerts = false;
            exApp.Workbooks.Add();          
            exWb = exApp.Workbooks.get_Item(1);
            
            exWs = exWb.Sheets.get_Item(1);
            tableLiq = new DataTable();
            tableGas = new DataTable();
        }

        private void checkTableForNullReferences()
        {
            for (int i = 0; i < tableLiq.Rows.Count; i++)
            {
                for (int j = 0; j < tableLiq.Columns.Count; j++)
                {
                    if (tableLiq.Rows[i][j] == null) tableLiq.Rows[i][j] = 0;
                }
            }
        }

        public void fillDataTableLiquid(IEnumerable<Chrom> chroms, IEnumerable<string> componentsName)
        {
            tableLiq = new DataTable();
            tableLiq.Columns.Add("Хроматограмма");

            /* for (int i = 0; i < componentsName.Count(); i++)
                 componentsName[i] = componentsName[i].ToLower();*/

            foreach (string name in componentsName)
                tableLiq.Columns.Add(name);

            foreach (Chrom chrom in chroms)
            {
                DataRow rowToAdd = tableLiq.NewRow();
                rowToAdd["Хроматограмма"] = chrom.name;
                /*foreach (Component component in chrom.components)
                {
                    if (componentsName.Contains(component.name.ToLower()))
                        rowToAdd[component.name] = component.concentration;
                }*/

                foreach (Component component in chrom.components)
                {
                    foreach (string name in componentsName)
                    {
                        if (component.name.ToLower() == name.ToLower())
                            rowToAdd[component.name] = component.concentration;
                    }
                }

                tableLiq.Rows.Add(rowToAdd);
            }

            checkTableForNullReferences();
        }
        public void fillDataTableGas(IEnumerable<Chrom> chroms)
        {
            List<string> componentsName = new List<string>();
            componentsName.AddRange(new string[] {"водород","метан", "этан", "этилен", "пропилен", "пропан", "изобутан", "н-бутан", "изопентан", "н-пентан"});
            foreach (Chrom chr in chroms)
            {
                var name = from n in chr.components select n.name.ToLower(); 
                componentsName.AddRange(name);
            }

            componentsName = new List<string>(componentsName.Distinct().ToArray());

            tableGas = new DataTable();
            tableGas.Columns.Add("Хроматограмма");


            foreach (string name in componentsName)
                tableGas.Columns.Add(name);

            foreach (Chrom chrom in chroms)
            {
                DataRow rowToAdd = tableGas.NewRow();
                rowToAdd["Хроматограмма"] = chrom.name;
              

                foreach (Component component in chrom.components)
                {
                    foreach (string name in componentsName)
                    {
                        if (component.name.ToLower() == name.ToLower())
                            rowToAdd[component.name] = component.concentration;
                    }
                }

                tableGas.Rows.Add(rowToAdd);
            }

            checkTableForNullReferences();
        }

        public void printDataInExcel()
        {
            int i = 0, j = 0;
            #region liquid
            for ( i = 0; i < tableLiq.Columns.Count; i++)
                exWs.Cells[1, i + 1] = tableLiq.Columns[i].ColumnName;
            for ( i = 0; i < tableLiq.Rows.Count; i++)
                exWs.Cells[i + 2, 1] = tableLiq.Rows[i][0];
            for ( i = 1; i < tableLiq.Columns.Count; i++)
            {
                for (j = 0; j < tableLiq.Rows.Count; j++)
                {
                    try
                    {
                        //table.Rows[j][i]
                        double value = double.Parse(tableLiq.Rows[j][i].ToString());
                        exWs.Cells[j + 2, i + 1].Value = value;
                    }
                    catch
                    {
                        exWs.Cells[j + 2, i + 1].Value = 0;
                    }
                }
            }
            #endregion

            #region Gas
            int startRow, startCol;
            startRow = tableLiq.Rows.Count + 2 ;
            startCol = tableLiq.Columns.Count + 1;

            for ( i = 0; i < tableGas.Columns.Count; i++)
                exWs.Cells[startRow, i + 1] = tableGas.Columns[i].ColumnName;
            for ( i = 0; i < tableGas.Rows.Count; i++)
                exWs.Cells[i + startRow + 1, 1] = tableGas.Rows[i][0];

            for (i = 1; i < tableGas.Columns.Count; i++)
            {
                for (j = 0; j < tableGas.Rows.Count; j++)
                {
                    try
                    {
                        //table.Rows[j][i]
                        double value = double.Parse(tableGas.Rows[j][i].ToString());
                        exWs.Cells[j + startRow + 1, i + 1].Value = value;
                    }
                    catch
                    {
                        exWs.Cells[j + startRow + 1, i + 1].Value = 0;
                    }
                }
            }
            #endregion 

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
