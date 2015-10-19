using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Dbg = System.Diagnostics.Debug;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook wb, bool saveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range dataRange = activeWorksheet.Range["data"];
            Excel.Range dataRows = dataRange.Rows;

            Dictionary<Tuple<string, string>, Tuple<List<string>, int>> data =
                new Dictionary<Tuple<string, string>, Tuple<List<string>, int>>();
            foreach (Excel.Range row in dataRows){
                Object id_obj = row.Cells[1,1].Value2;
                if (id_obj != null)
                {
                    string id = id_obj.ToString();
                    for (int i = 2; i < row.Columns.Count; i++ )
                    {
                        for (int j = i + 1; j <= row.Columns.Count; j++)
                        {
                            string a_investor = (string)row.Cells[1, i].Value2;
                            string b_investor = (string)row.Cells[1, j].Value2;

                            //handle ordering
                            if (!String.IsNullOrEmpty(a_investor) &&
                                !String.IsNullOrEmpty(b_investor))
                            {
                                if (a_investor.CompareTo(b_investor) > 0)
                                {
                                    string tmp = b_investor;
                                    b_investor = a_investor;
                                    a_investor = tmp;
                                }

                                Tuple<string, string> pair = new Tuple<string, string>(a_investor, b_investor);

                                if (data.ContainsKey(pair))
                                {
                                    int new_count = data[pair].Item2 + 1;
                                    List<string> new_list = data[pair].Item1;
                                    new_list.Add(id);

                                    data[pair] = new Tuple<List<string>, int>(new_list, new_count);
                                }
                                else
                                {
                                    List<string> investments = new List<string>() { id };
                                    Tuple<List<string>, int> new_value = new Tuple<List<string>, int>(investments, 1);
                                    data.Add(pair, new_value);
                                }
                            }
                        }
                    }
                }
            }

            
            int next_open_column = activeWorksheet.UsedRange.Columns.Count;
            activeWorksheet.Cells[1, next_open_column].Value2 = "Investor 1";
            activeWorksheet.Cells[1, next_open_column + 1].Value2 = "Investor 2";
            activeWorksheet.Cells[1, next_open_column + 2].Value2 = "Number of Common Investments";
            int output_row = 2;
            foreach (KeyValuePair<Tuple<string, string>, Tuple<List<string>, int>> entry in data)
            {
                int occurences = entry.Value.Item2;
                if (occurences > 1)
                {
                    Dbg.WriteLine(entry.Key.Item1 + ", " + entry.Key.Item2 + " " + entry.Value.Item2);
                }
                activeWorksheet.Cells[output_row, next_open_column].Value2 = entry.Key.Item1;
                activeWorksheet.Cells[output_row, next_open_column + 1].Value2 = entry.Key.Item2;
                activeWorksheet.Cells[output_row, next_open_column + 2].Value2 = entry.Value.Item2;
                output_row++;
            }
            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
