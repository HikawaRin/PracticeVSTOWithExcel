using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace code.ViewModel
{
    public class Certificate
    {
        public string Source { get; set; }

        public string Side { get; set; }

        public string MainItem { get; set; }

        public string SubItem { get; set; }

        public decimal Money { get; set; }

        public string Description { get; set; }
    }

    public class WorkSheetViewModel
    {
        private Excel.Workbook _workbook { get; set; }

        public List<Certificate> Certificates { get; set; }

        public WorkSheetViewModel()
        {
            _workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Certificates = new List<Certificate>();

            if(!_loaddata())
            {
                Certificates.Clear();
            }
        }

        public void OutputData()
        {
            Excel.Worksheet datasheet =  (Excel.Worksheet)_workbook.Worksheets.Add();
            datasheet.Name = "汇总";

            _insertHead(datasheet);
            Dictionary<string, decimal> JDatas = new Dictionary<string, decimal>();
            Dictionary<string, decimal> DDatas = new Dictionary<string, decimal>();
            int count = 2;
            List<string> headers = new List<string>();
            foreach (Certificate c in Certificates)
            {
                if (c.Side == "借方")
                {
                    if (JDatas.ContainsKey(c.MainItem))
                    {
                        JDatas[c.MainItem] += c.Money;
                    }
                    else
                    {
                        JDatas.Add(c.MainItem, c.Money);
                        bool status = false;
                        foreach (string header in headers)
                        {
                            if (header == c.MainItem)
                            {
                                status = true;
                                break;
                            }
                        }
                        if (!status)
                        {
                            headers.Add(c.MainItem);
                            datasheet.Range["A" + count.ToString()].Value = c.MainItem;
                            count++;
                        }
                    }
                }
                else
                {
                    if (DDatas.ContainsKey(c.MainItem))
                    {
                        DDatas[c.MainItem] += c.Money;
                    }
                    else
                    {
                        DDatas.Add(c.MainItem, c.Money);
                        bool status = false;
                        foreach (string header in headers)
                        {
                            if (header == c.MainItem)
                            {
                                status = true;
                                break;
                            }
                        }
                        if (!status)
                        {
                            headers.Add(c.MainItem);
                            datasheet.Range["A" + count.ToString()].Value = c.MainItem;
                            count++;
                        }
                    }
                }
            }
            for (int i = 2; i < count; i++)
            {
                string head = datasheet.Range["A" + i.ToString()].Value.ToString();
                if (JDatas.ContainsKey(head))
                {
                    datasheet.Range["B" + i.ToString()].Value = JDatas[head];
                }

                if (DDatas.ContainsKey(head))
                {
                    datasheet.Range["C" + i.ToString()].Value = DDatas[head];
                }
            }
        }

        private void _insertHead(Excel.Worksheet sheet)
        {
            sheet.Range["B1"].Value = "借方";
            sheet.Range["C1"].Value = "贷方";
        }

        private bool _loaddata()
        {
            foreach (Excel.Worksheet sheet in _workbook.Worksheets)
            {
                Excel.Range range = sheet.UsedRange;
                
                Excel.Range Description = range.Find("摘要");
                Excel.Range Mainitem = range.Find("科目");
                Excel.Range Subitem = range.Find("子目");
                Excel.Range JSide = range.Find("借方金额");
                Excel.Range DSide = range.Find("贷方金额");
                
                if (Description == null && 
                    Mainitem == null && 
                    Subitem == null && 
                    JSide == null && 
                    DSide == null)
                {
                    MessageBox.Show("跳过工作表:" + sheet.Name);
                    continue;
                }

                if (Description == null ||
                    Mainitem == null ||
                    Subitem == null ||
                    JSide == null ||
                    DSide == null)
                {
                    MessageBox.Show("工作表:" + sheet.Name + "存在格式问题");
                    return false;
                }

                    for (int i = Mainitem.Row + 1; i <= range.Rows.Count; i++)
                {
                    Certificate c = new Certificate
                    {
                        Source = sheet.Name
                    };
                    for (int j = range.Column; j <= range.Columns.Count; j++)
                    {
                        int ij = (int)'A' + j - 1;
                        string s = ((char)ij).ToString() + i.ToString();
                        Excel.Range cell = range.Range[s];
                        string cell_value = (cell.Value2 ?? "").ToString().Replace(" ", "");
                        if (string.IsNullOrWhiteSpace(cell_value))
                        {
                            continue;
                        }

                        if (j == Description.Column)
                        {
                            c.Description = cell_value;
                        }
                        else if (j == Mainitem.Column)
                        {
                            c.MainItem = cell_value;
                        }
                        else if (j == Subitem.Column)
                        {
                            c.SubItem = cell_value;
                        }
                        else if (j == JSide.Column)
                        {
                            c.Side = "借方";
                            if (decimal.TryParse(cell_value, out decimal result))
                            {
                                c.Money = result;
                            }
                        }
                        else if (j == DSide.Column)
                        {
                            if (c.Side == "借方")
                            {
                                Certificate jc = new Certificate
                                {
                                    Description = c.Description,
                                    MainItem = c.MainItem,
                                    Side = c.Side,
                                    Money = c.Money,
                                    Source = c.Source,
                                    SubItem = c.SubItem
                                };
                                if (!string.IsNullOrWhiteSpace(jc.MainItem) && jc.Money != 0)
                                {
                                    Certificates.Add(jc);
                                }
                            }
                            c.Side = "贷方";
                            if (decimal.TryParse(cell_value, out decimal result))
                            {
                                c.Money = result;
                            }
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(c.MainItem) && c.Money != 0)
                    {
                        Certificates.Add(c);
                    }
                }
            }
            return true;
        }
    }
}
