using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace code.ViewModel
{
    public class Sheet
    {
        public bool isSelect { get; set; }

        public string Name { get; set; }

        public Excel.Worksheet _Sheet { get; set; }
    }

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

        public List<Sheet> Sheets { get; set; }

        public List<Certificate> Certificates { get; set; }

        public WorkSheetViewModel()
        {
            _workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Sheets = new List<Sheet>();
            Certificates = new List<Certificate>();
            
            // if(!_loaddata())
            // {
            //     Certificates.Clear();
            // }
        }

        public bool LoadSheets()
        {
            Sheets.Clear();
            foreach (Excel.Worksheet sheet in _workbook.Worksheets)
            {
                Excel.Range range = sheet.UsedRange;

                Excel.Range A1 = range.Range["A1"];
                string Header = (A1.Value2 ?? "").ToString().Replace(" ", "");
                if (Header == "记账凭证")
                {
                    Sheet s = new Sheet
                    {
                        isSelect = false,
                        Name = sheet.Name,
                        _Sheet = sheet
                    };
                    Sheets.Add(s);
                }
            }
            return true;
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

        public bool _loaddata()
        {
            Certificates.Clear();

            foreach (Sheet sh in Sheets)
            {
                if (sh.isSelect == false)
                {
                    continue;
                }
                else
                {
                    Excel.Worksheet sheet = sh._Sheet; 
                    Excel.Range range = sheet.UsedRange;

                    // Excel.Range A1 = range.Range["A1"];
                    Excel.Range Description = range.Range["A3"];
                    string sDescription = (Description.Value2 ?? "").ToString().Replace(" ", "");
                    Excel.Range Mainitem = range.Range["B3"];
                    string sMainitem = (Mainitem.Value2 ?? "").ToString().Replace(" ", "");
                    Excel.Range Subitem = range.Range["C3"];
                    string sSubitem = (Subitem.Value2 ?? "").ToString().Replace(" ", "");
                    Excel.Range JSide = range.Range["D3"];
                    string sJSide = (JSide.Value2 ?? "").ToString().Replace(" ", "");
                    Excel.Range DSide = range.Range["E3"];
                    string sDSide = (DSide.Value2 ?? "").ToString().Replace(" ", "");

                    // if (Description == null &&
                    //     Mainitem == null &&
                    //     Subitem == null &&
                    //     JSide == null &&
                    //     DSide == null)
                    // {
                    //     MessageBox.Show("跳过工作表:" + sheet.Name);
                    //     continue;
                    // }

                    if (sDescription != "摘要" ||
                        sMainitem != "科目" ||
                        sSubitem != "子目" ||
                        sJSide != "借方金额" ||
                        sDSide != "贷方金额")
                    {
                        MessageBox.Show("工作表:" + sheet.Name + "存在格式问题");
                        return false;
                    }

                    string sumstr = "";
                    int cindex = 0;
                    while (sumstr != "合计")
                    {
                        cindex++;
                        string sumindex = "A" + cindex.ToString();
                        Excel.Range Sum = range.Range[sumindex];
                        sumstr = (Sum.Value2 ?? "").ToString().Replace(" ", "");
                    }

                    Excel.Range JSum = range.Range["D" + cindex.ToString()];
                    string JSumValueStr = (JSum.Value2 ?? "").ToString().Replace(" ", "");
                    decimal.TryParse(JSumValueStr, out decimal JSumValue);
                    Excel.Range DSum = range.Range["E" + cindex.ToString()];
                    string DSumValueStr = (DSum.Value2 ?? "").ToString().Replace(" ", "");
                    decimal.TryParse(DSumValueStr, out decimal DSumValue);
                    if (JSumValue != DSumValue)
                    {
                        sh._Sheet.Activate();
                        MessageBox.Show("工作表:" + sheet.Name + "借贷不相等");
                        Certificates.Clear();
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
            }
            return true;
        }
    }
}
