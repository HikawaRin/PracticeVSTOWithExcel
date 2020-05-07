using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using code.ViewModel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace code.View.WPF
{
    /// <summary>
    /// SidePanelView.xaml 的交互逻辑
    /// </summary>
    public partial class SidePanelView : UserControl
    {
        public WorkSheetViewModel WorkSheetViewModel { get; set; }

        public List<Sheet> SheetName;

        public SidePanelView()
        {
            InitializeComponent();
            WorkSheetViewModel = null;
            SheetName = new List<Sheet>();
        }

        private void InputDataButton_Click(object sender, RoutedEventArgs e)
        {
            // WorkSheetViewModel.Certificates.Clear();
            WorkSheetViewModel._loaddata();
            DataList.ItemsSource = WorkSheetViewModel.Certificates;

            List<string> SideList = (from c in WorkSheetViewModel.Certificates
                                     select c.Side).ToList<string>();
            SideList.Insert(0, "空");
            SideComboBox.ItemsSource = SideList.Distinct();
            List<string> MainItemList = (from c in WorkSheetViewModel.Certificates
                                         select c.MainItem).ToList<string>();
            MainItemList.Insert(0, "空");
            MainItemComboBox.ItemsSource = MainItemList.Distinct();
        }

        private void OutputButton_Click(object sender, RoutedEventArgs e)
        {
            WorkSheetViewModel.OutputData();
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {

            string SideString = "空";
            string MainItemString = "空";
            if (SideComboBox.SelectedItem != null)
            {
                SideString = SideComboBox.SelectedItem.ToString();
            }
            if (MainItemComboBox.SelectedItem != null)
            {
                MainItemString = MainItemComboBox.SelectedItem.ToString();
            }
            
            if (SideString == "空" && MainItemString == "空")
            {
                DataList.ItemsSource = WorkSheetViewModel.Certificates;
                return;
            }
            List<Certificate> SearchedList;
            if (SideString == "空" && MainItemString != "空")
            {
                SearchedList = (from item in WorkSheetViewModel.Certificates
                                where item.MainItem == MainItemString
                                select item).ToList<Certificate>();
            }
            else if (SideString != "空" && MainItemString == "空")
            {
                SearchedList = (from item in WorkSheetViewModel.Certificates
                                where item.Side == SideString
                                select item).ToList<Certificate>();
            }
            else
            {
                SearchedList = (from item in WorkSheetViewModel.Certificates
                                where item.Side == SideString &&
                                       item.MainItem == MainItemString
                                select item).ToList<Certificate>();
            }
            
            DataList.ItemsSource = SearchedList;
        }

        private void AllSheet_Click(object sender, RoutedEventArgs e)
        {
            bool state = (bool)AllSheet.IsChecked;
            foreach (Sheet s in SheetName)
            {
                s.isSelect = state;
            }
        }

        private void RefreshSheet_Click(object sender, RoutedEventArgs e)
        {
            if (WorkSheetViewModel is null)
            {
                WorkSheetViewModel = new WorkSheetViewModel();
            }
            
            WorkSheetViewModel.LoadSheets();
            SheetName.Clear();
            foreach (Sheet s in WorkSheetViewModel.Sheets)
            {
                SheetName.Add(s);
            }
            SheetName.Sort((left, right) =>
            {
                if(left.Name.Length > right.Name.Length)
                {
                    return 1;
                }else if (left.Name.Length < right.Name.Length)
                {
                    return -1;
                }
                else
                {
                    int i = 0;
                    while(i < left.Name.Length && left.Name[i] == right.Name[i])
                    {
                        i++;
                    }

                    if (i == left.Name.Length) return 0;

                    if (left.Name[i] > right.Name[i])
                    {
                        return 1;
                    }
                    else
                    {
                        return -1;
                    }
                }
            });
            SheetList.ItemsSource = SheetName;
        }

        private void FiltterBtn_Click(object sender, RoutedEventArgs e)
        {
            Regex regex = new Regex(Filtter.Text);
            SheetName.Clear();
            foreach (Sheet s in WorkSheetViewModel.Sheets)
            {
                if (regex.IsMatch(s.Name))
                {
                    SheetName.Add(s);
                }
            }
            SheetName.Sort((left, right) =>
            {
                if (left.Name[7] > right.Name[7])
                {
                    return 1;
                }
                else if (left.Name[7] == right.Name[7])
                {
                    return 0;
                }
                else
                {
                    return -1;
                }
            });
        }

        private void CheckBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Controls.CheckBox c = (System.Windows.Controls.CheckBox)sender;
            string s = (string)c.Content;
            foreach (Sheet she in WorkSheetViewModel.Sheets)
            {
                if (she.Name == s)
                {
                    she._Sheet.Activate();
                    return;
                }
            }
        }
    }
}
