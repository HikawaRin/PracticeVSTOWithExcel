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

namespace code.View.WPF
{
    /// <summary>
    /// SidePanelView.xaml 的交互逻辑
    /// </summary>
    public partial class SidePanelView : UserControl
    {
        public WorkSheetViewModel WorkSheetViewModel { get; set; }

        public SidePanelView()
        {
            InitializeComponent();
            WorkSheetViewModel = null;
        }

        private void InputDataButton_Click(object sender, RoutedEventArgs e)
        {
            WorkSheetViewModel = new WorkSheetViewModel();

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
    }
}
