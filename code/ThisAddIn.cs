using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using code.View.WinForm;

namespace code
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane MyPanel { get; set; }

        private int _index = -1;

        public void ChangeWindow(int index)
        {
            if (index != _index && !MyPanel.Visible)
            {
                MyPanel = this.CustomTaskPanes[index];
                MyPanel.Visible = true;

                _index = index;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            MyPanel = this.CustomTaskPanes.Add(new Panel(), "Side Panel");
            MyPanel.Width = 800;
            MyPanel.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
