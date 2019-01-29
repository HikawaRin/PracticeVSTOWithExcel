using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;

namespace code
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void SidePanelButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ChangeWindow(0);
        }
    }
}
