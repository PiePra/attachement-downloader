using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows;
using System.Collections.ObjectModel;

namespace test
{
    public partial class Outlook
    {
        private void Outlook_Load(object sender, RibbonUIEventArgs e)
        {
           
        }

        private void test_Click(object sender, RibbonControlEventArgs e)
        {
            MainWindow window = new MainWindow();
            window.Show();
        }
    }
}
