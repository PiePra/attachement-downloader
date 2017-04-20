using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace test
{
    public partial class Outlook
    {
        private void Outlook_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void test_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Test");
        }
    }
}
