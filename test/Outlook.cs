using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;

namespace test
{
    public partial class Outlook
    {
        private void Outlook_Load(object sender, RibbonUIEventArgs e)
        {
           
        }

        private void test_Click(object sender, RibbonControlEventArgs e)
        {

            List<MailItem> mls = new List<MailItem>();

            Application app = Globals.ThisAddIn.Application;
            Explorer exp = app.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                for(int i = 1; i<=exp.Selection.Count; i++)
                {
                    Object selObject = exp.Selection[i];
                    if (selObject is MailItem)
                    {
                        MailItem mailItem =
                            (selObject as MailItem);
                        mls.Add(mailItem);
                    }

                }
                MainWindow mwin = new MainWindow(mls);
                mwin.ShowDialog();

            }
            

        }
    }
}
