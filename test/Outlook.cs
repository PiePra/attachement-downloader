using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace test
{
    public partial class Outlook
    {
        private void Outlook_Load(object sender, RibbonUIEventArgs e)
        {
           
        }

        private void test_Click(object sender, RibbonControlEventArgs e)
        {

            List <Email> mls = new List<Email>();

            outlook.Application app = Globals.ThisAddIn.Application;
            outlook.Explorer exp = app.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                for(int i = 1; i<=exp.Selection.Count; i++)
                {
                    Object selObject = exp.Selection[i];
                    if (selObject is outlook.MailItem)
                    {
                        outlook.MailItem mailItem =
                            (selObject as outlook.MailItem);
                        mls.Add(new Email(mailItem));
                    }

                }
                Form1 Form = new Form1();
                Form.setmls(mls);
                Form.Show();

            }
            

        }
    }
}
