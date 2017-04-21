using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using outlook = Microsoft.Office.Interop.Outlook;

namespace test
{
    public class Email
    {
        private outlook.MailItem Mail;

        public Email(outlook.MailItem Mail)
        {
            this.Mail = Mail;
        }

        public outlook.MailItem getMail()
        {
            return Mail;
        }

        public override string ToString()
        {
            return Mail.Subject;
        }
    }
}
