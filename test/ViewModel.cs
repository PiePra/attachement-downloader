using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

namespace test
{
    public class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private ObservableCollection<MailItem> _Mails;
        private string test;
        private string homeDrive;
        private MailItem ml;


        public void fillMails()
        {
            _Mails = new ObservableCollection<MailItem>();
            Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Outlook.Explorer exp = app.ActiveExplorer();

            if (exp.Selection.Count > 0)
            {
                //MessageBox.Show(exp.Selection.Count.ToString());
                for (int i = 1; i <= exp.Selection.Count; i++)
                {
                    Object selObject = exp.Selection[i];
                    if (selObject is MailItem)
                    {
                        MailItem mailItem =
                            (selObject as MailItem);
                        if (mailItem.Subject == null) mailItem.Subject = "***Kein Betreff***";
                        _Mails.Add(mailItem);
                    }

                }
            }
        }



        public ViewModel()
        {
            homeDrive = Environment.GetEnvironmentVariable("HomeDrive");
            test = "Hallo";
        }

        private int zahl;

        public ObservableCollection<MailItem> Mails
        {
            get { return _Mails; }
            set
            {
                _Mails = value;
                OnNotifyPropertyChanged("Mails");
            }

        }

        public string Test { get => test; set => test = value; }
        public int Zahl { get => zahl; set => zahl = value; }
        public string HomeDrive { get => homeDrive; set => homeDrive = value; }
        public MailItem Ml { get => ml; set => ml = value; }

        void OnNotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(property));
        }

    }
}
