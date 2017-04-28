using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Xaml;
using Microsoft.Office.Interop.Outlook;

namespace test
{
    /// <summary>
    /// Interaktionslogik für UserControl1.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ViewModel Model = new ViewModel();
            Model.fillMails();
            this.DataContext = Model;
            InitializeComponent();

        }

        private void pickfolder_click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK) tbFolder.Text= dialog.SelectedPath;
            }
        }

        private void Button_Delete_Click(object sender, RoutedEventArgs e)
        {
            StackPanel panel = (StackPanel)((Button)sender).Parent;
            System.Windows.Forms.MessageBox.Show(panel.Parent.ToString());
            //((ObservableCollection<MailItem>)trvMails.DataContext).Remove();
        }
    }
}
