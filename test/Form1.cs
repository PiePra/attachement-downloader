using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using outlook = Microsoft.Office.Interop.Outlook;

namespace test
{
    public partial class Form1 : Form
    {
        private List<Email> mls;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fillListbox();
        }

        public void setmls(List<Email> mls)
        {
            this.mls = mls;
        }

        private void fillListbox()
        {
            foreach (Email ml in mls)
                listBox1.Items.Add(ml);
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
