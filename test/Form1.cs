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
            textBox1.Text = @"H:\";
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
            folderBrowserDialog1.ShowDialog();
            textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = listBox1.SelectedIndices.Count - 1; i >= 0; i--)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndices[i]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string link = textBox1.Text;
            this.Close();
            foreach (Email mail in listBox1.Items)
            {
                outlook.Attachments atts = mail.getMail().Attachments;
                outlook.MailItem ml = mail.getMail();

                if (atts.Count > 0)
                {
                    switch (ml.BodyFormat)
                    {
                        case outlook.OlBodyFormat.olFormatHTML:
                            string html = "<html><body>";
                            foreach (outlook.Attachment att in atts)
                            {
                                html += "<p><a href =\"" + link + att.FileName + "\">" + att.FileName + "</a></p>";
                                att.SaveAsFile(link + att.FileName);
                            }
                            html += "</body ></html>";
                            ml.HTMLBody = html + ml.HTMLBody;
                                break;
                    }
                }
            }
        }
    }
}
