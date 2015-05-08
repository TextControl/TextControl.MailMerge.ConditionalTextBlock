using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tx_mailmerge_condIF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings();
            ls.ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord;

            textControl1.Load("template.docx", TXTextControl.StreamType.WordprocessingML, ls);
        }

        private void createToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.ReadXml("database.xml", XmlReadMode.Auto);

            mailMerge1.Merge(ds.Tables[0]);
        }

        private void mailMerge1_DataRowMerged(object sender, TXTextControl.DocumentServer.MailMerge.DataRowMergedEventArgs e)
        {
            byte[] data;
            string sPlaceholderStartIF = "%%StartIF%%";
            string sPlaceholderEndIF = "%%EndIF%%";

            using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
            {
                tx.Create();
                tx.Load(e.MergedRow, TXTextControl.BinaryStreamType.InternalUnicodeFormat);

                foreach (TXTextControl.IFormattedText part in tx.TextParts)
                {
                    do
                    {
                        int start = part.Find(sPlaceholderStartIF, 0, TXTextControl.FindOptions.NoMessageBox);
                        int end = part.Find(sPlaceholderEndIF, start, TXTextControl.FindOptions.NoMessageBox);

                        if (start == -1)
                            continue;

                        part.Selection.Start = start;
                        part.Selection.Length = end - start + sPlaceholderEndIF.Length;
                        part.Selection.Text = "";
                    } while (part.Find(sPlaceholderStartIF, 0, TXTextControl.FindOptions.NoMessageBox) != -1);
                }

                tx.Save(out data, TXTextControl.BinaryStreamType.InternalUnicodeFormat);
            }

            e.MergedRow = data;
        }
    }
}
