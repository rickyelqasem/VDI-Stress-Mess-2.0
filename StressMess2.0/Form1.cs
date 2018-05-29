using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace StressMess2._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; 
            Int32 r =0;
            if (radioButton1.Checked)
            {
                r = DateTime.Now.Second;
                if (r <= 50)
                {
                    r = r + 50;
                }
                label3.Text = (r * 2).ToString();
                timer1.Interval = (r * 1000) * 2;
            }
            else
            {
                r = DateTime.Now.Second;
                if (r <= 50)
                {
                    r = r + 50;
                }
                label3.Text = numericUpDown1.Value.ToString();
                Int32 numpick = (Int32)numericUpDown1.Value;
                timer1.Interval = (numpick * 1000) *2 ;
            }

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Range oRng;

            oXL = new Excel.Application();
            oXL.Visible = true;

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            Int32 x;
            Int32 y;

            if (r >= 100)
                {
                    r = 100;
                }
            for (x = 1; x<r; x++)
            {
                for (y = 1; y < r; y++)
                {
                    Random random = new Random();
                    Int32 randomNumber = random.Next(0, 100);
                    Int32 second = DateTime.Now.Second;
                    randomNumber = randomNumber + second;
                    oSheet.Cells[x, y] = randomNumber.ToString();
                    oSheet.Application.Visible = true;
                }

            }
            oXL.ActiveWorkbook.Saved = true;
            oXL.Application.Quit();
            oXL.Quit();
            Word.Application wrd = new Word.Application();
            Word.Document doc = new Word.Document();
            wrd.Visible = true;
            doc = wrd.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            Word.Paragraph oPara1;
            oPara1 = doc.Content.Paragraphs.Add(ref oMissing);

            for (int i = 0; i < r; i++)
            {
                oPara1.Range.Text = "This text was inserted by Stress Mess";
                oPara1.Range.Font.Bold = 1;
                oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
                oPara1.Range.InsertParagraphAfter();
            }


            doc.Saved = true;
            doc.Close();
            wrd.Quit();


        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = true;
            label3.Text = numericUpDown1.Value.ToString();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = false;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
        }
    }
}
