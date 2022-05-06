using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordPrint
{
    
    public partial class Form2 : Form
    {
        private int succeedCount;
        private int failCount;
        private List<string> failList;

        private List<string> failReasonList;

        


        public Form2()
        {
            this.succeedCount = 0;
            this.failCount = 0;
            this.failList = new List<string> ();

            this.failReasonList = new List<string>();
            InitializeComponent();
        }

        public Form2(int succeedCount, int failCount, List<string> failList, List<string> failReasonList)
        {
            this.succeedCount = succeedCount;
            this.failCount = failCount;
            this.failList = failList;
            this.failReasonList = failReasonList;
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.label3.Text = this.succeedCount.ToString();
            this.label1.Text = this.failCount.ToString();
            this.richTextBox1.Text = string.Join("\r\n", this.failList.ToArray<string>());
            this.richTextBox2.Text = string.Join("\r\n", this.failReasonList.ToArray<string>());
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
