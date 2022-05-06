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
    public partial class InputExcel : Form
    {
        /// <summary>
        /// 文件路径 全路径
        /// </summary>
        public string FullStringFileName { get; set; }

        /// <summary>
        /// 是否覆盖之前的路径
        /// </summary>
        public bool OverHisData { get; set; }

        public InputExcel()
        {
            InitializeComponent();
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string fileName = this.textBox1.Text;
            if (fileName.ToLower().EndsWith(".xls") || fileName.ToLower().EndsWith(".xlsx"))
            {

            }
            else
            {
                MessageBox.Show("您选择的必须为EXCEL文件！！");
                return;
            }


            FullStringFileName = openFileDialog1.FileName;
            this.OverHisData = this.checkBox1.Checked;

            this.DialogResult = DialogResult.OK;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            textBox1.Text = openFileDialog1.FileName;
           
        }
    }
}
