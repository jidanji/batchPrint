using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Management;
using System.Threading;
using System.Windows.Forms;

using System.Data.OleDb;
using System.Linq;
using DataAccess;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;

namespace WordPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            if (!Directory.Exists(path))//判断是否存在
            {
                Directory.CreateDirectory(path);//创建新路径
            }
            List<string> printer = GetAllPrinter();
            foreach (var item in printer)
            {
                this.cmbPrinter.Items.Add(item);
            }
            DbHelperAccess.connStr = string.Format(@"Provider = Microsoft.Ace.OLEDB.12.0;Data Source = {0}", Application.StartupPath + "/数据文件/Access.accdb");
            LoadData();
        }

        public void LoadData()
        {
            try
            {
                this.dataGridView1.Rows.Clear();
                string[] titleArray = new string[7] { "打印序号", "考生号", "姓名", "录取专业", "开学年", "开学月", "开学日"  };
                for (int index = 0; index < titleArray.Length; index++)
                {
                    this.dataGridView1.Columns[index].HeaderCell.Value = titleArray[index];
                }
                DataTable StudentInfo = DbHelperAccess.ExecuteDatable("select * from StudentInfo where 考生号 like '%" + this.txtXueHao.Text + "%' order by 打印序号 asc ", new OleDbParameter[0]);
                foreach (DataRow row in StudentInfo.Rows)
                {
                    this.dataGridView1.Rows.Add(
                        row[titleArray[0]].ToString(),
                        row[titleArray[1]].ToString(),
                        row[titleArray[2]].ToString(),
                        row[titleArray[3]].ToString(),
                        row[titleArray[4]].ToString(),
                        row[titleArray[5]].ToString(),
                        row[titleArray[6]].ToString()
                        );
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

    /// <summary>
    /// 文档path
    /// </summary>
    string path = System.Windows.Forms.Application.StartupPath + "/请放入要打印的文件";
        bool isStart = false;
        bool pause = false;
        Thread t = null;

        private string tupianpath = (Application.StartupPath + "/图片信息");



        private bool btnLoadFile_Click(object sender, EventArgs e)
        {
          
            int isCanResult = IsCanPrint();
            if (isCanResult == 0)
            {
                return true;
            }
            else if (isCanResult == 1)
            {
                MessageBox.Show("文件夹中不能有除doc和docx格式的其他文件！");
                return false;
            }
            else if (isCanResult == 3)
            {
                MessageBox.Show("文件夹为空！请放入doc文件。");
                btnOpenFolder_Click(new object(), new EventArgs());
                return false;
            }
            return true;
        }

        /// <summary>
        /// 检查文件格式
        /// </summary>
        /// <returns></returns>
        private int IsCanPrint()
        {
            int result = 0;
            DirectoryInfo mydir = new DirectoryInfo(path);
            FileSystemInfo[] fsis = mydir.GetFileSystemInfos();
            if (fsis.Count() <= 0)
            {
                return 3;
            }
            foreach (FileSystemInfo fsi in fsis)
            {
                if (fsi is FileInfo)
                {
                    FileInfo fi = (FileInfo)fsi;
                    string[] fileNamearr = fsi.Name.Split('.');
                    if (fileNamearr.LastOrDefault() != "doc" && fileNamearr.LastOrDefault() != "docx")
                    {
                        result = 1;
                        break;
                    }
                }
            }
            return result;
        }

        

        /// <summary>
        /// 开始 启动一个线程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!btnLoadFile_Click(new object(), new EventArgs())) {
                return;
            }
            if (!string.IsNullOrEmpty(this.cmbPrinter.Text))
            {
                isStart = true;
                t = new Thread(new ThreadStart(PrintMain));
                t.Start();
                string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印开始……";
                listBoxLog.Items.Add(message);
                LogManage.WriteLog(LogManage.LogFile.Trace, message);
                this.btnStart.Enabled = false;
                this.btnBreak.Enabled = true;
                 
            }
            else
            {
                MessageBox.Show("请选择一个打印机！");
                this.cmbPrinter.DroppedDown = true;
            }
        }

        /// <summary>
        /// 停止方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBreak_Click(object sender, EventArgs e)
        {
            isStart = false;
            this.btnStart.Enabled = true;
            this.btnBreak.Enabled = false;
            string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "程序终止……";
            listBoxLog.Items.Add(message);
            LogManage.WriteLog(LogManage.LogFile.Trace, message);
        }

        private void WritePauseLog()
        {
            while (pause)
            {
                DateTime now = DateTime.Now;
                now = now.AddMinutes(Convert.ToInt32(txtRestTime));
                string message = now.ToString("MM-dd HH:mm:ss") + "程序暂停中：" + now.Subtract(DateTime.Now).ToString();
                this.listBoxLog.Items.Add(message);
                LogManage.WriteLog(LogManage.LogFile.Trace, message);
            }
        }

        /// <summary>
        /// 打印主线程方法
        /// </summary>
        private void PrintMain()
        {
            try
            {
                DirectoryInfo mydir = new DirectoryInfo(path);
                FileSystemInfo[] fsis = mydir.GetFileSystemInfos();
                int i = 0;
                foreach (FileSystemInfo fsi in fsis)
                {
                    if (i != 0 && i % Convert.ToInt32(txtCount.Text) == 0)
                    {
                        listBoxLog.Items.Add(DateTime.Now.ToString("MM-dd HH:mm:ss") + "已经打印" + i + "正在休息");
                        Thread.Sleep(Convert.ToInt32(txtRestTime.Text) * 1000);
                    }
                    i++;
                    if (!isStart)
                        break;
                    if (pause)
                    {
                        pause = false;
                    }
                    if (fsi is FileInfo)
                    {
                        FileInfo fi = (FileInfo)fsi;
                        PrintHelper printHelper = new PrintHelper();
                        printHelper.PrintMethodOther(fi.FullName, cmbPrinter.Text);
                        string message = DateTime.Now.ToString("MM-dd HH:mm:ss") + "正在打印" + fi.Name + "……";
                        listBoxLog.Items.Add(message);
                        LogManage.WriteLog(LogManage.LogFile.Trace, message);
                    }
                    listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;

                }
                listBoxLog.Items.Add(DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                LogManage.WriteLog(LogManage.LogFile.Trace, DateTime.Now.ToString("MM-dd HH:mm:ss") + "打印结束。");
                this.btnStart.Enabled = true;
                this.btnBreak.Enabled = false;
                listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogManage.WriteLog(LogManage.LogFile.Error, ex.Message);
            }

        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            WritePauseLog();
            pause = true;
            this.btnBreak.Enabled = false;
            this.btnStart.Enabled = true;
           
        }


        /// <summary>
        /// 获取所有打印机
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllPrinter()
        {
            ManagementObjectCollection queryCollection;
            string _classname = "SELECT * FROM Win32_Printer";
            Dictionary<string, ManagementObject> dict = new Dictionary<string, ManagementObject>();
            ManagementObjectSearcher query = new ManagementObjectSearcher(_classname);
            queryCollection = query.Get();
            List<string> result = new List<string>();
            foreach (ManagementObject mo in queryCollection)
            {
                string oldName = mo["Name"].ToString();
                result.Add(oldName);
            }
            return result;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            this.listBoxLog.Items.Clear();
        }

  

        private void 数据导入ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            InputExcel f1 = new WordPrint.InputExcel();
            f1.ShowDialog();
            if (f1.DialogResult == DialogResult.OK)
            {
                try
                {
                    string fileName = f1.FullStringFileName;
                    bool isOverHisData = f1.OverHisData;
                    if (isOverHisData)
                    {
                        string sql = "delete from StudentInfo";
                        DbHelperAccess.ExecuteNonQuery(sql, new OleDbParameter[0]);
                    }
                    DataTable dt1 = ExcelHelper.TransExcelToTable(fileName, "Sheet1");
                    foreach (DataRow row in dt1.Rows)
                    {
                        string sql = string.Format(@"insert into StudentInfo
(打印序号,
考生号,
姓名,
录取专业,
开学年,
开学月,
开学日
)values('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",
row["打印序号"].ToString(),
row["考生号"].ToString(),
row["姓名"].ToString(),
row["录取专业"].ToString(),
row["开学年"].ToString(),
row["开学月"].ToString(),
row["开学日"].ToString()
);
                        DbHelperAccess.ExecuteNonQuery(sql, new OleDbParameter[0]);
                        this.LoadData();
                    }
                    MessageBox.Show(string.Format("导入{0}条纪录！", dt1.Rows.Count));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

 


        private void btnOpenFolder_Click(object sender, EventArgs e)
        {

        }


 

        private void button1_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(System.Windows.Forms.Application.StartupPath + "/数据模板");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "数据处理错误，请仔细检查");
            }
          
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(System.Windows.Forms.Application.StartupPath + "/图片信息");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "数据处理错误，请仔细检查");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DbHelperAccess.ExecuteNonQuery("delete * from StudentInfo where 考生号 like '%" + this.txtXueHao.Text + "%'", new OleDbParameter[0]);
            this.LoadData();
        }

        public void Toprint()
        {
            int succeedCount = 0;
            int failCount = 0;
            List<string> failList   = new List<string>();
            List<string> failReasonList = new List<string>();
            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                try
                {
                    var imgurl = this.tupianpath + "/" + $"{row.Cells[1].Value.ToString()}.jpg";
                    if (!File.Exists(imgurl))
                    {
                        throw new Exception("文件不存在");
                    }
                    WriteIntoWord word = new WriteIntoWord();
                    string filePath = System.Windows.Forms.Application.StartupPath + "/数据模板/template.dot";
                    string[] titleArray = new string[6] { "考生号", "姓名", "录取专业", "开学年", "开学月", "开学日" };
                    object[] objArray1 = new object[] { Application.StartupPath, "/请放入要打印的文件/", row.Cells[1].Value, ".doc" };
                    string saveDocPath = string.Concat(objArray1);
                    word.OpenDocument(filePath);
                    try
                    {
                        word.WriteIntoDocument("Pic", imgurl);
                    }
                    catch (Exception err1)
                    {
                        word.Save_CloseDocument(saveDocPath);
                        throw err1;
                    }

                    for (int index = 0; index < titleArray.Length; index++)
                    {
                        try
                        {
                            var key = titleArray[index];
                            var val = row.Cells[index + 1].Value.ToString();
                            word.WriteIntoDocument(key, val);
                        }
                        catch (Exception err)
                        {
                            word.Save_CloseDocument(saveDocPath);
                            throw err;
                        }
                    }
                    word.Save_CloseDocument(saveDocPath);
                    succeedCount++;

                }
                catch (Exception ex1)
                {
                    failCount++;
                    failReasonList.Add(ex1.Message);
                    failList.Add(row.Cells[1].Value.ToString());
                }
            }
            Form2 form2 = new Form2(succeedCount, failCount, failList, failReasonList);
            form2.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                FileInfo[] files = new DirectoryInfo(Application.StartupPath + "/请放入要打印的文件/").GetFiles();
                for (int i = 0; i < files.Length; i++)
                {
                    files[i].Delete();
                }
                this.Toprint();
                Process.Start(this.path);
                
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void 数据导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(System.Windows.Forms.Application.StartupPath + "/请放入要打印的文件");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
    }
}


 
