using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace C_Excel
{
    public partial class output : Form
    {
        List<Member_QingJia.MemberChuQingStatistics> localMSs = null;
        List<String> duree = null;
        public output(List<Member_QingJia.MemberChuQingStatistics> mss)
        {
            InitializeComponent();
            localMSs = mss;
            Form1 fom = new Form1();
            duree = fom.LaDuree;
        }

        
        Member_QingJia MQJ = null;
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                if (this.Owner.MdiParent != null)
                {
                    duree = ((Form1)this.Owner.MdiParent).LaDuree;
                }
                else
                {
                    duree = ((Form1)this.Owner).LaDuree;
                }

                if (comboBox1.SelectedItem.ToString() == "TXT文件")
                {
                    StringBuilder str = new StringBuilder();
                    str.Append(duree[1] + "至" + duree[4] + "\t" + DateTime.Now.ToString());
                    str.Append(System.Environment.NewLine);

                    foreach (Member_QingJia.MemberChuQingStatistics MS in localMSs)
                    {
                        int badData = MS.BadData;
                        int dataInQuestion = MS.dataInQuestion;
                        int workerLate = MS.workerIsLate;
                        int workerNotSignOff = MS.workerNotSignOff;
                        int workerOnTime = MS.workerOnTime;
                        string workName = MS.workerName;

                        str.Append(workName + System.Environment.NewLine + "准时: " + workerOnTime + "天;\t ");
                        str.Append("迟到: " + workerLate + "天;");
                        str.Append("数据格式错误:  " + dataInQuestion + "天;\t ");
                        str.Append("无下班时间:  " + workerNotSignOff + "天;\t ");
                        str.Append("无数据: " + badData + "天;\t ");
                        str.Append(System.Environment.NewLine);
                    }
                    using (var sfd = new SaveFileDialog())
                    {
                        sfd.Filter = "文本文件(*.txt)|*.txt";
                        sfd.FileName = duree[0] + "至" + duree[4] + "考勤记录";
                        if (sfd.ShowDialog() == DialogResult.OK && sfd.FileName != "")
                        {
                            File.WriteAllText(sfd.FileName, str.ToString());
                        }
                    }

                }
                else if (comboBox1.SelectedItem.ToString() == "窗口中显示")
                {
                    MQJ = new Member_QingJia();
                    ShowTheResult(localMSs);
                }
                else if (comboBox1.SelectedItem.ToString() == "XML文件")
                {
                    MessageBox.Show("Oooops 功能未完成");
                }
                else if (comboBox1.SelectedItem.ToString() == "Excel文件")
                {
                    System.Data.DataTable ExcelTable = null;
                    ExcelTable = CreateDataSet(localMSs).Tables[0];

                    using (var sfd = new SaveFileDialog())
                    {
                        sfd.Filter = "Excel97-2003文件|*.xls;*.xlt;*.xltm|Excel2007-2010|*.xlsx|所有文件(*.*)|*.*";
                        sfd.FileName = duree[0] + "至" + duree[4] + "考勤记录统计";
                        if (sfd.ShowDialog() == DialogResult.OK && sfd.FileName != "")
                        {
                            SaveDataTableToExcel(ExcelTable, sfd.FileName);
                        }
                    }

                    //localMSs
                }

            }
        }

        private void ShowTheResult(List<Member_QingJia.MemberChuQingStatistics> MSs)
        {
            //
            TreeView treeView1 = new TreeView();
            treeView1.Size = new Size(this.Size.Width - 50, this.Size.Height - this.Location.Y - 30);
            treeView1.Location = new System.Drawing.Point(this.Location.X + 15, comboBox1.Location.Y + comboBox1.Height + 10);
            this.Controls.Add(treeView1);
            //
            foreach (Member_QingJia.MemberChuQingStatistics MS in MSs)
            {

                int badData = MS.BadData;
                int dataInQuestion = MS.dataInQuestion;
                int workerLate = MS.workerIsLate;
                int workerNotSignOff = MS.workerNotSignOff;
                int workerOnTime = MS.workerOnTime;
                string workName = MS.workerName;

                TreeNode mainNode = new TreeNode();
                mainNode.Text = workName;

                TreeNode subNode1 = new TreeNode("准时: " + workerOnTime + "天");
                subNode1.BackColor = Color.Green;

                TreeNode subNode2 = new TreeNode("迟到: " + workerLate + "天");
                subNode2.BackColor = Color.Red;

                TreeNode subNode3 = new TreeNode("数据格式错误: " + dataInQuestion + "天");
                subNode3.BackColor = Color.Yellow;

                TreeNode subNode4 = new TreeNode("无下班时间: " + workerNotSignOff + "天");
                subNode4.BackColor = Color.Violet;

                TreeNode subNode5 = new TreeNode("无数据: " + badData + "天");
                subNode5.ForeColor = Color.Silver;
                subNode5.BackColor = Color.Brown;


                mainNode.Nodes.Add(subNode1);
                mainNode.Nodes.Add(subNode2);
                mainNode.Nodes.Add(subNode3);
                mainNode.Nodes.Add(subNode4);
                mainNode.Nodes.Add(subNode5);
                treeView1.Nodes.Add(mainNode);
            }

        }

        private void output_Resize(object sender, EventArgs e)
        {
            foreach (Control c in this.Controls)
            {
                if (c is TreeView)
                {
                    c.Size = new Size(this.Size.Width - 50, this.Size.Height - c.Location.Y - 30);
                }
            }
        }

        public DataSet CreateDataSet(List<Member_QingJia.MemberChuQingStatistics> LMCQSs)
        {
            DataSet retenuDataSet = new DataSet();
            System.Data.DataTable tblDatas = new System.Data.DataTable("RateOfAttendance");
            
            tblDatas.Columns.Add("ID", Type.GetType("System.Int32"));
            tblDatas.Columns.Add("成员名", Type.GetType("System.String"));
            tblDatas.Columns.Add("准时", Type.GetType("System.Int32"));
            tblDatas.Columns.Add("迟到", Type.GetType("System.Int32"));
            tblDatas.Columns.Add("数据格式错误", Type.GetType("System.Int32"));
            tblDatas.Columns.Add("无下班时间", Type.GetType("System.Int32"));
            tblDatas.Columns.Add("无数据", Type.GetType("System.Int32"));

            int numWorker = 1;
            foreach (var Lmcqs in LMCQSs)
            {
                //tblDatas.Columns.Add(Lmcqs.workerName, Type.GetType("System.String"));

                tblDatas.Columns[0].AutoIncrement = true;
                tblDatas.Columns[0].AutoIncrementSeed = 1;
                tblDatas.Columns[0].AutoIncrementStep = 1;

                tblDatas.Rows.Add(new object[] { numWorker, Lmcqs.workerName
                    , Lmcqs.workerOnTime, Lmcqs.workerIsLate, Lmcqs.dataInQuestion
                    , Lmcqs.workerNotSignOff, Lmcqs.BadData });
                numWorker++;
            }

            retenuDataSet.Tables.Add(tblDatas);
            return retenuDataSet;
        }

        public static bool SaveDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int col = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }
                    }
                }

                int size = excelTable.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
                }
                //设置禁止弹出保存和覆盖的询问提示框 
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存工作簿 
                wBook.Save();
                //保存excel文件 
                app.Save(filePath);
                app.SaveWorkspace(filePath);
                app.Workbooks.Close();  
                app.Quit();
                System.GC.Collect();  
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
            }
            
        }

    }
}
