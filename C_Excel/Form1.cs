using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using System.Text.RegularExpressions;

//using Excel.

namespace C_Excel
{
    public partial class Form1 : Form
    {
        public class ExcelResponse
        {
            public bool IsSuccess { get; set; }
            public string Message { get; set; }
            public DataTable Item { get; set; }
        }

        public DateTime NowTime;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件.";
            fileDialog.Filter = "Excel97-2003文件|*.xls;*.xlt;*.xltm|Excel2007-2010|*.xlsx|所有文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = fileDialog.FileName;
            }

            var excelFilename = @fileName;

            var dtResponse = ReadData(excelFilename);

            if (dtResponse.IsSuccess)
            {
                lblMessage.Text = dtResponse.Message;
                if (null != dtResponse.Item)
                {
                    // LINQ to DataTable, Reading [B] col without title
                    var query = from item in dtResponse.Item.AsEnumerable()
                                //where item.Field<string>("F3") != "SHIT"
                                select item;//.Field<string>;

                    #region Business Logic

                    // Example: Add into listBox
                    if (query.Count() != 0)
                    {
                        foreach (var item in query)
                        {
                            listBox1.Items.Add(item.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("结果为空!");
                    }

                    // TODO: Your Batch Operation
                    // --

                    #endregion

                }
            }
            else
            {
                MessageBox.Show(dtResponse.Message, "错误:");
            }
        }
        
        
        /// <summary>
        /// Read first sheet in Excel 2007+ File
        /// </summary>
        /// <param name="excelFilename">Path to your Excel file</param>
        /// <returns>ExcelResponse, containing a datatable which is result</returns>
        private static ExcelResponse ReadData(string excelFilename)
        {
            try
            {
                if (!File.Exists(excelFilename))
                {
                    throw new IOException(string.Format("File {0} Not Exists!", excelFilename));
                }
                using (var conn = new OleDbConnection())
                {
                    //
                    /*if (Path.GetExtension(path) == ".xls")
            {
                oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;
                Data Source=" + path + ";
                Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"");
            }
            else if (Path.GetExtension(path) == ".xlsx")
            {
                oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source=" + path + ";
                Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");
            }*/
                    ///
                    if (Path.GetExtension(excelFilename) == ".xls")
                    {

                    }


                    conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + excelFilename + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=No\"";
                    conn.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", conn);
                    var ds = new DataSet();
                    da.Fill(ds);
                    if (null != ds.Tables[0])
                    {
                        return new ExcelResponse()
                        {
                            IsSuccess = true,
                            Message = "Query Successfully Completed",
                            Item = ds.Tables[0]
                        };
                    }
                    return new ExcelResponse
                    {
                        IsSuccess = true,
                        Message = "No Data in the Excel",
                        Item = null
                    };
                }
            }
            catch (OleDbException ex)
            {
                return new ExcelResponse
                {
                    IsSuccess = false,
                    Message = "Exception in OleDb Operation: " + ex.Message,
                    Item = null
                };
            }
            catch (Exception ex)
            {
                return new ExcelResponse
                {
                    IsSuccess = false,
                    Message = "Exception reading excel: " + ex.Message,
                    Item = null
                };
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            splitContainer1.IsSplitterFixed=false;// 1.FixedPanel=FixedPanel.Panel1
            StartTimer();
        }

        public void StartTimer()
        {
            this.timer1.Enabled = true;
            this.timer1.Start();
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            NowTime = DateTime.Now;
            lblMessage.Text = "欢迎!" + NowTime.ToString();
            toolStripStatusLabel1.Text = NowTime.ToString();
            this.Text = "通信所" + (Convert.ToInt32(NowTime.Month) - 1).ToString() + "月份考勤";
        }

      

        private void button3_Click(object sender, EventArgs e)
        {
            string FileName;

            FileName = OpenFile();

            if (FileName != "")
            {
                DataSet DS = LoadDataFromExcel(FileName);

                DataTable DT = DS.Tables[0];
                for (int i = 0; i < 4; i++)
                {
                    DT.Rows.Remove(DT.Rows[1]);
                    //;
                }

                DataTable subDT = DT.Copy();
                subDT.Clear();
                int b = 0; int c = 0;
                List<string> st = new List<string>();
                foreach (DataRow dr in DT.Rows)
                {
                    foreach (DataColumn dc in DT.Columns)
                    {
                        b++;
                        
                        List<string > MathGroup=new List<string>();
                        string a = dr[dc].ToString().Replace(" ", "");
                        //Convert.ToDateTime(textBox1.Text.Replace(" ", "").Substring(0, 8)).ToShortTimeString().ToString();
                        //if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(((20|21|22|23|[0-1]?\d):[0-5]?\d)-((20|21|22|23|[0-1]?\d):[0-5]?\d)$"))//&& isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(((20|21|22|23|[0-1]?\d):[0-5]?\d)-$") && isExMatch(dr[dc].ToString().Replace(" ", ""), @"^-((20|21|22|23|[0-1]?\d):[0-5]?\d)$")) //验证正则表达式
                        if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^((20|21|22|23|[0-1]?\d):([0-5]?\d))-((20|21|22|23|[0-1]?\d):[0-5]?\d)$", out MathGroup))
                        {
                            string temp = dr[dc].ToString();
                            dr[dc] = Convert.ToDateTime(dr[dc].ToString().Replace(" ", "").Substring(0, 5)).ToShortTimeString().ToString();
                            c++;
                            textBox2.Text = c.ToString();
                            //st.Add(MathGroup);

                        }
                        else
                        {
                            continue;
                        }
                        

                    }
                }
                DT.Columns.Remove(DT.Columns[1]);
                DT.Columns.Remove(DT.Columns[2]);
                DT.Columns.Remove(DT.Columns[5]);
                DT.Columns.Remove(DT.Columns[6]);
                dataGridView1.DataSource = DT;
            }



            /*for (int a = 0; a < 8; a++)
            {
                //subDT.Columns.Add(DT.Columns[a]);
            }

            DataColumn DIndex = DT.Columns.Add("ID", typeof(int));
            DIndex.AutoIncrement = true;
            DIndex.AutoIncrementSeed = -1;
            DIndex.AutoIncrementStep = -1;
            DIndex.ReadOnly = true;

                
            //if(DT.ta)
            dataGridView1.DataSource = subDT;
            //MessageBox.Show()
            Update();*/
        }

        public string OpenFile()
        {
            string fileName = null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件.";
            fileDialog.Filter = "Excel97-2003文件|*.xls;*.xlt;*.xltm|Excel2007-2010|*.xlsx|所有文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = fileDialog.FileName;
            }

            return fileName;
        }

        public static DataSet LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn;
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet1$]";//可是更改Sheet名称，比如sheet2，等等   

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Sheet1");
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception err)
            {
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }

        public bool isExMatch(string text, string patten, out List<string> Match)
        {
            bool _isMatch = false;
            Regex Patten = new Regex(patten);
            List<string> _match = new List<string>();
            //if (Regex.IsMatch(text, patten))
            if(Patten.Match(text).Success)
            {
                _isMatch = true;
                for (int num = 1; num < Patten.Match(text).Groups.Count; num++)
                {
                    _match.Add(Patten.Match(text).Groups[num].Value);
                }

            }
            else
                _isMatch = false;
            Match = _match;
            return _isMatch;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<string> b=new List<string>();
            string d = ""; 
            string a = textBox1.Text;
            //textBox2.Text = Convert.ToDateTime(textBox1.Text.Replace(" ", "").Substring(0, 8)).ToShortTimeString().ToString();
            //textBox2.Text=
            //if (Regex.IsMatch(textBox1.Text.Replace(" ", "").Substring(0, 4), @"^((20|21|22|23|[0-1]?\d):[0-5]?\d)$"))
            if (isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d):([0-5]?\d)-(20|21|22|23|[0-1]?\d):([0-5]?\d)$", out  b))
            {
                foreach (string c in b)
                {
                    d = d + c + "+++";
                }
                textBox2.Text = "true    " + d;
            }
            //return Regex.IsMatch(StrSource, @"^((20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d)$");
        }

    }
}