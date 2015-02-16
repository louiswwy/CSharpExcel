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

        #region 封装早晨和下午时间
        public class AMTime
        {
            private DateTime _amTime;
            public DateTime amTime
            {
                get { return _amTime; }
                set { _amTime = value; }
            }

            public AMTime(DateTime AmTime)
            {
                this.amTime = AmTime;
            }
        }

        public class PMTime
        {
            private DateTime _pmTime;
            public DateTime pmTime
            {
                get { return _pmTime; }
                set { _pmTime = value; }
            }

            public PMTime(DateTime PmTime)
            {
                this.pmTime = PmTime;
            }
        }
        #endregion


        public class WorkTime
        {
            private AMTime _amTime;
            public AMTime amTime
            {
                get { return this._amTime; }
                set { this._amTime = value; }
            }

            private PMTime _pmTime;
            public PMTime pmTime
            {
                get { return this._pmTime; }
                set { this._pmTime = value; }
            }

            public WorkTime(PMTime PmTime)
            {
                this.pmTime = pmTime;
            }
            public WorkTime(AMTime AmTime)
            {
                this.amTime = AmTime;
            }

            public WorkTime(AMTime AmTime, PMTime PmTime)
            {
                this.amTime = AmTime;
                this.pmTime = pmTime;
            }
        }
        public class Member_Communications
        {
            private string _name;
            public string name{
                get { return this._name; }
                set { this._name = value; }
            }

            private List<WorkTime> _workTime;
            public List<WorkTime> workTime
            {
                get { return this._workTime; }
                set { this._workTime = value; }
            }

            public Member_Communications()
            {                
            }
            //private List<DateTime> _
            public Member_Communications(string Name)
            {
                this.name = Name;
            }

            public Member_Communications(string Name, List<WorkTime> WorkTime)
            {
                this.name = Name;
                this.workTime = WorkTime;
            }
        }

        public List<WorkTime> listWorkTime; //上班时间
        public static List<Member_Communications> MemberSchedules=new List<Member_Communications>();
        //本地电脑时间.
        public DateTime NowTime;

        //最晚上班时间
        public DateTime LimitShowUpTime = Convert.ToDateTime("08:46:00");
        //最早下班时间
        public DateTime LimitDismissTime = Convert.ToDateTime("17:30:00");
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
                    conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + excelFilename + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1 ";
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

        //在界面底部显示现在时间,每秒刷新
        private void timer1_Tick(object sender, EventArgs e)
        {
            NowTime = DateTime.Now;
            lblMessage.Text = "欢迎!" + NowTime.ToString();
            toolStripStatusLabel1.Text = NowTime.ToString();
            this.Text = "通信所" + (Convert.ToInt32(NowTime.Month) - 1).ToString() + "月份考勤";
        }

      

        private void button3_Click(object sender, EventArgs e)
        {
            string FilePath = "";
            //文件路径
            FilePath = OpenFile();

            try
            {
                if (FilePath != "")
                {
                    DataSet DS = LoadDataFromExcel(FilePath);

                    DataTable DT = DS.Tables[0];

                    List<string> st = new List<string>();


                    List<Member_Communications> ListMemberSchedule = new List<Member_Communications>(); //全员列表
                    List<string> _checkedMemberName = new List<string>(); //已遍历过得员工的名字列表
                    List<string> MemberName = new List<string>();

                    foreach (DataRow dr in DT.Rows)
                    {
                        foreach (DataColumn dc in DT.Columns)
                        {
                            List<string> MathGroup = new List<string>();
                            string a = dr[dc].ToString().Replace(" ", "");

                            bool _begin = false;
                            //尚未便利过任何员工时 和 检测到的员工名称与列表中的最后一个不同
                            if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{3})$", out MemberName) || _begin == true)// && MemberName[0] != "通信所" && _appMemberName.Count == 0)
                            //|| (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out MemberName) && MemberName[0] != "通信所" && _appMemberName[_appMemberName.Count - 1] != MemberName[0]))
                            {
                                _begin = true;
                                if (MemberName.Count != 0 && MemberName[0] != "通信所")// &&  (_appMemberName.Count == 0||_appMemberName[_appMemberName.Count - 1] != MemberName[0]))
                                {
                                    MessageBox.Show("" + MemberName[0]);

                                    if (_checkedMemberName.Count == 0)//|| _checkedMemberName[_checkedMemberName.Count - 1].name != MemberName[0])
                                    {
                                        listWorkTime = new List<WorkTime>();
                                        _checkedMemberName.Add(MemberName[0]);
                                        Member_Communications _memberSchedule = new Member_Communications(MemberName[0]);
                                        MemberSchedules.Add(_memberSchedule);
                                    }
                                    else if (MemberSchedules.Count != 0 && MemberSchedules[MemberSchedules.Count - 1].name.ToString() != MemberName[0])
                                    {
                                        //_checkedMemberName[_checkedMemberName.Count - 1].workTime=lis
                                        Member_Communications _memberSchedule = new Member_Communications(MemberName[0]);
                                        MemberSchedules[MemberSchedules.Count - 1].workTime = listWorkTime;
                                        MemberSchedules.Add(_memberSchedule);
                                        listWorkTime.Clear();
                                    }
                                }
                            }
                                    //实例化MemberSchedule;


                                    #region usefull

                                    //当时间格式为xx:xx-yy:yy时
                                    if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))
                                    /*|| isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d$", out MathGroup)
                                    || isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out MathGroup)
                                    || isExMatch(dr[dc].ToString().Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))*/
                                    {
                                        //foreach (string str in MathGroup)
                                        {
                                            AMTime _amt1 = new AMTime(Convert.ToDateTime(Convert.ToDateTime(MathGroup[0]).ToShortTimeString()));
                                            PMTime _pmt1 = new PMTime(Convert.ToDateTime(Convert.ToDateTime(MathGroup[1]).ToShortTimeString()));
                                            WorkTime _workTime = new WorkTime(_amt1, _pmt1);
                                            listWorkTime.Add(_workTime);
                                        }
                                    }

                                    //时间格式为xx:xx:xx时 //^[\u4e00-\u9fa5]{3}
                                    else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d[\u4e00-\u9fa5]{0,4}$", out MathGroup))
                                    {
                                        //foreach (string str in MathGroup) 
                                        {
                                            MessageBox.Show("1");
                                            dr[dc] = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                                            AMTime _amt1 = new AMTime(Convert.ToDateTime(Convert.ToDateTime(MathGroup[0]).ToShortTimeString()));
                                            WorkTime _workTime = new WorkTime(_amt1);
                                            listWorkTime.Add(_workTime);
                                        }
                                    }
                                    //时间格式为xx:xx-
                                    else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out MathGroup))
                                    {
                                        //foreach (string str in MathGroup)
                                        {
                                            dr[dc] = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                                            AMTime _amt1 = new AMTime(Convert.ToDateTime(Convert.ToDateTime(MathGroup[0]).ToShortTimeString()));
                                            WorkTime _workTime = new WorkTime(_amt1);
                                            listWorkTime.Add(_workTime);
                                        }
                                    }
                                    else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))
                                    {
                                        //foreach (string str in MathGroup)
                                        {
                                            dr[dc] = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                                            PMTime _pmt1 = new PMTime(Convert.ToDateTime(Convert.ToDateTime(MathGroup[0]).ToShortTimeString()));
                                            WorkTime _workTime = new WorkTime(_pmt1);
                                            listWorkTime.Add(_workTime);
                                        }
                                    }
                                    //if (isExMatch(textBox1.Text.Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out b))
                                    /*else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out MathGroup))
                                    {
                                        dr[dc] = "星期" + MathGroup[1];
                                    }*/
                                    else if (dr[dc].ToString().Replace(" ", "") == "")
                                    {
                                        dr[dc] = dr[dc];
                                    }
                                    else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^[0].\d*$", out MathGroup))
                                    {
                                        MessageBox.Show("单元格样式需要修改!", "注意", MessageBoxButtons.OK);
                                        break;
                                        //dr[dc] = "-+-" +Convert.ToDateTime(Convert.ToDateTime(dr[dc])).ToString() + "--error--" "-+-";
                                    }
                                    //u4E00-\u9FA5 与上级重复
                                    /*else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{3})$", out MathGroup))
                                    {
                                        dr[dc] = dr[dc] + "-" + MathGroup[0];
                                    }*/

                                    else
                                    {
                                        continue;
                                    }
                                    #endregion
                                //}
                            //}

                        }
                    }
                    dataGridView1.DataSource = DT;

                    WorkingPassion(MemberSchedules);
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
            catch (Exception ex)
            {

                MessageBox.Show("" + ex);
            }
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
                //fileType=fileDialog.f
            }

            return fileName;
        }

        public static DataSet LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn;
                //         Provider=Microsoft.Ace.OleDb.12.0;"  Provider=Microsoft.Jet.OLEDB.4.0                     12/8
                strConn = "Provider=Microsoft.Ace.OleDb.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
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

        public bool CompareTime(DateTime time1, DateTime time2)
        {
            bool _isLater = false;
            //time1比time2晚时
            if(DateTime.Compare(time1,time2)>=0)
            {
                _isLater = true;
            }
            return _isLater;
        }

        //考勤情况
        public void WorkingPassion(List<Member_Communications> MCs)
        {    
            
            string StatueOfEmployer = null;

            foreach (Member_Communications mc in MCs)
            {
                string EmployerName = mc.name;
                List<WorkTime> _listWorkTime = new List<WorkTime>();
                _listWorkTime = mc.workTime;
                int _inTime = 0;
                int _noSignOff = 0;
                int _late = 0;
                int _question = 0;
                foreach (WorkTime wt in _listWorkTime)
                {
                    if (wt.amTime != null && wt.pmTime != null)
                    {
                        if (CompareTime(Convert.ToDateTime(wt.amTime.ToString()), LimitShowUpTime))
                        {
                            _late++;
                        }
                        else
                        {
                            _inTime++;
                        }
                    }
                    else if (wt.amTime == null)
                    {
                        _question++;
                    }
                    else if (wt.pmTime == null)
                    {
                        _noSignOff++;
                    }
                }
                string _statueOfEmployer = EmployerName + "\n \t准时到达:" + _inTime + "次.\n \t迟到:" + _late + "次. \n \t无早上数据:" + _question + "次. \n \t无下午数据:" + _noSignOff + "次";
                StatueOfEmployer = StatueOfEmployer + _statueOfEmployer;
            }
            textBox2.Text = StatueOfEmployer;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<string> b=new List<string>();
            string d = ""; 
            string a = textBox1.Text;
            //textBox2.Text = Convert.ToDateTime(textBox1.Text.Replace(" ", "").Substring(0, 8)).ToShortTimeString().ToString();
            //textBox2.Text=
            //if (Regex.IsMatch(textBox1.Text.Replace(" ", "").Substring(0, 4), @"^((20|21|22|23|[0-1]?\d):[0-5]?\d)$"))
            if (isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b) 
                //|| isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d:[0-5]?\d)$", out b) 
                || isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out b)
                || isExMatch(textBox1.Text.Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b))
            {
                foreach (string c in b)
                {
                    d = d + c + "+++";
                }

                CompareTime(Convert.ToDateTime(b[0]), Convert.ToDateTime(b[1]));
                //textBox2.Text = "true    " + d;
            }
            else if (isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d$", out b))
            {
                textBox2.Text = "true    " + b[0];

            }
            else if (isExMatch(textBox1.Text.Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out b))
            {
                textBox2.Text = "true    " + b[0];
            }
            //
            else if (isExMatch(textBox1.Text.Replace(" ", ""), @"(^[\u4e00-\u9fa5]{3})$", out b))
            {
                textBox2.Text = "true    " + b[0];
            }
            //return Regex.IsMatch(StrSource, @"^((20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d)$");
        }

    }
}