﻿using System;
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

using System.Xml;

using System.Text.RegularExpressions;
using System.Data.Odbc;
using System.Reflection;

/*using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;
using NPOI.SS.UserModel;
*/
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
            private TimeSpan _am_Time;
            public TimeSpan amTime
            {
                get { return _am_Time; }
                set { _am_Time = value; }
            }

            public AMTime(TimeSpan AmTime)
            {
                this.amTime = AmTime;
            }
        }

        public class PMTime
        {
            private TimeSpan _pmTime;
            public TimeSpan pmTime
            {
                get { return _pmTime; }
                set { _pmTime = value; }
            }

            public PMTime(TimeSpan PmTime)
            {
                this.pmTime = PmTime;
            }
        }
        #endregion

        private List<string> _laDuree;
        public List<string> LaDuree
        {
            get { return this._laDuree; }
            set { this._laDuree = value; }
        }

        public bool DataIsSet = true;

        public class WorkTime
        {
            private List<string> _stringDate;
            public List<string> _Date
            {
                get { return this._stringDate; }
                set { this._stringDate = value; }
            }
            private AMTime _time_am;
            public AMTime amTime
            {
                get { return this._time_am; }
                set { this._time_am = value; }
            }

            private PMTime _time_pm;
            public PMTime pmTime
            {
                get { return this._time_pm; }
                set { this._time_pm = value; }
            }

            public WorkTime()
            {
            }
            public WorkTime(List<string> StringDate)
            {
                this._Date = StringDate;
            }

            public WorkTime(List<string> StringDate, PMTime PmTime)
            {
                this._Date = StringDate;
                this.pmTime = PmTime;
            }
            public WorkTime(List<string> StringDate, AMTime AmTime)
            {
                this._Date = StringDate;
                this.amTime = AmTime;
            }

            public WorkTime(List<string> StringDate, AMTime AmTime, PMTime PmTime)
            {
                this._Date = StringDate;
                this.amTime = AmTime;
                this.pmTime = PmTime;
            }
        }

        public class Member_Departement_Communications
        {
            private string _name;
            public string name
            {
                get { return this._name; }
                set { this._name = value; }
            }

            private List<WorkTime> _workTime;
            public List<WorkTime> workTime
            {
                get { return this._workTime; }
                set { this._workTime = value; }
            }


            public Member_Departement_Communications()
            {
            }
            //private List<DateTime> _
            public Member_Departement_Communications(string Name)
            {
                this.name = Name;
            }

            public Member_Departement_Communications(string Name, List<WorkTime> WorkTime)
            {
                this.name = Name;
                this.workTime = WorkTime;
            }
        }

        public class WorkTimeState : WorkTime
        {

            private string _leaveReason;
            public string leaveReason
            {
                get { return _leaveReason; }
                set { this._leaveReason = value; }
            }

            public WorkTimeState(List<string> StringDate, string LeaveReason)
            {
                this._Date = StringDate;
                this.leaveReason = LeaveReason;
            }

            public WorkTimeState(List<string> StringDate, PMTime PmTime, string LeaveReason)
            {
                this._Date = StringDate;
                this.pmTime = PmTime;
                this.leaveReason = LeaveReason;
            }
            public WorkTimeState(List<string> StringDate, AMTime AmTime, string LeaveReason)
            {
                this._Date = StringDate;
                this.amTime = AmTime;
                this.leaveReason = LeaveReason;
            }

            public WorkTimeState(List<string> StringDate, AMTime AmTime, PMTime PmTime, string LeaveReason)
            {
                this._Date = StringDate;
                this.amTime = AmTime;
                this.pmTime = PmTime;
                this.leaveReason = LeaveReason;
            }
            /*public WorkTimeState(WorkTime WorkTime, string LeaveReason)
            {
                this.workTime = WorkTime;
                this.leaveReason = LeaveReason;
            }*/
        }

        public class Member_Dep_Com_WorkingState : Member_Departement_Communications
        {
            private int isLate;
            private int onTime;
            private int inQuestion;
            private int notSignOff;

            private List<WorkTimeState> _workTimeState;
            public List<WorkTimeState> workTimeState
            {
                get { return this._workTimeState; }
                set { this._workTimeState = value; }
            }

            public int workerIsLate
            {
                get { return isLate; }
                set { isLate = value; }
            }

            public int workerOnTime
            {
                get { return onTime; }
                set { onTime = value; }
            }
            public int dataInQuestion
            {
                get { return inQuestion; }
                set { inQuestion = value; }
            }
            public int workerNotSignOff
            {
                get { return notSignOff; }
                set { notSignOff = value; }
            }

            public Member_Dep_Com_WorkingState(string WorkerName, List<WorkTimeState> WorkTime_State, int WorkerIsLate, int WorkerOnTime, int DataInQuestion, int WorkerNotSignOff)
            {
                this.name = WorkerName;
                this.workTimeState = WorkTime_State;
                this.workerIsLate = WorkerIsLate;
                this.workerOnTime = WorkerOnTime;
                this.dataInQuestion = DataInQuestion;
                this.workerNotSignOff = WorkerNotSignOff;
            }
        }

        FunctionsCS fcs = new FunctionsCS();

        public List<WorkTime> listWorkTime; //上班时间

        /*
        public List<int> ListNotEmptyCol = new List<int>();
        */
        public List<Member_Departement_Communications> ListMemberSchedule;//=new List<Member_Departement_Communications>();
        //ListMemberSchedule
        //本地电脑时间.
        public DateTime NowTime;

        //最晚上班时间
        public static DateTime _limitShowUpTime;// = Convert.ToDateTime("08:46:00");       
        public DateTime SetLimShowUpTime
        {
            get { return _limitShowUpTime; }
            set { _limitShowUpTime = value; }
        }

        //最早下班时间
        public static DateTime _limitDismissTime;// = Convert.ToDateTime("17:30:00");
        public DateTime SetLimDissmisTime
        {
            get { return _limitDismissTime; }
            set { _limitDismissTime = value; }
        }

        public Form1()
        {
            InitializeComponent();
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
                    oledbConn = new OleDbConnection("Provider=Micrsoft.Jet.OLEDB.4.0;
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
            外勤ToolStripMenuItem.Enabled = false;
            LaDuree = new List<string>();
            this.Text = "通信所考勤记录";
            //splitContainer1.IsSplitterFixed = false;// 1.FixedPanel=FixedPanel.Panel1
            StartTimer();
            LoadWorkTime();
        }


        #region Timer
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
            toolStripStatusLabel1.Text = "";
            //this.Text = "通信所考勤记录";
        }
        #endregion

        string start;
        string end;

        string startY;
        string startM;
        string StartD;

        string endY;
        string endM;
        string endD;

        //当遇到名字或者 "11 一 "类似的格式时.
        private void button3_Click(object sender, EventArgs e)
        {
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
            fileDialog.Dispose();
            return fileName;
        }

        public static DataSet LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn;
                //         Provider=Microsoft.Ace.OleDb.12.0;"  Provider=Microsoft.Jet.OLEDB.4.0                     12/8
                strConn = "Provider=Microsoft.ACE.OlEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";

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

        #region 处理xml
        public void LoadWorkTime()
        {
            try
            {
                var folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                //var configPath = Path.Combine(folderPath, "DataFile");

                string pathXmlFile = "";
                //configPath + @"\WorkTime.xml"

#if DEBUG
                pathXmlFile = @"..\..\DataFile\WorkTime.xml";
#else
                pathXmlFile = folderPath + @"\WorkTime.xml";
#endif
                StringBuilder sttr=new StringBuilder();

                XmlDocument xmldoc = new XmlDocument();
                if (System.IO.File.Exists(pathXmlFile))
                {
                    sttr.Append("-1-");
                    sttr.Append(pathXmlFile);
                    xmldoc.Load(pathXmlFile);    //读取指定的XML文档                    
                }
                else
                {
                    sttr.Append("-2-");
                    sttr.Append(pathXmlFile);
                    CreateXml(xmldoc, pathXmlFile);//如果程序没有找到xml文件，则新建一个
                }


                XmlNode NodeWorkTime = xmldoc.DocumentElement;  //读取xml的根节点

                foreach (XmlNode node in NodeWorkTime.ChildNodes)//循环子节点
                {
                    switch (node.Name)
                    {
                        case "ShowUpTime":
                            if (node.InnerText != "")
                            {
                                this.SetLimShowUpTime = Convert.ToDateTime(node.InnerText);
                            }
                            else
                            {
                                node.InnerText = "8:46";
                                xmldoc.Save(pathXmlFile);
                            }

                            break;

                        case "dissmisTime":
                            if (node.InnerText != "")
                            {
                                this.SetLimDissmisTime = Convert.ToDateTime(node.InnerText);
                            }
                            else
                            {
                                node.InnerText = "17:30";
                                xmldoc.Save(pathXmlFile);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex, "错误");
            }
        }

        //创建WorkTime.xml文件
        public void CreateXml(XmlDocument xmldoc, string path)
        {
            //xmldoc = new XmlDocument();
            XmlNode xmlnode = xmldoc.CreateNode(XmlNodeType.XmlDeclaration, "", "");//加入XML的声明段落
            xmldoc.AppendChild(xmlnode);
            XmlElement xmlelem = xmldoc.CreateElement("workTime");

            XmlElement xmlChilElem1 = xmldoc.CreateElement("ShowUpTime");
            XmlElement xmlChilElem2 = xmldoc.CreateElement("dissmisTime");
            xmlelem.AppendChild(xmlChilElem1);
            xmlelem.AppendChild(xmlChilElem2);

            xmldoc.AppendChild(xmlelem);
            xmldoc.Save(path);// @"..\..\DataFile\WorkTime.xml");
        }

        /*public void XmlMemberList()
        {
            try
            {
                XmlDocument xmldoc = new XmlDocument();
                try
                {
                    xmldoc.Load(@"..\..\DataFile\ListMemberName.xml");    //读取指定的XML文档                    
                }
                catch (Exception)
                {
                    CreateXml(xmldoc, @"..\..\DataFile\ListMemberName.xml");//如果程序没有找到xml文件，则新建一个
                }
                XmlNode NodeWorkTime = xmldoc.DocumentElement;  //读取xml的根节点

                foreach (XmlNode node in NodeWorkTime.ChildNodes)//循环子节点
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex, "错误");
            }
        }*/
        #endregion



        //比较两个时间的早晚, 测试时间比时限大时返回true
        public bool CompareTime(TimeSpan time1, TimeSpan TimeLimit)
        {
            bool _isLater = false;
            //time1比time2晚时
            if (time1 >= TimeLimit)
            {
                _isLater = true;
            }
            return _isLater;
        }

        //考勤情况
        public void WorkingPassion(List<Member_Departement_Communications> MCs)
        {

            string StateOfEmployer = null;

            foreach (Member_Departement_Communications mc in MCs)
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
                        //将数据转换为timeSpan格式


                        DateTime _timeLim = this.SetLimShowUpTime;
                        TimeSpan ts = _timeLim.TimeOfDay;

                        if (CompareTime(wt.amTime.amTime, ts))
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
                string _stateOfEmployer = EmployerName + System.Environment.NewLine + "\t准时到达:"
                    + _inTime + "次." + System.Environment.NewLine + " \t迟到:" + _late + "次." + System.Environment.NewLine + " \t无早上数据:"
                    + _question + "次. " + System.Environment.NewLine + "\t无下午数据:" + _noSignOff + "次. " + System.Environment.NewLine + System.Environment.NewLine;
                StateOfEmployer = StateOfEmployer + _stateOfEmployer;
            }
            //textBox2.Text = StateOfEmployer;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            string FilePath = "";
            //文件路径
            FilePath = OpenFile();

            PrintData(FilePath);
        }

        public List<int> FindColNumber()
        {
            return null;
        }

        public void PrintData(string filePath)
        {
            OdbcConnection conn = this.GetConnection(filePath);
            //查询语句，就是SQL语句嘛
            string strComm = "select * from [Sheet1$]";
            //创建查询命令，也很熟悉吧
            OdbcCommand comm = new OdbcCommand(strComm, conn);
            //别忘了，访问Excel也是要打开连接的
            conn.Open();
            //Reader这个类就再熟悉不过了吧，和SqlDataReader基本上是一样的
            OdbcDataReader reader = comm.ExecuteReader();
            //Console.WriteLine("姓名\t学号\t年龄\t性别");

            //读取Reader中的数据，打印到屏幕上
            if (reader != null)
            {
                while (reader.Read())
                {
                    StringBuilder strLine = new StringBuilder();
                    for (int i = 0; i < reader.FieldCount; ++i)
                    {
                        strLine.Append(reader[i].ToString() + "\t");
                    }
                    Console.WriteLine(strLine.ToString());
                }
            }
        }

        private OdbcConnection GetConnection(string FilePath)
        {
            //连接字符串
            //string strConn = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=D:\\test.xls;DefaultDir=c:\\mypath";
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
            //创建连接，和SQL Server差不多，就是SqlConnection变成了OdbcConnection
            OdbcConnection conn = new OdbcConnection(strConn);
            return conn;
        }



        private void B_Calendar_Click(object sender, EventArgs e)
        {

        }

        private void B_Test_Click(object sender, EventArgs e)
        {
            List<string> b = new List<string>();
            string d = "";
            string a = textBox1.Text;
            //textBox2.Text = Convert.ToDateTime(textBox1.Text.Replace(" ", "").Substring(0, 8)).ToShortTimeString().ToString();
            //textBox2.Text=
            //if (Regex.IsMatch(textBox1.Text.Replace(" ", "").Substring(0, 4), @"^((20|21|22|23|[0-1]?\d):[0-5]?\d)$"))
            if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b)
                //|| isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d:[0-5]?\d)$", out b) 
                || fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b)
                || fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b))
            {
                foreach (string c in b)
                {
                    d = d + c + "+++";

                }
                //textBox2.Text = d;
                //CompareTime(Convert.ToDateTime(b[0]), Convert.ToDateTime(b[1]));
                //textBox2.Text = "true    " + d;
            }
            //^(([1-9]{1})|([0-1][0-9])|([1-2][0-3])):([0-5][0-9])$

            //else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^(([0-1]?[0-9])|([2][0-3])):([0-5]?[0-9])(:([0-5]?[0-9]))?-(([0-1]?[0-9])|([2][0-3])):([0-5]?[0-9])(:([0-5]?[0-9]))?$", out b))
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^([1-9]{1}|[0-1][0-9]|[1-2][0-3]):([0-5][0-9])-([1-9]{1}|[0-1][0-9]|[1-2][0-3]):([0-5][0-9])$", out b))
            {
                foreach (string c in b)
                {
                    d = d + c + "+++";

                }
                //textBox2.Text = "true    " + d;

            }
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out b))
            {
                //textBox2.Text = "true    " + b[0];
            }
            //
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out b))
            {
                //textBox2.Text = "true    " + b[0];
            }
            else
            {
                MessageBox.Show("no match");
            }
            //return Regex.IsMatch(StrSource, @"^((20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d)$");
        }

        //设置deadline时间。
        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WfSetting wf = new WfSetting();
            //wf.Show();
            wf.Owner = this;
            wf.ShowDialog(this);

            wf.Close();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
        }

        Member_QingJia MQJ = null;
        private void 外勤ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MQJ = new Member_QingJia();
            MQJ.Owner = this;
            MQJ.ShowDialog(this);
            MQJ.Close();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult messageboxResult = MessageBox.Show("确认退出?", "注意", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (messageboxResult == DialogResult.Yes)
            {
                this.Close();
            }
        }


        private void toolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            string FilePath = "";
            //文件路径
            FilePath = OpenFile();
            Member_Departement_Communications memberSchedule;
            ListMemberSchedule = new List<Member_Departement_Communications>(); //全员列表
            List<string> _checkedMemberName = new List<string>(); //已遍历过得员工的名字列表
            List<string> MemberName = new List<string>();
            WorkTime wt;
            listWorkTime = new List<WorkTime>();//员工每日出勤时间

            int countDay = 1;

            try
            {
                if (FilePath != "" && FilePath != null)
                {
                    DataSet DS = LoadDataFromExcel(FilePath);

                    DataTable DT = DS.Tables[0];

                    List<string> st = new List<string>();

                    int Nrow = 0; int Ncol = 0;
                    foreach (DataRow dr in DT.Rows)
                    {
                        Nrow++;
                        Ncol = 0;
                        foreach (DataColumn dc in DT.Columns)
                        {
                            Ncol++;

                            //单元格不为空时
                            if (dr[dc].ToString() != "" && dr[dc].ToString() != null)
                            {
                                //xxxx-xx-xx-xxxx-xx-xx
                                //记录文件所记录月份.
                                if (fcs.isExMatch(dr[dc].ToString().Replace(" ", ""), @"^((\d{4})-([0,1]?\d)-([0,3]?\d))--(\d{4}-([0,1]?\d)-([0,3]?\d))$", out MemberName))
                                {

                                    start = MemberName[0];
                                    end = MemberName[4];

                                    startY = MemberName[1];
                                    startM = MemberName[2];
                                    StartD = MemberName[3];

                                    endY = MemberName[4];
                                    endM = MemberName[5];
                                    endD = MemberName[6];

                                    foreach (string _str in MemberName)
                                    {
                                        LaDuree.Add(_str);
                                    }

                                    this.Text = "通信所" + start + "至" + end + "考勤记录";
                                }
                                //当数据为2或3位汉字时 记录为姓名
                                if (fcs.isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out MemberName) && MemberName[0] != "通信所" && MemberName[0] != "赵煜")//|| _begin == true)// && MemberName[0] != "通信所" && _appMemberName.Count == 0)
                                {
                                    countDay = 1;
                                    string memberName = "";
                                    memberName = MemberName[0];
                                    //尚未遍历,列表为空
                                    if (_checkedMemberName.Count == 0)
                                    {
                                        _checkedMemberName.Add(memberName);//记录员工名称 
                                    }
                                    else if (_checkedMemberName.Count != 0 && _checkedMemberName[_checkedMemberName.Count - 1] != memberName && !_checkedMemberName.Contains(memberName)) //或者发现列表中尚未出现的员工名称时
                                    {
                                        memberSchedule = new Member_Departement_Communications(_checkedMemberName[_checkedMemberName.Count - 1], listWorkTime);

                                        ListMemberSchedule.Add(memberSchedule);

                                        _checkedMemberName.Add(memberName);//记录员工名称 
                                        listWorkTime = new List<WorkTime>();
                                    }

                                    continue;
                                }
                                //ListMemberSchedule;

                                //读取当单元格数据为 “数字（2位）汉字（一位）” 时读取下一行，同一排单元格的数据
                                List<string> inDate = new List<string>();
                                if (fcs.isExMatch(dr[dc].ToString().Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out inDate))
                                {

                                    string a = inDate[0];
                                    string b = inDate[1];
                                    if (countDay <= Convert.ToInt32(endD))
                                    {
                                        countDay++;

                                        string strColName = dc.ColumnName.ToString();

                                        DataRow seleRow = DT.Rows[Nrow];
                                        string dataInCol = seleRow[dc].ToString();
                                        //string str = dtc.[Nrow + 1];

                                        if (dataInCol.Replace(" ", "") == "-" || dataInCol.Replace(" ", "") == "")
                                        {
                                            wt = new WorkTime(inDate);
                                        }
                                        else
                                        {
                                            bool isWorkingOk;
                                            wt = fcs.ConvertStringToDateTime(dataInCol, inDate, out isWorkingOk);
                                            if (!isWorkingOk)
                                            {
                                                MessageBox.Show("数据读取错误.", "警告", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                            }
                                        }

                                        listWorkTime.Add(wt);
                                        //MessageBox.Show("" + DateZh.ToString() + "---" + strColName + "---" + dataInCol);
                                        //Convert.ToDateTime(StartTime.ToString());
                                        //continue; 
                                    }
                                    else
                                    {

                                    }
                                }

                            }
                        }
                    }

                    dataGridView1.DataSource = DT;
                    //加入表中最后一个成员的信息
                    memberSchedule = new Member_Departement_Communications(_checkedMemberName[_checkedMemberName.Count - 1], listWorkTime);
                    ListMemberSchedule.Add(memberSchedule);

                    MQJ = new Member_QingJia();
                    MQJ.MdiParent = this;
                    MQJ.Show();
                    this.Size = new Size(MQJ.Width + 30, MQJ.Height + 100);

                    //formMember.ShowDialog(this);
                    //formMember.Close();
                    外勤ToolStripMenuItem.Enabled = DataIsSet;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("" + ex);
            }
        }
        output formOutPut = null;
        private void 输出ToolStripMenuItem_Click(object sender, EventArgs e)
        {//MemberChuQingList
            if (MQJ != null)
            {
                //MQJ = new Member_QingJia();

                formOutPut = new output(MQJ.MemberChuQingList);
                formOutPut.Owner = this;
                formOutPut.StartPosition = FormStartPosition.Manual;
                formOutPut.Location = (Point)new Size(this.Location.X + this.Width, this.Location.Y);
                formOutPut.Height = this.Height;
                formOutPut.Show();
            }
            else
            {
                MessageBox.Show("需要先导入考勤文件", "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void Form1_LocationChanged(object sender, EventArgs e)
        {
            if (formOutPut != null)
            {
                formOutPut.Location = (Point)new Size(this.Location.X + this.Width, this.Location.Y);

            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (formOutPut != null)
            {
                Size _size = formOutPut.Size;
                formOutPut.Size = new Size(_size.Width, this.Size.Height);
                formOutPut.Location = (Point)new Size(this.Location.X + this.Width, this.Location.Y);

            }
        }

        /*
        private void button1_Click(object sender, EventArgs e)
        {
            string filePath=null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件.";
            fileDialog.Filter = "Excel97-2003文件|*.xls;*.xlt;*.xltm|Excel2007-2010|*.xlsx|所有文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = fileDialog.FileName;
                //fileType=fileDialog.f
            }
            DataTable dt = ExcelToTableForXLSX(filePath);
            dataGridView1.DataSource = dt;
        }*/
        /*
        #region Excel2003
        /// <summary>  
        /// 将Excel文件中的数据读出到DataTable中(xls)  
        /// </summary>  
        /// <param name="file"></param>  
        /// <returns></returns>  
        public static DataTable ExcelToTableForXLS(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);
                ISheet sheet = hssfworkbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLS(header.GetCell(i) as XLS.HSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;  
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLS(sheet.GetRow(i).GetCell(j) as XLS.HSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>  
        /// 将DataTable数据导出到Excel文件中(xls)  
        /// </summary>  
        /// <param name="dt"></param>  
        /// <param name="file"></param>  
        public static void TableToExcelForXLS(DataTable dt, string file)
        {
            XLS.HSSFWorkbook hssfworkbook = new XLS.HSSFWorkbook();
            ISheet sheet = hssfworkbook.CreateSheet("Test");

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            hssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>  
        /// 获取单元格类型(xls)  
        /// </summary>  
        /// <param name="cell"></param>  
        /// <returns></returns>  
        private static object GetValueTypeForXLS(XLS.HSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.BLANK: //BLANK:  
                    return null;
                case CellType.BOOLEAN: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.NUMERIC: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.STRING: //STRING:  
                    return cell.StringCellValue;
                case CellType.ERROR: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.FORMULA: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
        #endregion  
        */
        /*
        #region Excel2007
        /// <summary>
        /// 将Excel文件中的数据读出到DataTable中(xlsx)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static DataTable ExcelToTableForXLSX(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                ISheet sheet = xssfworkbook.GetSheetAt(0);

                //表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLSX(header.GetCell(i) as XSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLSX(sheet.GetRow(i).GetCell(j) as XSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 将DataTable数据导出到Excel文件中(xlsx)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void TableToExcelForXLSX(DataTable dt, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet("Test");

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            xssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 获取单元格类型(xlsx)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueTypeForXLSX(XSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:
                    return null;
                case CellType.Boolean: //BOOLEAN:
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:
                    return cell.NumericCellValue;
                case CellType.String: //STRING:
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:
                default:
                    return "=" + cell.CellFormula;
            }
        }

        #endregion
        */


    }
}
