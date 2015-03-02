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

using System.Xml;

using System.Text.RegularExpressions;
using System.Data.Odbc;



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
            public TimeSpan _amTime
            {
                get { return _am_Time; }
                set { _am_Time = value; }
            }

            public AMTime(TimeSpan AmTime)
            {
                this._amTime = AmTime;
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




        public class WorkTime
        {
            private AMTime _time_am;
            public AMTime _amTime
            {
                get { return this._time_am; }
                set { this._time_am = value; }
            }

            private PMTime _time_pm;
            public PMTime _pmTime
            {
                get { return this._time_pm; }
                set { this._time_pm = value; }
            }

            public WorkTime()
            {
            }

            public WorkTime(PMTime PmTime)
            {
                this._pmTime = PmTime;
            }
            public WorkTime(AMTime AmTime)
            {
                this._amTime = AmTime;
            }

            public WorkTime(AMTime AmTime, PMTime PmTime)
            {
                this._amTime = AmTime;
                this._pmTime = PmTime;
            }
        }

        public class Member_Departement_Communications
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



        FunctionsCS fcs = new FunctionsCS();

        public List<WorkTime> listWorkTime; //上班时间

        /*
        public List<int> ListNotEmptyCol = new List<int>();
        */
        public  List<Member_Departement_Communications> ListMemberSchedule;//=new List<Member_Departement_Communications>();
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

        #region but1
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
                            _boxList.Items.Add(item.ToString());
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
        #endregion

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
            splitContainer1.IsSplitterFixed = false;// 1.FixedPanel=FixedPanel.Panel1
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
            toolStripStatusLabel1.Text = NowTime.ToString();
            this.Text = "通信所考勤记录";
        }
        #endregion

        //当遇到名字或者 "11 一 "类似的格式时.
        private void button3_Click(object sender, EventArgs e)
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

            try
            {
                if (FilePath != "")
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
                                if(fcs.isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(\d{4}-[0,1]?\d-[0,3]?\d)--(\d{4}-[0,1]?\d-[0,3]?\d)$", out MemberName))
                                {
                                    string a = MemberName[0];
                                    string b = MemberName[1];

                                    this.Text = "通信所" + a + "至" + b + "考勤记录";
                                }
                                //当数据为2或3位汉字时
                                if (fcs.isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out MemberName) && MemberName[0] != "通信所"&& MemberName[0] != "赵煜")//|| _begin == true)// && MemberName[0] != "通信所" && _appMemberName.Count == 0)
                                {
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
                                    DateTime currentDate = NowTime;
                                    int lastMoth = Convert.ToInt32(currentDate.Month) - 1;
                                    int currentyear = Convert.ToInt32(currentDate.Year);
                                    /*
                                    //将数据转换为 年/月/日/星期 格式
                                    StringBuilder DateZh = new StringBuilder();
                                    DateZh.Append(currentyear.ToString() + "-" + lastMoth.ToString() + "-" + inDate[0] + "-星期" + inDate[1]);

                                    StringBuilder StartTime = new StringBuilder();
                                    StartTime.Append(currentyear.ToString() + "," + lastMoth.ToString() + "," + inDate[0]);

                                    StringBuilder StopTime = new StringBuilder();
                                    StopTime.Append(currentyear.ToString() + "," + lastMoth.ToString() + "," + inDate[0]);
                                    
                                    dr[dc] = DateZh.ToString();
                                    */

                                    string strColName = dc.ColumnName.ToString();
                                    
                                    DataRow seleRow = DT.Rows[Nrow];
                                    string dataInCol = seleRow[dc].ToString();
                                    //string str = dtc.[Nrow + 1];
                                    
                                    if (dataInCol.Replace(" ", "") == "-" || dataInCol.Replace(" ", "") == "")
                                    {
                                        AMTime _errorA = new AMTime(Convert.ToDateTime("0:00:00").TimeOfDay);
                                        PMTime _errorP = new PMTime(Convert.ToDateTime("0:00:00").TimeOfDay);
                                        wt = new WorkTime();
                                    }
                                    else
                                    {
                                        wt = fcs.ConvertStringToDateTime(dataInCol);
                                    }

                                    listWorkTime.Add(wt);
                                    //MessageBox.Show("" + DateZh.ToString() + "---" + strColName + "---" + dataInCol);
                                    //Convert.ToDateTime(StartTime.ToString());
                                    //continue;
                                }
                            }
                        }

                        #region 111
                        /*foreach (DataRow dr in DT.Rows)
                        {
                            Nrow++;
                            Ncol = 0;
                            foreach (DataColumn dc in DT.Columns)
                            {
                                Ncol++;
                                List<string> MathGroup = new List<string>();
                                string a = dr[dc].ToString().Replace(" ", "");

                                if (dr[dc].ToString() == "" || dr[dc].ToString() == null)
                                {
                                    continue;
                                }
                                bool _begin = false;
                                //尚未便利过任何员工时 和 检测到的员工名称与列表中的最后一个不同
                                if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out MemberName) || _begin == true)// && MemberName[0] != "通信所" && _appMemberName.Count == 0)
                                //|| (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out MemberName) && MemberName[0] != "通信所" && _appMemberName[_appMemberName.Count - 1] != MemberName[0]))
                                {
                                    _begin = true;
                                    if (MemberName.Count != 0 && MemberName[0] != "通信所")// &&  (_appMemberName.Count == 0||_appMemberName[_appMemberName.Count - 1] != MemberName[0]))
                                    {
                                        //MessageBox.Show("" + MemberName[0] + "__" + Nrow.ToString() + "__" + Ncol.ToString());

                                        if (_checkedMemberName.Count == 0)//|| _checkedMemberName[_checkedMemberName.Count - 1].name != MemberName[0])
                                        {
                                            listWorkTime = new List<WorkTime>();
                                            _checkedMemberName.Add(MemberName[0]);
                                            Member_Communications _memberSchedule = new Member_Communications(MemberName[0]);
                                            MemberSchedules.Add(_memberSchedule);
                                            continue;
                                        }
                                        else if (MemberSchedules.Count != 0 && MemberSchedules[MemberSchedules.Count - 1].name.ToString() != MemberName[0])
                                        {
                                            //_checkedMemberName[_checkedMemberName.Count - 1].workTime=lis
                                            Member_Communications _memberSchedule = new Member_Communications(MemberName[0]);
                                            MemberSchedules[MemberSchedules.Count - 1].workTime = listWorkTime;
                                            MemberSchedules.Add(_memberSchedule);
                                            //listWorkTime.Clear();
                                            listWorkTime = new List<WorkTime>();
                                            continue;
                                        }
                                    }
                                }
                                //实例化MemberSchedule;


                                #region usefull

                                //当时间格式为xx:xx-yy:yy时
                                if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))
                                        //|| isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d$", out MathGroup)
                                        //|| isExMatch(dr[dc].ToString().Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out MathGroup)
                                        //|| isExMatch(dr[dc].ToString().Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))

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
                                        MessageBox.Show("1:" + MathGroup[0]);
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
                                //时间格式为-xx:xx
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

                                else if (isExMatch(textBox1.Text.Replace(" ", ""), @"^([0-3]\d)一|二|三|四|五|六|日$", out MathGroup))
                                        //else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out MathGroup))
                                {
                                    dr[dc] = "星期" + MathGroup[1];
                                }
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
                                        //else if (isExMatch(dr[dc].ToString().Replace(" ", ""), @"(^[\u4e00-\u9fa5]{3})$", out MathGroup))
                                        //{
                                        //    dr[dc] = dr[dc] + "-" + MathGroup[0];
                                        //}                                
                                else
                                {
                                    continue;
                                }
                                #endregion
                                    //}
                                //}

                            }
                        }*/
                        #endregion

                        dataGridView1.DataSource = DT;

                        //
                    }
                    memberSchedule = new Member_Departement_Communications(_checkedMemberName[_checkedMemberName.Count - 1], listWorkTime);
                    ListMemberSchedule.Add(memberSchedule);
                    WorkingPassion(ListMemberSchedule);


                }
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

        #region 处理xml
        public void LoadWorkTime()
        {
            try
            {
                XmlDocument xmldoc = new XmlDocument();
                try
                {
                    xmldoc.Load(@"..\..\DataFile\WorkTime.xml");    //读取指定的XML文档                    
                }
                catch (Exception)
                {
                    CreateXml(@"..\..\DataFile\WorkTime.xml");//如果程序没有找到xml文件，则新建一个
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
                                xmldoc.Save(@"..\..\DataFile\WorkTime.xml");
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
                                xmldoc.Save(@"..\..\DataFile\WorkTime.xml");
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
        public void CreateXml(string path)
        {
            XmlDocument xmldoc = new XmlDocument();
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

        public void XmlMemberList()
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
                    CreateXml(@"..\..\DataFile\ListMemberName.xml");//如果程序没有找到xml文件，则新建一个
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
        }
        #endregion

        /*
        //判定正则表达式，返回值由正则表达式确定
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
        */

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
            
            string StatueOfEmployer = null;

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
                    if (wt._amTime != null && wt._pmTime != null)
                    {
                        //将数据转换为timeSpan格式
                        

                        DateTime _timeLim = this.SetLimShowUpTime;
                        TimeSpan ts = _timeLim.TimeOfDay;

                        if (CompareTime(wt._amTime._amTime, ts))
                        {
                            _late++;
                        }
                        else
                        {
                            _inTime++;
                        }
                    }
                    else if (wt._amTime == null)
                    {
                        _question++;
                    }
                    else if (wt._pmTime == null)
                    {
                        _noSignOff++;
                    }
                }
                string _statueOfEmployer = EmployerName + System.Environment.NewLine + "\t准时到达:"
                    + _inTime + "次." + System.Environment.NewLine + " \t迟到:" + _late + "次." + System.Environment.NewLine + " \t无早上数据:"
                    + _question + "次. " + System.Environment.NewLine + "\t无下午数据:" + _noSignOff + "次. " + System.Environment.NewLine + System.Environment.NewLine;
                StatueOfEmployer = StatueOfEmployer + _statueOfEmployer;
            }
            textBox2.Text = StatueOfEmployer;
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
            Calender cal = new Calender();
            cal.Show();
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
                || fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out b)
                || fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out b))
            {
                foreach (string c in b)
                {
                    d = d + c + "+++";
                }

                //CompareTime(Convert.ToDateTime(b[0]), Convert.ToDateTime(b[1]));
                //textBox2.Text = "true    " + d;
            }
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d$", out b))
            {
                textBox2.Text = "true    " + b[0];

            }
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"^([0-3]\d)(一|二|三|四|五|六|日)$", out b))
            {
                textBox2.Text = "true    " + b[0];
            }
            //
            else if (fcs.isExMatch(textBox1.Text.Replace(" ", ""), @"(^[\u4e00-\u9fa5]{2,3})$", out b))
            {
                textBox2.Text = "true    " + b[0];
            }
            else
            {
                MessageBox.Show("");
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
            textBox1.Text = this.SetLimShowUpTime.ToString();
        }

        private void 外勤ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Member_QingJia MQJ = new Member_QingJia();
            MQJ.Owner = this;
            MQJ.ShowDialog(this);
            MQJ.Close();
        }
    }
}