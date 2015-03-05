using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C_Excel
{
    public partial class Member_QingJia : Form
    {
        List<string> ListOfMemberName = new List<string>();


        public class Leave_Reason_Date
        {
            private string day;
            private string reason;

            public string date
            {
                get { return day; }
                set { this.day = value; }
            }

            public string leaveReason
            {
                get { return reason; }
                set { this.reason = value; }
            }

            public Leave_Reason_Date(string DateInMonth, string ReasonLeave)
            {
                this.date = DateInMonth;
                this.leaveReason = ReasonLeave;
            }
        }

        public class Member_Leave
        {
            private string _name;
            public string workerName
            {
                get { return _name; }
                set { this._name = value; }
            }

            private Leave_Reason_Date _memberLeave;
            public Leave_Reason_Date memberLeave
            {
                get { return _memberLeave; }
                set { this._memberLeave = value; }
            }

            public Member_Leave(string NameOfMember, Leave_Reason_Date LRD)
            {
                this.workerName = NameOfMember;
                this.memberLeave = LRD;
            }

        }

        public class MemberChuQingState
        {
            private string name;
            private int isLate;
            private int onTime;
            private int inQuestion;
            private int notSignOff;
            private int badData;
            public string workerName
            {
                get { return name; }
                set { this.name = value; }
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
            public int BadData
            {
                get { return badData; }
                set { badData = value; }
            }
            public MemberChuQingState(string WorkerName,int WorkerLate,int WorkerOnTime,int DateQuestion,int WorkerDidntSignOff,int numBadData)
            {
                this.name = WorkerName;
                this.workerIsLate = WorkerLate;
                this.workerOnTime = WorkerOnTime;
                this.dataInQuestion = DateQuestion;
                this.workerNotSignOff = WorkerDidntSignOff;
                this.BadData = numBadData;
            }
        }

        private bool _startProcecs = false;
        /*************/
        bool _ChuChai = false;
        bool _ShiJia = false;
        bool _Vacance = false;
        /*************/

        public List<Control> listComponant = new List<Control>();
        public List<string> listPassedMember = new List<string>(); //以遍历成员名称
        public List<MemberChuQingState> MemberChuQingList = new List<MemberChuQingState>();

        List<Member_Leave> Member_NotShowUp = new List<Member_Leave>();
        int _dayInWeek = 0;
        public Member_QingJia()
        {
            InitializeComponent();            
        }


        private void B_Valide_Click(object sender, EventArgs e)
        {
            int nonVisited = 0;
            StringBuilder str = new StringBuilder();
            if (listPassedMember.Count() == ((Form1)this.Owner).ListMemberSchedule.Count())
            {
                //LoadMemberLeaveList();
                WorkingPassion(((Form1)this.Owner).ListMemberSchedule);
            }
            else
            {
                //显示没有设置出差/事假人员名单.
                foreach (Form1.Member_Departement_Communications lmdc in ((Form1)this.Owner).ListMemberSchedule)
                {
                    if (!listPassedMember.Contains(lmdc.name))
                    {
                        nonVisited++;
                        str.Append("\t" + lmdc.name + System.Environment.NewLine);
                    }
                }
                MessageBox.Show("还有" + nonVisited + "名员工休假状况未定义"+System.Environment.NewLine+"分别是:"+System.Environment.NewLine+
                str.ToString(), "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void B_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Member_QingJia_Load(object sender, EventArgs e)
        {
            int b = ((Form1)this.Owner).ListMemberSchedule.Count();
            toolStripProgressBar1.Maximum = b;
            toolStripProgressBar1.Style = ProgressBarStyle.Blocks;

            toolState.Text = "";
            toolinfor.Text = "";
            toolTimer.Text = "";
            toolStripStatusLabel4.Text = "";
            toolStripStatusLabel5.Text = "";
            this.timer1.Enabled = true;
            timer1.Start();
            if (((Form1)this.Owner).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.Owner).LaDuree;
                string _str = a[0] + " -- " + a[4];

                //显示label
                DisplayTime(_str);
                
                
                groupBox1.Text = a[2] + "月月历";
                foreach (Form1.Member_Departement_Communications item in ((Form1)this.Owner).ListMemberSchedule)
                {
                    ListOfMemberName.Add(item.name);
                    comboxMember.Items.Add(item.name);
                }
                fileTheCalendar();

                //MessageBox.Show(a);
            }
            else
            {
                MessageBox.Show("需要先倒入Excel文件.", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                this.Close();
            }
            
        }

        private bool _eventAdded = false;
        private void comboxMember_SelectionChangeCommitted(object sender, EventArgs e)
        {

            if (!listPassedMember.Contains(this.comboxMember.SelectedItem.ToString()) || listPassedMember.Count == 0)
            {
                _startProcecs = true;
                //刷新按钮颜色.
                fileTheCalendar();
                _ChuChai = false;
                _ShiJia = false;
                _Vacance = false;
                B_ChuChai.FlatStyle = FlatStyle.Standard;
                B_ShiJia.FlatStyle = FlatStyle.Standard;
                B_Vacance.FlatStyle = FlatStyle.Standard;

                B_ChuChai.BackColor = Color.Yellow;
                B_ShiJia.BackColor = Color.Aqua;
                B_Vacance.BackColor = Color.Orange;

                string _name = ((Form1)this.Owner).ListMemberSchedule[comboxMember.SelectedIndex].name;
                listPassedMember.Add(_name);
                List<Form1.WorkTime> _lWorkTime = ((Form1)this.Owner).ListMemberSchedule[comboxMember.SelectedIndex].workTime;
                //MessageBox.Show(_name);
                if (_startProcecs!=true)
                {
                toolinfor.Text = "可以通过点击上方按钮来为每个员工定义出差或休假日期.";

                }
                toolState.Text = "选中:" + _name;
                if (_name != "")
                {
                    List<Form1.WorkTime> LFWT = ((Form1)this.Owner).ListMemberSchedule[comboxMember.SelectedIndex].workTime;
                    foreach (Control c in this.panel1.Controls)
                    {

                        if (c is Button)
                        {
                            if (listComponant.Contains(c))//当按钮表示所需月份日期时
                            {
                                c.Enabled = true;
                                //c.BackColor = Color.Control;
                                #region 将按钮文字改为日期和考勤时间 并同时改变按钮颜色
                                foreach (Form1.WorkTime FWT in LFWT)
                                {
                                    bool enretard = false;
                                    bool alheure = false;
                                    bool parPresente = false;
                                    if (FWT._Date.Count != 0)
                                    {
                                        List<string> text = FWT._Date;
                                        if (c.Text != "" && c.Text != null)
                                        {
                                            string[] yearMonthDay = c.Text.Split(new[] { " " }, StringSplitOptions.None); //将按钮名字分为两个部分,由"空格"分割
                                            string[] Day = yearMonthDay[0].Split(new[] { "-" }, StringSplitOptions.None); //将按钮名字分为three个部分,由"-"分割
                                            if (Convert.ToInt32(text[0]) == Convert.ToInt32(Day[2]))// || Convert.ToInt32(text[1]) == Convert.ToInt32(Day[2]))
                                            {
                                                string buttonText = c.Text;
                                                DateTime defautTime = Convert.ToDateTime(yearMonthDay[0]);
                                                DateTime morningTime;
                                                DateTime afterTime;
                                                string morningText = "";
                                                string afterText = "";

                                                Form1 fom = new Form1();
                                                DateTime LimitMorningTime = fom.SetLimShowUpTime;
                                                DateTime AfternoonTime = fom.SetLimDissmisTime;

                                                if ((FWT.amTime == null) && (FWT.pmTime == null))
                                                {
                                                    parPresente = true;
                                                    c.BackColor = Color.Violet;
                                                }
                                                else
                                                {
                                                    if (FWT.amTime == null)
                                                    {
                                                        c.BackColor = Color.Pink;
                                                    }
                                                    else
                                                    {
                                                        morningTime = Convert.ToDateTime(DateTime.Now.ToShortDateString()) + FWT.amTime.amTime;
                                                        morningText = FWT.amTime.amTime.ToString();
                                                        string text1 = morningTime.ToShortTimeString();
                                                        string text2 = LimitMorningTime.ToShortTimeString();
                                                        //if (TimeSpan.TryParse(text1, out interval))
                                                        if (DateTime.Compare(morningTime, LimitMorningTime) == 1)
                                                        {
                                                            //第一个时间比第二个时间大
                                                            c.BackColor = Color.Red;
                                                        }
                                                        else
                                                        {
                                                            //第一个时间比第二个时间小
                                                            c.BackColor = Color.Lime;
                                                        }
                                                    }
                                                    if (FWT.pmTime == null)
                                                    {
                                                        c.BackColor = Color.Pink;
                                                    }
                                                    else
                                                    {
                                                        afterTime = defautTime + FWT.pmTime.pmTime;
                                                        afterText = FWT.pmTime.pmTime.ToString();
                                                    }
                                                }
                                                c.Text = buttonText + System.Environment.NewLine + morningText + "-" + afterText.ToString();


                                                break;
                                            }
                                            
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                #endregion 
                                if (_eventAdded != true)
                                {
                                    c.MouseClick += c_MouseClick;
                                }
                            }
                        }
                        
                    }
                    _eventAdded = true;
                }
                else
                {

                }
            }
            int a = listPassedMember.Count();
            toolStripProgressBar1.Value = a;
        }

        void c_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                string text = ((Control)sender).Text;
                int lenText = text.Count();
                toolinfor.Text = "设置每名员工请假或出差的时间." ;
                string[] partText = text.Split(new[] { "\r\n" }, StringSplitOptions.None); //将按钮名字分为两个部分,由"空格"分割
                string toolInformation = partText[0];// +" 星期" + partText[1] + " ";

                
                if (_ShiJia == true)
                {
                    if (((Control)sender).BackColor != Color.Aqua)
                    {
                        ((Control)sender).BackColor = Color.Aqua;
                        toolState.Text = toolInformation + comboxMember.SelectedItem.ToString() + "申请事假";

                        Leave_Reason_Date _reasonAndDate=new Leave_Reason_Date(partText[0],"事假");
                        Member_Leave _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                    else
                    {
                        RemoveFromList(comboxMember.SelectedItem.ToString(), partText, "事假");//取消选定,并从列表中删除

                        ((Control)sender).BackColor = this.B_Valide.BackColor;

                    }
                }
                else if (_ChuChai == true)
                {
                    if (((Control)sender).BackColor != Color.Yellow)
                    {
                        ((Control)sender).BackColor = Color.Yellow;
                        toolState.Text = toolInformation + comboxMember.SelectedItem.ToString() + "申请出差";

                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "出差");
                        Member_Leave _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                    else
                    {
                        RemoveFromList(comboxMember.SelectedItem.ToString(), partText, "事假"); //取消选定,并从列表中删除

                        ((Control)sender).BackColor = this.B_Valide.BackColor;
                    }
                }
                else if (_Vacance == true)
                {
                    if (((Control)sender).BackColor != Color.Orange)
                    {
                        ((Control)sender).BackColor = Color.Orange;
                        toolState.Text = toolInformation + comboxMember.SelectedItem.ToString() + "申请放假";

                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "放假");
                        Member_Leave _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                    else
                    {
                        RemoveFromList(comboxMember.SelectedItem.ToString(), partText, "事假");//取消选定,并从列表中删除

                        ((Control)sender).BackColor = this.B_Valide.BackColor;
                    }
                }
                
            }
        }

        public void RemoveFromList(string name,string[] Date,string reason)
        {
            Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(Date[0], reason);
            Member_Leave _memberChuQing = new Member_Leave(name, _reasonAndDate);
            int index = 0;
            foreach (Member_Leave M_L in Member_NotShowUp)
            {
                if (M_L.workerName == name && M_L.memberLeave.date == Date[0] && M_L.memberLeave.leaveReason == reason)
                {
                    Member_NotShowUp.RemoveAt(index);
                    break;
                }
                index++;
            }
            
            string toolInformation = Date[0] + " 星期" + Date[1] + " ";
            toolState.Text = "取消 " + toolInformation + " " + comboxMember.SelectedItem.ToString() + "申请事假"; 
        }
        private void Member_QingJia_Resize(object sender, EventArgs e)
        {

            if (((Form1)this.Owner).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.Owner).LaDuree;
                string _str = a[0] + " -- " + a[3];
                DisplayTime(_str);

            }
        }

        #region change methode, not good
        public void LoadMemberLeaveList()
        {

            foreach (Control c in this.panel1.Controls)
            {

                if (c is Button)
                {
                    string text = c.Text;
                    string[] partText = text.Split(new[] { "\r\n" }, StringSplitOptions.None); //将按钮名字分为两个部分,由"空格"分割

                    Member_Leave _memberChuQing;

                    if (c.BackColor == Color.Aqua)
                    {
                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "事假");
                        _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                    else if (c.BackColor == Color.Yellow)
                    {
                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "出差");
                        _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                    else if (c.BackColor == Color.Orange)
                    {
                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "放假");
                        _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        Member_NotShowUp.Add(_memberChuQing);
                    }
                }
            }
        }
        #endregion

        public void WorkingPassion(List<Form1.Member_Departement_Communications> MCs)
        {
            Form1 form1 = new Form1();
            DateTime DissmisTime = form1.SetLimDissmisTime;
            DateTime LimShowUpTime = form1.SetLimShowUpTime;

            foreach (Form1.Member_Departement_Communications mc in MCs)
            {
                string EmployerName = mc.name;
                List<Form1.WorkTime> _listWorkTime = new List<Form1.WorkTime>();
                _listWorkTime = mc.workTime;
                int _inTime = 0;
                int _noSignOff = 0;
                int _late = 0;
                int _question = 0;
                int _noData = 0;
                foreach (Form1.WorkTime wt in _listWorkTime)
                {
                    //复制list<>
                    //var _tempChuQing = Member_NotShowUp.ToList(); 对新/老 list的更改都会对另一个进行变更.
                    List<Member_Leave> _tempChuQing = new List<Member_Leave>(Member_NotShowUp.Count);   //新建一个成员出差日期表
                    Member_NotShowUp.ForEach((item) =>
                    {
                        _tempChuQing.Add(new Member_Leave(item.workerName, item.memberLeave));
                    });

                    if (wt._Date != null)//只计算本月的数据
                    {
                        //InTheList(List<Member_Leave> LML,string memberName, string textAnalys)
                        string workDay = ((Form1)this.Owner).LaDuree[1] + "-" + ((Form1)this.Owner).LaDuree[2] + "-" + wt._Date[0]; //从excel文件中读取的日期

                        if (!InTheList(_tempChuQing, EmployerName, workDay))
                        {
                            if (wt.amTime != null && wt.pmTime != null)
                            {
                                //将数据转换为timeSpan格式
                                DateTime _timeLim = LimShowUpTime;
                                TimeSpan limitTimeSpan = _timeLim.TimeOfDay;

                                if (form1.CompareTime(wt.amTime.amTime, limitTimeSpan))
                                {
                                    _late++;
                                }
                                else
                                {
                                    _inTime++;
                                }
                            }
                            else if (wt.amTime == null && wt.pmTime == null)
                            {
                                _noData++;
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
                        else
                        {

                        }


                    }
                }
                MemberChuQingState MCQS = new MemberChuQingState(EmployerName, _late, _inTime, _question, _noSignOff, _noData);
                MemberChuQingList.Add(MCQS);
            }
        }

        private bool InTheList(List<Member_Leave> LML,string memberName, string textAnalys)
        {
            bool inTheList = false;
            foreach (Member_Leave item in LML) //全员出差表
            {
                if (textAnalys == item.memberLeave.date && item.workerName == memberName)
                {
                    inTheList = true;
                    break;
                }
            }
            return inTheList;
        }

        private static SizeF TextSize(string text, Font txtFnt)
        {
            SizeF txtSize = new SizeF();
            // The size returned is 'Size(int width, int height)' where width and height
            // are the dimensions of the string in pixels
            Size s = System.Windows.Forms.TextRenderer.MeasureText(text, txtFnt);
            // Value based on normal DPI settings of 96
            txtSize.Width = (float)Math.Ceiling((float)s.Width / 96f * 100f);
            txtSize.Height = (float)Math.Ceiling((float)s.Height / 96f * 100f);
            return txtSize;
        }

        public void DisplayTime(string str)
        {
            label8.Text = str;
            Font textFont = new Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point);
            SizeF _size = TextSize(str, textFont);

            Point newP = new Point(Convert.ToInt32(this.Size.Width - _size.Width) / 2, 9);
            label8.Location = newP;
            groupBox2.Location = new Point(Convert.ToInt32(this.Size.Width - groupBox2.Width) / 2, 25);
            B_Valide.Location = new Point(this.Width / 2 - 101, 384);
            B_Cancel.Location = new Point(this.Width / 2 + 44, 384);
        }

        DateTime[] listDT;

        public void fileTheCalendar()
        {
            List<string> a = ((Form1)this.Owner).LaDuree;
            int numberButton = 0;   //以遍历的按钮的个数
            int numberDayOfWeek = 1;    //星期
            int numberDayInMonth = 1;
            string textFirstDayOfMonth = a[0];
            DateTime FirstDayOfMonth = Convert.ToDateTime(textFirstDayOfMonth);
            int FirstDayOfMonthInWeek = Convert.ToInt32(FirstDayOfMonth.DayOfWeek); //文件所记录月份的第一天是星期几

            string[] DayOfWeek = { "", "一", "二", "三", "四", "五", "六", "日" }; //DayOfWeek的取值范围是1-7
            //为调整按钮名称
            bool _pass = false;
            foreach (Control c in this.panel1.Controls)
            {

                if (c is Button)
                {
                    c.BackColor = B_Valide.BackColor; //按钮的默认颜色随系统设置而变动.
                    c.Enabled = false;
                    toolinfor.Text = "需要选择员工名称才能进行下一步操作";
                    if ((_pass == true || numberDayOfWeek == FirstDayOfMonthInWeek) && (numberDayInMonth <= Convert.ToInt32(a[6]))) //从文件记录月份的第一天是星期几开始记录
                    {
                        _pass = true;
                        //string 

                        FunctionsCS fcs = new FunctionsCS();
                        if (numberButton <= Convert.ToInt32(a[6]))
                        {
                            string nameButton = a[1] + "-" + a[2] + "-" + (numberButton + 1); //按钮名称格式为 年-月- 日
                            c.Text = nameButton;

                        }
                        numberButton++;
                        numberDayInMonth++;
                        listComponant.Add(c);
                    }
                    else
                    {
                        c.Text = "";
                    }
                    if (numberDayOfWeek == 7) //星期日
                    {
                        numberDayOfWeek = 0;
                        c.BackColor = Color.Orange;
                    }
                    if (numberDayOfWeek == 6) //星期六
                    {
                        c.BackColor = Color.Orange;

                    }
                    numberDayOfWeek++;//dayofweek递增
                }
                //


            }


        }

        private void B_ChuChai_Click(object sender, EventArgs e)
        {
            

            if (B_ChuChai.FlatStyle != FlatStyle.Flat)
            {
                _ChuChai = true;
                B_ChuChai.FlatStyle = FlatStyle.Flat; //been selected
                B_ChuChai.BackColor = Color.Olive;

                //如果 事假 按钮被选中,那么恢复事假按钮为未选中,并改变_ShiJia的值
                if (B_ShiJia.FlatStyle == FlatStyle.Flat)
                {
                    B_ShiJia.FlatStyle = FlatStyle.Standard;
                    B_ShiJia.BackColor = Color.Aqua;
                    _ShiJia = false;
                }
                if (B_Vacance.FlatStyle == FlatStyle.Flat)
                {
                    B_Vacance.FlatStyle = FlatStyle.Standard;
                    B_Vacance.BackColor = Color.Orange;
                    _Vacance = false;
                }
            }
            else
            {
                B_ChuChai.FlatStyle = FlatStyle.Standard;
                B_ChuChai.BackColor = Color.Yellow;
                _ChuChai = false;
            }
        }

        private void B_ShiJia_Click(object sender, EventArgs e)
        {
            

            if (B_ShiJia.FlatStyle != FlatStyle.Flat)
            {
                _ShiJia = true;
                B_ShiJia.FlatStyle = FlatStyle.Flat;//been selected
                B_ShiJia.BackColor = Color.Teal;
                //如果 出差 按钮被选中,那么恢复事假按钮为未选中,并改变_ShiJia的值
                if (B_ShiJia.FlatStyle == FlatStyle.Flat)
                {
                    B_ChuChai.FlatStyle = FlatStyle.Standard;
                    B_ChuChai.BackColor = Color.Yellow;
                    _ChuChai = false;
                }
                if (B_Vacance.FlatStyle == FlatStyle.Flat)
                {
                    B_Vacance.FlatStyle = FlatStyle.Standard;
                    B_Vacance.BackColor = Color.Orange;
                    _Vacance = false;
                }
            }
            else
            {
                B_ShiJia.FlatStyle = FlatStyle.Standard;
                B_ShiJia.BackColor = Color.Aqua;
                _ShiJia = false;
            }
        }

        private void B_Vacance_Click(object sender, EventArgs e)
        {
            if (B_Vacance.FlatStyle != FlatStyle.Flat)
            {
                _Vacance = true;
                B_Vacance.FlatStyle = FlatStyle.Flat;//been selected
                B_Vacance.BackColor = Color.DarkGoldenrod;
                //如果 出差 按钮被选中,那么恢复事假按钮为未选中,并改变_ShiJia的值
                if (B_ChuChai.FlatStyle == FlatStyle.Flat)
                {
                    B_ChuChai.FlatStyle = FlatStyle.Standard;
                    B_ChuChai.BackColor = Color.Yellow;
                    _ChuChai = false;
                }
                if (B_ShiJia.FlatStyle == FlatStyle.Flat)
                {
                    B_ShiJia.FlatStyle = FlatStyle.Standard;
                    B_ShiJia.BackColor = Color.Aqua;
                    _ShiJia = false;
                }
            }
            else
            {
                B_Vacance.FlatStyle = FlatStyle.Standard;
                B_Vacance.BackColor = Color.Orange;
                _Vacance = false;
            }
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolTimer.Text = DateTime.Now.ToString();
        }

        private void comboxMember_SelectedIndexChanged(object sender, EventArgs e)
        {
        }



    } 
}
