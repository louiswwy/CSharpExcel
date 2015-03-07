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

        public class MemberChuQingStatistics
        {
            private string name;
            private int isLate = 0;
            private int onTime = 0;
            private int inQuestion = 0;
            private int notSignOff = 0;
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
            public MemberChuQingStatistics(string WorkerName,int WorkerLate,int WorkerOnTime,int DateQuestion,int WorkerDidntSignOff,int numBadData)
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
        public List<MemberChuQingStatistics> MemberChuQingList = new List<MemberChuQingStatistics>();

        Form1 fom = new Form1();
        DateTime LimitMorningTime;  // = fom.SetLimShowUpTime;
        DateTime AfternoonTime; // = fom.SetLimDissmisTime;

        List<Member_Leave> Member_NotShowUp = new List<Member_Leave>();
        int _dayInWeek = 0;
        public Member_QingJia()
        {
            InitializeComponent();            
        }

        public delegate void MyDelegate();
        //定义事件
        public event MyDelegate PantResultatEvent;


        private void B_Valide_Click(object sender, EventArgs e)
        {
            int nonVisited = 0;

            StringBuilder str = new StringBuilder();
            if (listPassedMember.Count() == ((Form1)this.MdiParent).ListMemberSchedule.Count())
            {
                fom.DataIsSet = true;
                output formOutPut = new output(MemberChuQingList);
                formOutPut.Owner = this;
                formOutPut.Show();
                bool a=formOutPut.IsDisposed;
                //this.Close();
            }
            else
            {
                //显示没有设置出差/事假人员名单.
                foreach (Form1.Member_Departement_Communications lmdc in ((Form1)this.MdiParent).ListMemberSchedule)
                {
                    if (!listPassedMember.Contains(lmdc.name))
                    {
                        nonVisited++;
                        str.Append("\t" + lmdc.name + System.Environment.NewLine);
                    }
                }
                MessageBox.Show("还有" + nonVisited + "名员工休假状况未定义" + System.Environment.NewLine + "分别是:" + System.Environment.NewLine +
                str.ToString(), "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void B_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Member_QingJia_Load(object sender, EventArgs e)
        {
            int b = ((Form1)this.MdiParent).ListMemberSchedule.Count();
            toolStripProgressBar1.Maximum = b;
            toolStripProgressBar1.Style = ProgressBarStyle.Blocks;

            toolStripStatusLabel1.Text = "";
            toolState.Text = "";
            toolinfor.Text = "";
            toolTimer.Text = "";

            分割线1.Text = "";
            分割线2.Text = "";
            分割线3.Text = "";
            分割线4.Text = "";

            this.timer1.Enabled = true;
            timer1.Start();
            if (((Form1)this.MdiParent).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.MdiParent).LaDuree;
                string _str = a[0] + " -- " + a[4] + " ";

                //显示label
                string defautTitle = this.Text;
                this.Text = _str + defautTitle;

                groupBox1.Text = a[2] + "月月历";
                foreach (Form1.Member_Departement_Communications item in ((Form1)this.MdiParent).ListMemberSchedule)
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

            if (comboxMember.SelectedItem != null && comboxMember.SelectedItem != "")
            {
                //if (/*!listPassedMember.Contains(this.comboxMember.SelectedItem.ToString()) ||*/ listPassedMember.Count == 0)
                //{
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

                string _name = ((Form1)this.MdiParent).ListMemberSchedule[comboxMember.SelectedIndex].name;

                if (!listPassedMember.Contains(_name))
                {
                    listPassedMember.Add(_name);
                }

                //List<Form1.WorkTime> _lWorkTime = ((Form1)this.MdiParent).ListMemberSchedule[comboxMember.SelectedIndex].workTime;


                toolinfor.Text = "可以通过点击上方按钮来为每个员工定义出差或休假日期.";

                toolStripStatusLabel1.Text = "以设置:" + listPassedMember.Count + "/" + ((Form1)this.MdiParent).ListMemberSchedule.Count + "成员";
                toolState.Text = "选中:" + _name;
                if (_name != "")
                {
                    PaintAndCalcule(_name);

                    #region 1111
                    /*
                List<Form1.WorkTime> LFWT = ((Form1)this.MdiParent).ListMemberSchedule[comboxMember.SelectedIndex].workTime;

                //定义
                MemberChuQingStatistics memberState;
                int _inTime = 0;    //准时
                int _noSignOff = 0; //没有打卡下班
                int _late = 0;  //迟到
                int _question = 0;  //数据有问题
                int _noData = 0;    //没有数据

                int passedbutton = 0;
                foreach (Control c in this.panel1.Controls)
                {

                    if (c is Button)
                    {
                        if (listComponant.Contains(c))//当按钮表示所需月份日期时
                        {
                            c.Enabled = true;
                            //c.BackColor = Color.Control;

                            #region 将按钮文字改为日期和考勤时间 并同时改变按钮颜色 同时记录次数

                            //foreach (Form1.WorkTime FWT in LFWT)
                            Form1.WorkTime FWT = LFWT[passedbutton];
                            passedbutton++;
                            //{

                            bool enretard = false;
                            bool alheure = false;
                            bool pasPresente = false;
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

                                        //Form1 fom = new Form1();
                                        LimitMorningTime = fom.SetLimShowUpTime;
                                        AfternoonTime = fom.SetLimDissmisTime;
                                        string aa = Convert.ToDateTime(yearMonthDay[0]).DayOfWeek.ToString();
                                        if (Convert.ToDateTime(yearMonthDay[0]).DayOfWeek.ToString() != "Saturday"
                                            && Convert.ToDateTime(yearMonthDay[0]).DayOfWeek.ToString() != "Sunday")
                                        {
                                            if ((FWT.amTime == null) && (FWT.pmTime == null)) //没有记录数据
                                            {
                                                pasPresente = true;
                                                c.BackColor = Color.Violet;
                                                _noData++;
                                            }
                                            else
                                            {
                                                if (FWT.amTime == null) //缺失上午的数据
                                                {
                                                    c.BackColor = Color.Pink;
                                                    _question++;
                                                }
                                                else if (FWT.amTime != null && FWT.pmTime != null)  //有上午的时间时
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
                                                        _late++;
                                                    }
                                                    else
                                                    {
                                                        //第一个时间比第二个时间小
                                                        c.BackColor = Color.Lime;
                                                        _inTime++;
                                                    }
                                                }
                                                else if (FWT.amTime != null && FWT.pmTime == null)
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
                                                        _late++;
                                                    }
                                                    else
                                                    {
                                                        //第一个时间比第二个时间小
                                                        c.BackColor = Color.Lime;
                                                        _inTime++;
                                                    }
                                                }
                                                if (FWT.pmTime == null) //缺失下午的数据
                                                {
                                                    //c.BackColor = Color.Pink;
                                                    _noSignOff++;
                                                }
                                                else
                                                {
                                                    afterTime = defautTime + FWT.pmTime.pmTime;
                                                    afterText = FWT.pmTime.pmTime.ToString();
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (FWT.amTime != null)
                                            {
                                                morningText = FWT.amTime.amTime.ToString();
                                            }
                                            if (FWT.pmTime != null)
                                            {
                                                afterText = FWT.pmTime.pmTime.ToString();
                                            }
                                        }
                                        c.Text = buttonText + System.Environment.NewLine + morningText + "-" + afterText.ToString();


                                        //break;
                                    }

                                }
                            }
                            else
                            {
                                break;
                            }
                            //}

                            #endregion
                            memberState = new MemberChuQingStatistics(_name, _late, _inTime, _question, _noSignOff, _noData);
                            MemberChuQingList.Add(memberState);

                            B_InTime.Text = "准时" + System.Environment.NewLine + memberState.workerOnTime + "天";
                            B_Late.Text = "迟到" + System.Environment.NewLine + memberState.workerIsLate + "天";
                            B_NotShowUp.Text = "旷工" + System.Environment.NewLine + memberState.BadData + "天";
                            B_Question.Text = "未知" + System.Environment.NewLine + memberState.dataInQuestion + "天";
                            B_NoSignOff.Text = "未打卡下班" + System.Environment.NewLine + memberState.workerNotSignOff + "天";

                            if (_eventAdded != true)
                            {
                                c.MouseClick += c_MouseClick;
                            }
                        }
                    }

                }
                _eventAdded = true;
                 */
                    #endregion

                }
                else
                {

                }
                //}
                int a = listPassedMember.Count();
                toolStripProgressBar1.Value = a;
            }
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

                        Leave_Reason_Date _reasonAndDate = new Leave_Reason_Date(partText[0], "事假");
                        Member_Leave _memberChuQing = new Member_Leave(comboxMember.SelectedItem.ToString(), _reasonAndDate);
                        #region aaa
                        /*
                        string[] _tempText=((Control)sender).Text.Split(new[] { "\r\n" }, StringSplitOptions.None); //将按钮名字分为两个部分,由"\r\n"分割
                        string[] _tempText1=_tempText[1].Split(new[] { "-" }, StringSplitOptions.None); //将按钮名字分为两个部分,由"\r\n"分割
                        string partText1=_tempText1[0]; //早上
                        string partText2=_tempText1[1]; //晚上

                        if (partText1 == null && partText2 == null)
                        {
                            foreach (MemberChuQingStatistics MCQS in MemberChuQingList)
                            {
                                if (MCQS.workerName == comboxMember.SelectedItem.ToString())
                                {
                                    MCQS.BadData -= 1;
                                    break;
                                }
                            }
                        }
                        else if (partText1 != null) //如果出差/请假则不记录出勤状态
                        {

                            //AfternoonTime = fom.SetLimDissmisTime;
                            LimitMorningTime = fom.SetLimShowUpTime;
                            TimeSpan time1 = Convert.ToDateTime(partText1).TimeOfDay;
                            TimeSpan limitTime = LimitMorningTime.TimeOfDay;
                            if (time1 >= limitTime) //如果已经被记录为迟到着删除.
                            {
                                foreach (MemberChuQingStatistics MCQS in MemberChuQingList)
                                {
                                    if (MCQS.workerName == comboxMember.SelectedItem.ToString())
                                    {
                                        MCQS.workerIsLate -= 1;
                                        break;
                                    }
                                }
                            }
                            else  //如果记录为准时
                            {
                                foreach (MemberChuQingStatistics MCQS in MemberChuQingList)
                                {
                                    if (MCQS.workerName == comboxMember.SelectedItem.ToString())
                                    {
                                        MCQS.workerOnTime -= 1;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (partText2 == null)  //如果出差/请假则不记录出勤状态   
                        {
                            foreach (MemberChuQingStatistics MCQS in MemberChuQingList)
                            {
                                if (MCQS.workerName == comboxMember.SelectedItem.ToString())
                                {
                                    MCQS.workerNotSignOff -= 1;
                                    break;
                                }
                            }
                        }
                        */
                        #endregion

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
                if (_Vacance == true || _ChuChai == true || _ShiJia)
                {
                    PaintAndCalcule(comboxMember.SelectedItem.ToString());
                }
            }
        }

        private void PaintAndCalcule(string MemberShown)
        {
            List<Form1.WorkTime> LFWT = ((Form1)this.MdiParent).ListMemberSchedule[comboxMember.SelectedIndex].workTime;

            //定义
            MemberChuQingStatistics memberState;
            int _inTime = 0;    //准时
            int _noSignOff = 0; //没有打卡下班
            int _late = 0;  //迟到
            int _question = 0;  //数据有问题
            int _noData = 0;    //没有数据

            int passedbutton = 0;
            foreach (Control c in this.panel1.Controls)
            {

                if (c is Button)
                {
                    if (listComponant.Contains(c))//当按钮表示所需月份日期时
                    {
                        c.Enabled = true;
                        //c.BackColor = Color.Control;

                        #region 将按钮文字改为日期和考勤时间 并同时改变按钮颜色 同时记录次数

                        //foreach (Form1.WorkTime FWT in LFWT)
                        Form1.WorkTime FWT = LFWT[passedbutton];
                        passedbutton++;
                        //{

                        bool enretard = false;
                        bool alheure = false;
                        bool pasPresente = false;
                        if (FWT._Date.Count != 0)
                        {
                            List<string> textReaded = FWT._Date;
                            if (c.Text != "" && c.Text != null)
                            {
                                string[] yearMonthDay = c.Text.Split(new[] { "" }, StringSplitOptions.None); //将按钮名字分为两个部分,由"空格"分割
                                string[] DayAndTimes = yearMonthDay[0].Split(new[] { "\r\n" }, StringSplitOptions.None); //将按钮名字分为three个部分,由"\r\n"分割
                                string[] DaysButton = DayAndTimes[0].Split(new[] { "." }, StringSplitOptions.None); //将按钮名字分为three个部分,由"-"分割
                                string[] AmPmTimes = DayAndTimes[1].Split(new[] { "-" }, StringSplitOptions.None); //将按钮名字分为three个部分,由"-"分割

                                //Day = yearMonthDay[0].Split(new[] { "\r\n" }, StringSplitOptions.None); //将按钮名字分为three个部分,由"\r\n"分割
                                //
                                if (Convert.ToInt32(textReaded[0]) == Convert.ToInt32(DaysButton[2]))// || Convert.ToInt32(text[1]) == Convert.ToInt32(Day[2]))
                                {
                                    bool _jumpToNextLoop = false;
                                    foreach (Member_Leave M_L in Member_NotShowUp)
                                    {
                                        List<string> a = ((Form1)this.MdiParent).LaDuree;
                                        if (M_L.workerName == comboxMember.SelectedItem.ToString() 
                                            && M_L.memberLeave.date == a[1] + "." + a[2] + "." + DaysButton[2])
                                        {
                                            _jumpToNextLoop = true;
                                            if (M_L.memberLeave.leaveReason == "出差")
                                            {
                                                c.BackColor = Color.Yellow;
                                            }
                                            else if (M_L.memberLeave.leaveReason == "放假")
                                            {
                                                c.BackColor = Color.Aqua;
                                            }
                                            else if (M_L.memberLeave.leaveReason == "节假日")
                                            {
                                                c.BackColor = Color.Orange;
                                            }

                                            break;
                                        }
                                        
                                    }

                                    //如果这个按钮所代表的日期,员工已请假则不记录,继续检查下一个按钮

                                    if (_jumpToNextLoop == true)
                                    {
                                        continue;
                                    }
                                    string buttonText = c.Text;
                                    DateTime defautTime = Convert.ToDateTime((yearMonthDay[0].Split(new[] { "\r\n" }, StringSplitOptions.None)[0]));
                                    DateTime morningTime;
                                    DateTime afterTime;
                                    string morningText = "";
                                    string afterText = "";

                                    //Form1 fom = new Form1();
                                    LimitMorningTime = fom.SetLimShowUpTime;
                                    AfternoonTime = fom.SetLimDissmisTime;

                                    string daytime = yearMonthDay[0].Split(new[] { "\r\n" }, StringSplitOptions.None)[0];

                                    //如果不是周末 或 请假
                                    if (Convert.ToDateTime(daytime).DayOfWeek.ToString() != "Saturday"
                                        && Convert.ToDateTime(daytime).DayOfWeek.ToString() != "Sunday")
                                    {
                                        if ((FWT.amTime == null) && (FWT.pmTime == null)) //没有记录数据
                                        {
                                            pasPresente = true;
                                            
                                                c.BackColor = Color.Violet;
                                                _noData++; 
                                            
                                        }
                                        else
                                        {
                                            if (FWT.amTime == null) //缺失上午的数据
                                            {
                                                    c.BackColor = Color.Pink;
                                                    _question++;
                                            }
                                            else if (FWT.amTime != null && FWT.pmTime != null)  //有上午的时间时
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
                                                    _late++;
                                                }
                                                else
                                                {
                                                    //第一个时间比第二个时间小
                                                    c.BackColor = Color.Lime;
                                                    _inTime++;
                                                } 
                                            
                                            }
                                            else if (FWT.amTime != null && FWT.pmTime == null)
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
                                                    _late++;
                                                }
                                                else
                                                {
                                                    //第一个时间比第二个时间小
                                                    c.BackColor = Color.Lime;
                                                    _inTime++;
                                                } 
                                                
                                            }
                                            if (FWT.pmTime == null) //缺失下午的数据
                                            {
                                                //c.BackColor = Color.Pink;
                                                _noSignOff++;
                                            }
                                            else
                                            {
                                                afterTime = defautTime + FWT.pmTime.pmTime;
                                                afterText = FWT.pmTime.pmTime.ToString();
                                            }
                                        }
                                    }
                                        //周六日
                                    else
                                    {
                                        if (FWT.amTime != null)
                                        {
                                            morningText = FWT.amTime.amTime.ToString();
                                        }
                                        if (FWT.pmTime != null)
                                        {
                                            afterText = FWT.pmTime.pmTime.ToString();
                                        }
                                    }
                                    c.Text = buttonText.Split(new[] { "\r\n" }, StringSplitOptions.None)[0] +System.Environment.NewLine + morningText + "-" + afterText.ToString();


                                    //break;
                                }

                            }
                        }
                        else
                        {
                            break;
                        }
                        
                        #endregion
                        

                        

                        if (_eventAdded != true)
                        {
                            c.MouseClick += c_MouseClick;
                        }
                    }
                }
            }
            memberState = new MemberChuQingStatistics(MemberShown, _late, _inTime, _question, _noSignOff, _noData);
            MemberChuQingList.Add(memberState);

            B_InTime.Text = "准时" + System.Environment.NewLine + memberState.workerOnTime + "天";
            B_Late.Text = "迟到" + System.Environment.NewLine + memberState.workerIsLate + "天";
            B_NotShowUp.Text = "旷工" + System.Environment.NewLine + memberState.BadData + "天";
            B_Question.Text = "未知" + System.Environment.NewLine + memberState.dataInQuestion + "天";
            B_NoSignOff.Text = "未打卡下班" + System.Environment.NewLine + memberState.workerNotSignOff + "天";
            _eventAdded = true;
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

            if (((Form1)this.MdiParent).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.MdiParent).LaDuree;
                string _str = a[0] + " -- " + a[3];

            }
        }

        #region change methode, not good
        /*
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
         */
        #endregion

        public void WorkingPassion(List<Form1.Member_Departement_Communications> MCs)
        {
            AfternoonTime = fom.SetLimDissmisTime;
            LimitMorningTime = fom.SetLimShowUpTime;

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
                        string workDay = ((Form1)this.MdiParent).LaDuree[1] + "-" + ((Form1)this.MdiParent).LaDuree[2] + "-" + wt._Date[0]; //从excel文件中读取的日期

                        if (!InTheList(_tempChuQing, EmployerName, workDay))
                        {
                            if (wt.amTime != null && wt.pmTime != null)
                            {
                                //将数据转换为timeSpan格式
                                DateTime _timeLim = LimitMorningTime;
                                TimeSpan limitTimeSpan = _timeLim.TimeOfDay;

                                if (fom.CompareTime(wt.amTime.amTime, limitTimeSpan))
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
                //MemberChuQingState MCQS = new MemberChuQingState(EmployerName, _late, _inTime, _question, _noSignOff, _noData);
                //MemberChuQingList.Add(MCQS);
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

        DateTime[] listDT;

        public void fileTheCalendar()
        {
            List<string> a = ((Form1)this.MdiParent).LaDuree;
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

                        //FunctionsCS fcs = new FunctionsCS();
                        if (numberButton <= Convert.ToInt32(a[6])) //a[6]是月的天数
                        {
                            string nameButton = a[1] + "." + a[2] + "." + (numberButton + 1) + System.Environment.NewLine + "-"; //按钮名称格式为 年-月- 日
                            c.Text = nameButton;

                        }
                        numberButton++;
                        numberDayInMonth++;
                        listComponant.Add(c);
                    }

                    else if (numberDayInMonth > Convert.ToInt32(a[6])) //下个月
                    {
                        //int dayLeft = numberDayInMonth - Convert.ToInt32(a[6]);
                        int _tMoi = Convert.ToInt32(a[2]) + 1;
                        int _tYear = Convert.ToInt32(a[1]);
                        int _tDate = numberDayInMonth - Convert.ToInt32(a[6]);
                        numberDayInMonth++;
                        c.Text = "" + _tYear + "-" + _tMoi + "-" + _tDate;
                    }
                    else
                    {
                        //上个月
                        //界面上第一个按钮表示是星期一,如果当月第一天不是星期一,则表示该按钮表示上个月的倒数第.
                        if (numberDayOfWeek < FirstDayOfMonthInWeek && _pass == false)
                        {
                            int dayLeft = FirstDayOfMonthInWeek - numberDayOfWeek - 1;
                            int _tMoi = 0;
                            int _tYear = 0;
                            int _tDate = 0;
                            //
                            if (Convert.ToInt32(a[2]) == 1)
                            {
                                _tYear = Convert.ToInt32(a[1]) - 1;
                                _tMoi = 12;
                                _tDate = 31 - dayLeft;
                            }
                            else
                            {
                                _tYear = Convert.ToInt32(a[1]);
                                _tMoi = Convert.ToInt32(a[2]);
                                int daysInOneMonth = System.DateTime.DaysInMonth(_tYear, Convert.ToInt32(a[2]));
                                _tDate = daysInOneMonth - dayLeft;
                                //每个月的天数
                                
                            }

                            c.Text = "" + _tYear + "-" + _tMoi + "-" + _tDate;
                        }
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


            if (_startProcecs == true)
            {
                if (B_ChuChai.FlatStyle != FlatStyle.Flat)
                {
                    _ChuChai = true;
                    B_ChuChai.FlatStyle = FlatStyle.Flat; //been selected
                    B_ChuChai.BackColor = Color.Olive;
                    toolinfor.Text = "设置出差日期";

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
                    toolinfor.Text = "";
                }
            }
        }

        private void B_ShiJia_Click(object sender, EventArgs e)
        {


            if (_startProcecs == true)
            {
                if (B_ShiJia.FlatStyle != FlatStyle.Flat)
                {
                    _ShiJia = true;
                    B_ShiJia.FlatStyle = FlatStyle.Flat;//been selected
                    B_ShiJia.BackColor = Color.Teal;
                    toolinfor.Text = "设置事假日期";
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
                    toolinfor.Text = "";
                }
            }

        }

        private void B_Vacance_Click(object sender, EventArgs e)
        {

            if (_startProcecs == true)
            {
                if (B_Vacance.FlatStyle != FlatStyle.Flat)
                {
                    _Vacance = true;
                    B_Vacance.FlatStyle = FlatStyle.Flat;//been selected
                    B_Vacance.BackColor = Color.DarkGoldenrod;
                    toolinfor.Text = "设置节假日日期";
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
                    toolinfor.Text = "";
                }
            }

        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolTimer.Text = DateTime.Now.ToString();
        }

        private void comboxMember_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void Member_QingJia_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult messageboxResult = MessageBox.Show("确认关闭么?", "注意", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (messageboxResult == DialogResult.OK)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }



    } 
}
