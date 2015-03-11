using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace C_Excel
{
    public partial class WfSetting : Form
    {
        public WfSetting()
        {
            InitializeComponent();
        }
        //public static DateTime LimitShowUpTime;// = Convert.ToDateTime("08:46:00");
        //最早下班时间
        //public static DateTime LimitDismissTime;// = Convert.ToDateTime("17:30:00");
        public class WorkPlan
        {
            private DateTime _showUp;
            public DateTime showUpAt
            {
                get { return _showUp; }
                set { _showUp = value; }
            }

            private DateTime _dissmis;
            public DateTime dissmisAt
            {
                get { return _dissmis; }
                set { _dissmis = value; }
            }

            public WorkPlan()
            {

            }

            public WorkPlan(DateTime ShowUpTime,DateTime DissmisTime)
            {
                this.showUpAt = ShowUpTime;
                this.dissmisAt = DissmisTime;
            }

        }

        XmlDocument xmldoc;
        //XmlNode xmlnode;
        //XmlElement xmlelem;

        /*
         if (!System.IO.File.Exists(this.textBox_xlsPath.Text))
         */


        private void B_Valide_Click(object sender, EventArgs e)
        {
            if(T_ShowUp.Text!="" && T_Dissmis.Text != "")
            {
                try
                {
                    xmldoc = new XmlDocument();//@"..\..\Book.xml")

#if debug
                    string XmlFilePath = @"..\..\DataFile\WorkTime.xml";


#else
                    var folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                    string XmlFilePath = folderPath + @"\WorkTime.xml";//Environment.CurrentDirectory + @"..\..\..\..\DataFile\WorkTime.xml";

#endif

                    xmldoc.Load(XmlFilePath);    //读取指定的XML文档
                    
                    XmlElement xe = xmldoc.DocumentElement;//获取xml文档的根xmlelement

                    if (xe.HasChildNodes)
                    {
                        ///三种查询xml文件数据的方式
                        /// 
                        //string strPath = string.Format("workTime/ShowUpTime");
                        //XmlElement selectXe = (XmlElement)xmldoc.SelectSingleNode(strPath);//selectSingleNode 根据XPath表达式,获得符合条件的第一个节点.
                        ///
                        //string a = selectXe.GetElementsByTagName("ShowUpTime")[0].InnerText;
                        ///
                        //XmlNodeList list = xmldoc.GetElementsByTagName("ShowUpTime");//要查询的的节点名  

                        XmlNodeList listShowUpTime = xmldoc.GetElementsByTagName("ShowUpTime");//要查询的的节点名  
                        XmlNodeList listDissmisTime = xmldoc.GetElementsByTagName("dissmisTime");//要查询的的节点名  

                        string showUpT = listShowUpTime[0].InnerText.ToString();
                        string dissmisT = listDissmisTime[0].InnerText.ToString();
                                                                               //@"^(20|21|22|23|[0-1]?\d:[0-5]?\d)$"
                        if (isExMatch(T_ShowUp.Text.ToString().Replace(" ", ""), @"(20|21|22|23|[0-1]?\d:[0-5]?\d)$")
                            && isExMatch(T_Dissmis.Text.ToString().Replace(" ", ""), @"(20|21|22|23|[0-1]?\d:[0-5]?\d)$"))
                        {
                            if (showUpT != T_ShowUp.Text.ToString() || dissmisT != T_Dissmis.Text.ToString())
                            {
                                xmldoc.GetElementsByTagName("ShowUpTime")[0].InnerText = T_ShowUp.Text.ToString();
                                xmldoc.GetElementsByTagName("dissmisTime")[0].InnerText = T_Dissmis.Text.ToString();
                                xmldoc.Save(XmlFilePath);

                                //LimitShowUpTime = Convert.ToDateTime(T_ShowUp.Text.ToString() + ":00");
                                //LimitDismissTime = Convert.ToDateTime(T_Dissmis.Text.ToString() + ":00");

                                
                                //MessageBox.Show(""+ Form1._limitShowUpTime.ToShortTimeString());
                                //传递值回form1
                                ((Form1)this.Owner).SetLimShowUpTime = Convert.ToDateTime(T_ShowUp.Text.ToString());
                                ((Form1)this.Owner).SetLimDissmisTime = Convert.ToDateTime(T_Dissmis.Text.ToString());
                                MessageBox.Show("以储存下列变更:" + System.Environment.NewLine + "上班时间为:  0"
                                    + T_ShowUp.Text.ToString() + ";" + System.Environment.NewLine + "下班时间为  " 
                                    + T_Dissmis.Text.ToString() + ";", "通知", MessageBoxButtons.OK
                                    , MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                            }
                            else
                            {
                                MessageBox.Show("数据未改变", "通知", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            MessageBox.Show("数据格式不正确", "通知",MessageBoxButtons.OK,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button1);
                        }
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("" + ex, "错误！");
                }
            }
        }

        private void B_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void WfSetting_Load(object sender, EventArgs e)
        {
            try
            {
                
                xmldoc = new XmlDocument();
#if DEBUG
                
                string ProgramInstalPath = "";
                ProgramInstalPath = Application.StartupPath.ToString();

                string XmlFilePath = @"..\..\DataFile\WorkTime.xml";
#else
                //string XmlFilePath = @"..\..\..\..\DataFile\WorkTime.xml";


                string ProgramInstalPath = "";
                //ProgramInstalPath = Application.CommonAppDataPath.ToString();
                var folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                string XmlFilePath = folderPath + @"\WorkTime.xml";//Environment.CurrentDirectory + @"..\..\..\..\DataFile\WorkTime.xml";

#endif

                if (System.IO.File.Exists(XmlFilePath))
                {
                    xmldoc.Load(XmlFilePath);    //读取指定的XML文档
                    XmlNode NodeWorkTime = xmldoc.DocumentElement;  //读取xml的根节点

                    foreach (XmlNode node in NodeWorkTime.ChildNodes)//循环子节点
                    {
                        switch (node.Name)
                        {
                            case "ShowUpTime":
                                if (node.InnerText != "")
                                {
                                    T_ShowUp.Text = node.InnerText;
                                }

                                break;

                            case "dissmisTime":
                                if (node.InnerText != "")
                                {
                                    T_Dissmis.Text = node.InnerText;
                                }
                                break;
                        }
                    }
                }
                else
                {


                    /*string path = XmlFilePath;

                    //XDocument xdoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"),
                    //
                    //                               new XElement("Root", "root"));

                    ///<workTime>
                    /// <ShowUpTime>8:46</ShowUpTime>
                    /// <dissmisTime>17:30</dissmisTime>
                    ///</workTime>
                    XElement root = new XElement("workTime",

                        new XElement("ShowUpTime", "8:46"),
                        new XElement("dissmisTime", "17:30")
                        );

                    root.Save(path);*/
                    MessageBox.Show(XmlFilePath);

                }
            }
            catch (Exception ex)
            {                
                MessageBox.Show("" + ex, "错误");
            }
        }

        public bool isExMatch(string text, string patten)
        {
            bool _isMatch = false;
            Regex Patten = new Regex(patten);
            List<string> _match = new List<string>();
            //if (Regex.IsMatch(text, patten))
            if (Patten.Match(text).Success)
            {
                _isMatch = true;
                for (int num = 1; num < Patten.Match(text).Groups.Count; num++)
                {
                    _match.Add(Patten.Match(text).Groups[num].Value);
                }

            }
            else
            {
                _isMatch = false;
            }

            return _isMatch;
        }
    }
}
