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
        public Member_QingJia()
        {
            InitializeComponent();
            foreach (Control c in this.panel1.Controls)
            {
                if (c is Button)
                {
                    c.MouseClick += c_MouseClick;
                }
            }
        }

        void c_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ((Control)sender).BackColor=Color.Lavender;
            }
        }

        List<string> ListOfMemberName = new List<string>();
        public class Member_ChuQing
        {
            private string _day;
            public string workerName
            {
                get { return _day; }
                set { this._day = value; }
            }

            //private 
        }

        /*************/
        bool _ChuChai = false;
        bool _ShiJia = false;
        bool _Vacance = false;
        /*************/



        private void B_Valide_Click(object sender, EventArgs e)
        {

        }

        private void B_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Member_QingJia_Load(object sender, EventArgs e)
        {

            if (((Form1)this.Owner).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.Owner).LaDuree;
                string _str = a[0] + " -- " + a[4];

                DisplayTime(_str);
                fileTheCalendar(a);

                groupBox1.Text = a[2] + "月月历";
                foreach (Form1.Member_Departement_Communications item in ((Form1)this.Owner).ListMemberSchedule)
                {
                    ListOfMemberName.Add(item.name);
                    comboxMember.Items.Add(item.name);
                }

                //MessageBox.Show(a);
            }
            else
            {
                MessageBox.Show("需要先倒入Excel文件.", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                this.Close();
            }
            
        }

        private void comboxMember_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string _name = ((Form1)this.Owner).ListMemberSchedule[comboxMember.SelectedIndex].name;
            List<Form1.WorkTime> _lWorkTime = ((Form1)this.Owner).ListMemberSchedule[comboxMember.SelectedIndex].workTime;
            MessageBox.Show(_name);
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

        public void fileTheCalendar(List<string> textMoi)
        {
            //长方形75,25 
            //间距 6
            int Year = Convert.ToInt32( textMoi[1].ToString());
            int Month =Convert.ToInt32( textMoi[2].ToString());
            int DayS = Convert.ToInt32( textMoi[3].ToString());
            int DayE = Convert.ToInt32( textMoi[6].ToString());
            string[,] tableMonth;

            listDT = new DateTime[DayE - DayS + 1];

            for (int DayInMonth = 0; DayInMonth < DayE; DayInMonth++)
            {
                DateTime dt = new DateTime(Year, Month, DayInMonth+1);

                listDT[DayInMonth] = dt;

                
            }

            int num = 0;
            int colNum = 0;


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
                B_Vacance.BackColor = Color.DarkOrange;
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

        private void Member_QingJia_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Click(object sender, EventArgs e)
        {

        }




    } 
}
