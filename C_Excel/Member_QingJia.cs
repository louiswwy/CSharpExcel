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
        }

        List<string> ListOfMemberName = new List<string>();




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
                string _str = a[0] + " -- " + a[3];

                Font textFont = new Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point);
                SizeF _size = TextSize(_str, textFont);

                // Settings to generate a New Label
                Label lbl = new Label();   // Create the Variable for Label
                lbl.Name = "MyNewLabelID"; // Identify your new Label
                lbl.Text = _str;

                // Create Variables to Define "X" and "Y" Locations
                //var lblLocX = lbl.Location.X;
                //var lblLoxY = lbl.Location.Y;

                lbl.SetBounds((this.Size.Width - Convert.ToInt32(_size.Width)) / 2, 0, Convert.ToInt32(_size.Width), Convert.ToInt32(_size.Height));
                //Set your Label Location Here
                //lblLocX = 500;
                //lblLoxY = 77;

                this.Controls.Add(lbl);
                //label8.Location = 
                //label8.Location = new Point(delete.Location.Y, 40);//设置纵坐标y 


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

        private void Member_QingJia_Resize(object sender, EventArgs e)
        {

            /*
            if (((Form1)this.Owner).LaDuree.Count != 0)
            {
                List<string> a = ((Form1)this.Owner).LaDuree;
                string _str = a[0] + " -- " + a[3];

                Font textFont = new Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point);
                SizeF _size = TextSize(_str, textFont);

                // Settings to generate a New Label
                Label lbl = new Label();   // Create the Variable for Label
                lbl.Name = "MyNewLabelID"; // Identify your new Label
                lbl.Text = _str;

                // Create Variables to Define "X" and "Y" Locations
                var lblLocX = lbl.Location.X;
                var lblLoxY = lbl.Location.Y;

                lbl.SetBounds((this.Size.Width - Convert.ToInt32(_size.Width)) / 2, 0, Convert.ToInt32(_size.Width), Convert.ToInt32(_size.Height));
                //Set your Label Location Here
                lblLocX = 500;
                lblLoxY = 77;

                this.Controls.Add(lbl);
            }*/
        }
    } 
}
