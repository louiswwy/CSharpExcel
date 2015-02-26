﻿using System;
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
    public partial class Calender : Form
    {
        public Calender()
        {
            InitializeComponent();
        }

        List<DateTime> ListDate = new List<DateTime>();
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectionRange sr = new SelectionRange();
            sr.Start = DateTime.Parse(this.textBox1.Text);
            sr.End = DateTime.Parse(this.textBox2.Text);
            this.monthCalendar1.SelectionRange = sr;
        }
        
        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            if (monthCalendar1.SelectionRange.Start.Date.Month==monthCalendar1.SelectionRange.End.Date.Month)
            {
                this.Text = monthCalendar1.SelectionRange.Start.Date.Month.ToString() + "月放假日期"; 
            }
            else
            {
                this.Text = monthCalendar1.SelectionRange.Start.Date.Month.ToString() + "月 -" 
                    + monthCalendar1.SelectionRange.End.Date.Month.ToString() + "月放假日期"; 
            }
            string[] NameWeeks = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            listBox1.Items.Clear();
            this.textBox1.Text = monthCalendar1.SelectionRange.Start.Date.ToShortDateString();
            this.textBox2.Text = monthCalendar1.SelectionRange.End.Date.ToShortDateString();
            if (!ListDate.Contains(monthCalendar1.SelectionRange.Start.Date))
            ListDate.Add(monthCalendar1.SelectionRange.Start.Date);

            if (!ListDate.Contains(monthCalendar1.SelectionRange.End.Date))
            ListDate.Add(monthCalendar1.SelectionRange.End.Date);
            foreach (DateTime dt in ListDate)
            {
                StringBuilder str = new StringBuilder();
                str.Append(dt.ToShortDateString() + "  " + NameWeeks[Convert.ToInt32(dt.DayOfWeek)]);
                listBox1.Items.Add(str.ToString());
            }
            
            //listBox1.Items.Add(monthCalendar1.SelectionRange.End.Date.ToShortDateString());
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {

        }
    }
}
