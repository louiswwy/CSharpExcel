using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace C_Excel
{
    class FunctionsCS
    {
        private Form1.WorkTime _dateTime;
        public Form1.WorkTime dateTime
        {
            get { return _dateTime; }
            set { this._dateTime = value; }
        }


        public Form1.WorkTime ConvertStringToDateTime(string strTime, List<string> strDate)
        {
            //Form1.WorkTime dateTime;// = new Form1.WorkTime();
            List<string> MathGroup = new List<string>();
            Form1.AMTime AmTime;
            Form1.PMTime PmTime;

            //当时间格式为xx:xx-yy:yy时
            if (isExMatch(strTime.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))
            {
                AmTime = new Form1.AMTime(Convert.ToDateTime(MathGroup[0]).TimeOfDay);
                PmTime = new Form1.PMTime(Convert.ToDateTime(MathGroup[1]).TimeOfDay);
                dateTime = new Form1.WorkTime(strDate, AmTime, PmTime);//(AmTime, PmTime);

            }

            //时间格式为xx:xx:xx时，只提取前4位数字
            else if (isExMatch(strTime.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d$", out MathGroup))
            {

                //string temp = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                AmTime = new Form1.AMTime(Convert.ToDateTime(MathGroup[0]).TimeOfDay);//.ToShortTimeString()));
                dateTime = new Form1.WorkTime(strDate, AmTime);
                //listWorkTime.Add(_workTime);

            }

            //时间格式为xx:xx:xx 汉字（0-4位）时 //^[\u4e00-\u9fa5]{3}
            else if (isExMatch(strTime.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d):[0-5]?\d[\u4e00-\u9fa5]{0,4}$", out MathGroup))
            {
                //string temp = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                AmTime = new Form1.AMTime(Convert.ToDateTime(MathGroup[0]).TimeOfDay);//.ToShortTimeString()));
                dateTime = new Form1.WorkTime(strDate, AmTime);
                //listWorkTime.Add(_workTime);

            }

            //时间格式为xx:xx-
            else if (isExMatch(strTime.Replace(" ", ""), @"^(20|21|22|23|[0-1]?\d:[0-5]?\d)-$", out MathGroup))
            {
                //dr[dc] = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                AmTime = new Form1.AMTime(Convert.ToDateTime(MathGroup[0]).TimeOfDay);//.ToShortTimeString()));
                dateTime = new Form1.WorkTime(strDate, AmTime);
                //listWorkTime.Add(_workTime);

            }
            //时间格式为-xx:xx
            else if (isExMatch(strTime.Replace(" ", ""), @"^-(20|21|22|23|[0-1]?\d:[0-5]?\d)$", out MathGroup))
            {
                //dr[dc] = Convert.ToDateTime(MathGroup[0]).ToShortTimeString().ToString();
                PmTime = new Form1.PMTime((Convert.ToDateTime(MathGroup[0])).TimeOfDay);//.ToShortTimeString()));
                dateTime = new Form1.WorkTime(strDate, PmTime);
                //listWorkTime.Add(_workTime);

            }

                //20:xx匹配20:xx
            else if (isExMatch(strTime.Replace(" ", ""), @"^([1-9]{1}|[0-1][0-9]|[1-2][0-3]):([0-5][0-9])-([1-9]{1}|[0-1][0-9]|[1-2][0-3]):([0-5][0-9])$", out MathGroup))
            {
                List<string> SplitText = new List<string>();
                for (int a = 0; a < MathGroup.Count; a++)
                {
                    if (MathGroup[a] != null && MathGroup[a] != "")
                    {
                        SplitText.Add(MathGroup[a]);
                    }
                }
                AmTime = new Form1.AMTime(Convert.ToDateTime(MathGroup[0] + ":" + MathGroup[1]).TimeOfDay);
                PmTime = new Form1.PMTime(Convert.ToDateTime(MathGroup[2] + ":" + MathGroup[3]).TimeOfDay);
                dateTime = new Form1.WorkTime(strDate, AmTime, PmTime);//(AmTime, PmTime);
            }
            else
            {
                string aaa = strTime;
                AmTime = new Form1.AMTime(Convert.ToDateTime(strTime).TimeOfDay);//.ToShortTimeString()));
                dateTime = new Form1.WorkTime(strDate, AmTime);
            }


            return dateTime;
        }

        public bool isExMatch(string text, string patten, out List<string> Match)
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
                _isMatch = false;
            Match = _match;
            return _isMatch;
        }

    }
}
