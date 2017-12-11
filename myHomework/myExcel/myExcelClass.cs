using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace myExcel
{
    public static class myExcelClass
    {
        public static void encrypt(this Excel.Range rng, Excel.Range key)
        {
            Excel.Range rng2 = rng.Offset[2, 0];
            //int a = (int) rng.Value;
            //string b = rng.Text;
            rng2.Value = (int)rng.Value ^ (int)key.Value;
        }
        public static List<string> getContent(string web)
        {
            XElement xml = XElement.Load(web);
            string txt = "";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index1) => txt += index1.ToString() + ";" + m.Element("title").Value + "\r\n")
                .Where((n, index2) => index2 < 5)
                .ToList();
            return list;
        }
        public static void writhContent(this Excel.Range rng, string web)
        {
            //var list = getContent(rng.Value);
            //if (rng.Value = null)
            //{
            //    rng.Value = web;
            //}
            //else
            //{
            //    list = getContent(web);
            //}
            var list = getContent(web);
            for (int i = 0; i < list.Count; i++)
            {
                rng.Offset[i + 1, 0].Value = i;
                rng.Offset[i + 1, 1].Value = list[i];
            }
        }
        public static DateTime getTimeFirstStr()
        {
            DateTime getTimeFirst;
            getTimeFirst = DateTime.Now;
            //string timeFirst = getTimeFirst.ToString("yyyy年M月d日 H时m分s秒 fff毫秒");
            return getTimeFirst;
        }
        public static string getTimeDiff(DateTime getTimeFist)
        {
            DateTime getTimeSecond = DateTime.Now;
            TimeSpan getTimeDiff = getTimeFist > getTimeSecond ?
                getTimeFist - getTimeSecond :
                getTimeSecond - getTimeFist;
            string timeDiff = string.Format("间隔时间：{0}天{1}时{2}分{3}秒 {4}毫秒",
                getTimeDiff.Days, getTimeDiff.Hours,
                getTimeDiff.Minutes, getTimeDiff.Seconds,
                getTimeDiff.Milliseconds);
            return timeDiff;
        }
    }
}
