using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AVC_ClareData.PublicClass
{
    public class DateTimeOperation
    {
        /// <summary>
        /// 获取系统当前日期对应周度，周度获取原则：若今年1月1日不为星期一，则去年包含最后一个星期一的不足7天的日子归到今年的第一周
        /// </summary>
        /// <returns>返回周度格式为yyWww，即形如14W01的格式</returns>
        public static string Week()
        {
            DateTime dtNow = DateTime.Now;
            int weekOfYear = 1;
            for (int i = dtNow.DayOfYear - 1; i > 0; i--)
            {
                DateTime temp = dtNow.AddDays(-i);
                if (temp.DayOfWeek.ToString().ToLower() == "sunday")
                    weekOfYear++;
            }
            string week;
            if (weekOfYear < 10)
                week = dtNow.Year.ToString().Substring(2, 2) + "W0" + weekOfYear.ToString();
            else
                week = dtNow.Year.ToString().Substring(2, 2) + "W" + weekOfYear.ToString();
            return week;
        }

        /// <summary>
        /// 获取指定日期对应周度，周度获取原则：若今年1月1日不为星期一，则去年包含最后一个星期一的不足7天的日子归到今年的第一周
        /// </summary>
        /// <param name="dateTime">日期</param>
        /// <returns>返回周度格式为yyWww，即形如14W01的格式</returns>
        public static string Week(DateTime dateTime)
        {
            DateTime dtNow = dateTime;
            int weekOfYear = 1;
            for (int i = dtNow.DayOfYear - 1; i > 0; i--)
            {
                DateTime temp = dtNow.AddDays(-i);
                if (temp.DayOfWeek.ToString().ToLower() == "sunday")
                    weekOfYear++;
            }
            string week;
            if (weekOfYear < 10)
                week = dtNow.Year.ToString().Substring(2, 2) + "W0" + weekOfYear.ToString();
            else
                week = dtNow.Year.ToString().Substring(2, 2) + "W" + weekOfYear.ToString();
            return week;
        }

    }
}
