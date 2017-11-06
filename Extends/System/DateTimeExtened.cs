using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System
{
    /// <summary>
    /// 时间扩展
    /// </summary>
    public static partial class DateTimeExtened
    {
        /// <summary>
        /// 根据日期计算日期周数（以周日为一周的第一天）
        /// </summary>
        /// <param name="dateTime">日期</param>
        /// <returns>日期周数</returns>
        public static int WeekOfYear(this DateTime dateTime)
        {
            int day = DateTime.Parse(string.Format("{0}-1-1 0:0:0", dateTime.Year)).DayOfWeek.GetHashCode() - 1;
            int week = (dateTime.DayOfYear + day) / 7 + 1;
            return week;
        }

        /// <summary>
        /// 日期时间的所在周的第一天
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static DateTime FirstDayOfWeek(this DateTime dateTime)
        {
            return dateTime.Date.AddDays(0 - dateTime.DayOfWeek.GetHashCode());
        }

        /// <summary>
        /// 日期时间的所在周的最后一天
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static DateTime LastDayOfWeek(this DateTime dateTime)
        {
            return dateTime.Date.AddDays(6 - dateTime.DayOfWeek.GetHashCode());
        }

        /// <summary>
        /// 转换为时间昵称
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static string ToNickString(this DateTime dateTime)
        {
            if (dateTime.Date == DateTime.Now.Date)
            {
                var span = DateTime.Now.Subtract(dateTime);
                if (span.TotalMinutes < 30)
                {
                    return Math.Ceiling(span.TotalMinutes) + " 分钟前";
                }
                if (span.TotalMinutes < 60)
                {
                    return "半小时前";
                }

                return dateTime.ToString("HH:mm");
            }

            if (dateTime.AddDays(1).Date == DateTime.Now.Date)
            {
                return "昨天 " + dateTime.ToString("HH:mm");
            }
            if (dateTime.AddDays(2).Date == DateTime.Now.Date)
            {
                return "前天 " + dateTime.ToString("HH:mm");
            }

            if (dateTime.Year == DateTime.Now.Year)
            {
                return dateTime.ToString("MM月dd日");
            }
            return dateTime.ToString("yyyy年MM月dd日");
        }

        /// <summary>
        /// 转为日期格式
        /// </summary>
        /// <param name="dateTime">时间</param>
        /// <returns></returns>
        public static string ToDateString(this DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd");
        }
    }
}
