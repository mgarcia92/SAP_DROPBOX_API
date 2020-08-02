using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneratorSap
{
    static class DateTimeExt
    {
        public static DateTime FirstDateInWeek(this DateTime dt)
        {
            while (dt.DayOfWeek != System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.FirstDayOfWeek)
                dt = dt.AddDays(-1);
            return dt;
        }

        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
            DayOfWeek fdow = ci.DateTimeFormat.FirstDayOfWeek;
            return DateTime.Today.AddDays(-(DateTime.Today.DayOfWeek - fdow));
        }
    }
}
