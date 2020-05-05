using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ReportMaster
{
    public class SourceDataRow
    {
        public DateTime RegDate { get; set; }
        public int Number { get; set; }
        public int PlanValue { get; set; }
        public int ActualValue { get; set; }
    }

    public class ReportDataRow
    {
        public DateTime RegDate { get; set; }
        public string WeekDay { get { return CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(RegDate.DayOfWeek); } }
        public int Number { get; set; }
        public int PlanValue { get; set; }
        public int ActualValue { get; set; }
        public int UnactedValue 
        {
            get { return PlanValue - ActualValue;  }
        }

        public string GetWeekDay()
        {
            string day = "";

            switch (RegDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    day = "Пн";
                    break;
                case DayOfWeek.Tuesday:
                    day = "Вт";
                    break;
                case DayOfWeek.Wednesday:
                    day = "Ср";
                    break;
                case DayOfWeek.Thursday:
                    day = "Чт";
                    break;
                case DayOfWeek.Friday:
                    day = "Пт";
                    break;
                case DayOfWeek.Saturday:
                    day = "Сб";
                    break;
                case DayOfWeek.Sunday:
                    day = "Вс";
                    break;
            }

            return day;
        }
    }
}
