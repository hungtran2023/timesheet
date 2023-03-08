using AIS.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AIS.Domain.Common.Helper
{
    public static class DateTimeHelper
    {
        private static TimeSpan _startWorkTime = new TimeSpan(8, 0, 0);

        public static String FormatMinutesToTimeString(double minutes)
        {
            var result = "";
            int ConvertToHours = Convert.ToInt16(Math.Truncate(minutes / 60));
            int ConvertToMinutes = Convert.ToInt16(minutes % 60);
            if (ConvertToHours > 1)
            {
                result += String.Format("{0} Hours ", ConvertToHours);
            }
            else if (ConvertToHours == 1)
            {
                result += String.Format("{0} Hour ", ConvertToHours);
            }

            if (ConvertToMinutes > 1)
            {
                result += String.Format("{0} Minutes ", ConvertToMinutes);
            }
            else if (ConvertToMinutes == 1)
            {
                result += String.Format("{0} Minute ", ConvertToMinutes);
            }
            return result;
        }

        public static int BusinessDaysUntil(DateTime StartDate, DateTime EndDate ,DateTime[] holidays)
        {
            int businessDays = 0;
            for (DateTime i = StartDate.Date; i <= EndDate.Date; i = i.AddDays(1))
            {
                if (isWorkday(i,holidays))
                {
                    businessDays++;
                }
            }
            return businessDays;
        }

        public static double GetStartDateOffMinutes(DateTime StartDate, DateTime EndDate, double hours)
        {
            if ((StartDate.DayOfWeek != DayOfWeek.Sunday && StartDate.DayOfWeek != DayOfWeek.Saturday))
            {
                var endBreakTime = StartDate.Date.AddHours(13);
                var startBreakTime = StartDate.Date.AddHours(12);
                if (StartDate.Date == EndDate.Date)
                {
                    if (StartDate > startBreakTime && StartDate < endBreakTime)
                    {
                        StartDate = endBreakTime;
                    }
                    if (EndDate > startBreakTime && EndDate < endBreakTime)
                    {
                        EndDate = startBreakTime;
                    }
                    if (EndDate >= endBreakTime && StartDate <= startBreakTime)
                    {
                        EndDate = EndDate.AddHours(-1);
                    }
                    return (EndDate - StartDate).TotalMinutes;
                }
                if (StartDate > startBreakTime && StartDate < endBreakTime)
                {
                    StartDate = endBreakTime;
                }
                var regularWorkTime = StartDate.Date.AddHours(_startWorkTime.TotalHours + hours);
                var alreadyWorkTimespan = StartDate.Subtract(StartDate.Date.AddTicks(_startWorkTime.Ticks));
                if (alreadyWorkTimespan.TotalHours >= 5)
                {
                    alreadyWorkTimespan = alreadyWorkTimespan.Subtract(new TimeSpan(1, 0, 0));
                }
                var alreadyWorkHours = _startWorkTime.TotalHours + alreadyWorkTimespan.TotalHours;
                TimeSpan startDateOffTime = regularWorkTime.Subtract(StartDate.Date.AddHours(alreadyWorkHours));
                double startDateOffMinute = startDateOffTime.TotalMinutes;
                return startDateOffMinute;
            }
            return 0;
        }

        public static double GetEndDateOffMinutes(DateTime EndDate )
        {
            if (EndDate.DayOfWeek != DayOfWeek.Sunday || EndDate.DayOfWeek != DayOfWeek.Saturday)
            {
                var endBreakTime = EndDate.Date.AddHours(13);
                if (EndDate >= endBreakTime)
                {
                    EndDate = EndDate.AddHours(-1);
                }
                var endDateOffTime = EndDate.Subtract(EndDate.Date.AddHours(_startWorkTime.TotalHours));
                double endDateOffMinute = endDateOffTime.TotalMinutes;
                return endDateOffMinute;
            }
            return 0;
        }

        public static double ConvertMinutesToRoundHours(double minutes)
        {
            return Math.Round((minutes / 60), 2);
        }

        public static String ToFormat(this DateTime value)
        {
            return value.ToString("dd/MM/yyyy - hh:mm tt");
        }

        public static bool isWorkday(DateTime date, DateTime[] holidays)
        {
            return isWeekdays(date) && !holidays.Contains(date.Date);
        }

        public static bool isWeekdays(this DateTime date)
        {
            return date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday;

        }

        public static Dictionary<DateTime, Double> FillHoursAndDate(DateTime StartDate, DateTime EndDate, decimal offHours)
        {
            var results = new Dictionary<DateTime, double>();
            for (DateTime i = StartDate.Date; i <= EndDate.Date; i = i.AddDays(1))
            {
                if (i.DayOfWeek != DayOfWeek.Saturday && i.DayOfWeek != DayOfWeek.Sunday)
                {
                    results.Add(i.Date, (double)offHours);
                }
            }
            return results;
        }

        public static void FillHoursAndDate(List<ATC_SalaryStatus> listOfSalaryStatus, DateTime StartDate, DateTime EndDate, ref Dictionary<DateTime, Double> results)
        {
            if (listOfSalaryStatus.Count() == 1)
            {
                var offHours = Convert.ToDouble(listOfSalaryStatus.First().ATC_WorkingHours.Hours);
                for (DateTime i = StartDate.Date; i <= EndDate.Date; i = i.AddDays(1))
                {
                    if (i.DayOfWeek != DayOfWeek.Saturday && i.DayOfWeek != DayOfWeek.Sunday)
                    {
                        results.Add(i.Date, offHours);
                    }
                }
                return;
            }
            else
            {
                foreach (var item in listOfSalaryStatus)
                {
                    ATC_SalaryStatus previousSalaryStatus = item;
                    ATC_SalaryStatus nextSalaryStatus = null;
                    DateTime startDateTemp = DateTime.Now;
                    DateTime endDateTemp = DateTime.Now;
                    double offHours = Convert.ToDouble(previousSalaryStatus.ATC_WorkingHours.Hours);
                    try
                    {
                        nextSalaryStatus = listOfSalaryStatus[listOfSalaryStatus.IndexOf(item) + 1];
                    }
                    catch (Exception)
                    {
                        nextSalaryStatus = null;
                    }
                    if (nextSalaryStatus == null)
                    {
                        startDateTemp = previousSalaryStatus.SalaryDate.Date;
                        endDateTemp = EndDate.Date.AddDays(-1);
                    }
                    else
                    {
                        if (StartDate.Date >= previousSalaryStatus.SalaryDate.Date)
                        {
                            startDateTemp = StartDate.Date.AddDays(1);
                        }
                        else if (StartDate.Date < previousSalaryStatus.SalaryDate.Date)
                        {
                            startDateTemp = previousSalaryStatus.SalaryDate.Date.AddDays(1);
                        }
                        endDateTemp = nextSalaryStatus.SalaryDate.Date.AddDays(-1);
                    }
                    for (var i = startDateTemp.Date; i <= endDateTemp.Date; i = i.AddDays(1))
                    {
                        if (i.DayOfWeek != DayOfWeek.Saturday && i.DayOfWeek != DayOfWeek.Sunday)
                        {
                            results.Add(i.Date, offHours);
                        }
                    }
                }
            }
        }
    }
}
