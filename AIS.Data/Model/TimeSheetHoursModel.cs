using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class TimeSheetHoursModel
    {
        public decimal TotalHours { get; set; }

        public decimal TotalHourOverTime { get; set; }


    }
    public class TimeSheetViewLeaveModel
    {
        public int YearAN { get; set; }

        public DateTime DateFrom { get; set; }

        public DateTime DateTo { get; set; }

        public decimal MonthstoCalLeavedue { get; set; }

        public string ANType { get; set; }
       
        public decimal WorkingHours { get; set; }

        public decimal MoreHours { get; set; }

        public int MaxLeaveDue { get; set; }

        public decimal RateperYear { get; set; }

        public decimal AfterExpired { get; set; }

        public decimal KeepPassYear { get; set; }

        public int NumberOfYear { get; set; }

        public decimal ApplicationBy { get; set; }

        public decimal BeforeExpired { get; set; }

        public decimal RateByYTD { get; set; }

        public decimal Leavedue { get; set; }
    }

    public class TimeSheetExpiredViewLeave
    {
        public int ExpiredDay { get; set; }

        public int ExpiredMonth { get; set; }
    }

}
