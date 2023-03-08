using System;
using System.Linq;
using AIS.Data;
using AIS.Domain.AbsenceRequest;
using System.Web;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Constants;

namespace AIS.Domain.AnualLeaveDays
{
    public class AnualLeaveDaysService : IAnualLeaveDaysService
    {
        protected double CurrentRate
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.CurrentRate].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double BalanceDay
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.BalanceDay].ToString());
                    return Math.Round(value,2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double BalanceLastYear
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.BalanceLastYear].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double LeaveUntilDay
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.LeaveUntilDay].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double TotalHours
        {
            get
            {
                return HttpContext.Current.Session[StringConstants.TotalHours] == null ?  0 : Math.Round(Convert.ToDouble(HttpContext.Current.Session[StringConstants.TotalHours].ToString()),2);
            }
        }
        protected double AnualLeaveCurrentYear
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.AnnualLeaveCurrentYear].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double AnualLeaveReserved
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.AnnualLeaveReserved].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
        protected double BalanceHours
        {
            get
            {
                try
                {
                    var value = Convert.ToDouble(HttpContext.Current.Session[StringConstants.BalanceHours].ToString());
                    return Math.Round(value, 2);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }

        public AnualLeaveDaysService()
        {
        }

        public double GetCurrentRate() {

            var result = CurrentRate;
            return result;
        }

        public double GetAnualLeaveBalance()
        {
            var result = BalanceDay;
            return result;
        }

        public double GetAnualLeaveBalanceLastYear()
        {
            var result = BalanceLastYear;
            return result;
        }

        public double GetLeaveUntilDay()
        {
            var result = LeaveUntilDay;
            return result;
        }

        public double GetTotalHours()
        {
            var result = TotalHours;
            return result;
        }

        public double GetAnualLeaveCurrentYear()
        {
            var result = AnualLeaveCurrentYear;
            return result;
        }

        public double GetAnualLeaveReserved()
        {
            var result = AnualLeaveReserved;
            return result;
        }

        public double GetBalanceHours()
        {
            var result = BalanceHours;
            return result;
        }
    }
}
