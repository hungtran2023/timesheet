using System;
using AIS.Data;
using System.Linq;
using AIS.Domain.Base;

namespace AIS.Domain.Holiday
{
    public class HolidayService : Service<ATC_Holiday> , IHolidayService
    {
        public HolidayService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public DateTime[] GetAll() {
            var listOfHolidays =  from days in FindAll().ToList()
                   select new
                   {
                      date = new DateTime(days.sYear , days.sMonth , days.sDay)
                   }.date;
            return listOfHolidays.ToArray();
        }
    }
}
