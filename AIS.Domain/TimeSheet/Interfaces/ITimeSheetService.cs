using System;
using AIS.Data;
using System.Collections.Generic;
using AIS.Domain.Base;

namespace AIS.Domain.TimeSheet
{
    public interface ITimeSheetService : IService<ATC_Timesheet>
    {
        IEnumerable<Object> GetDaysOff( int Year);
        IEnumerable<object> ListOfTimesheet(String year);
    }
}
