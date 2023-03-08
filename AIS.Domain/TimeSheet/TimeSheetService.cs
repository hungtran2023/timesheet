using System;
using AIS.Data;
using System.Linq;
using System.Collections.Generic;
using AIS.Domain.Base;
using AIS.Domain.Employee;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;

namespace AIS.Domain.TimeSheet
{
    public class TimeSheetService : Service<ATC_Timesheet>, ITimeSheetService
    {
        public TimeSheetService(IUnitOfWork _unitofwork)
            : base(_unitofwork)
        {

        }
        public IEnumerable<object> ListOfTimesheet(String year) {
            return GetRepository().GetListByType(StringConstants.TableTimesheet + year) as IEnumerable<object>;
        }

        public IEnumerable<Object> GetDaysOff(int Year)
        {
            String temp = Year.ToString();
            if (Year == DateTime.Now.Year)
            {
                temp = String.Empty;
            }
            return (IEnumerable<Object>)GetListByType(StringConstants.TableTimesheet + temp);
        }
    }
}
