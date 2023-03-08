using System;
using System.Collections.Generic;
using AIS.Data;
using AIS.Domain.Base;

namespace AIS.Domain.Employee
{
    public interface IEmployeeService : IService<ATC_Employees>
    {
        ATC_Employees FindLogiUser(int id);
        IEnumerable<ATC_Employees> StaffsForManager(int ManagerId);
        decimal GetWorkingTime(DateTime date, int staffId);
        String GetFullName(int userId);
        String GetFullName(ATC_Employees user);
        decimal CurrentWorkingHours(int StaffId);
        IEnumerable<ATC_SalaryStatus> GetSalaryStatusBetween(DateTime from, DateTime to, int StaffId);
    }
}
