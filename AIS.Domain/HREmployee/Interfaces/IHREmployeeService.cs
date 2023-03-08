using AIS.Data;
using AIS.Data.Model;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;

namespace AIS.Domain.HREmployee
{
    public interface IHREmployeeService : IService<HR_Employee>
    {
        HR_Employee GetEmployeeInfoById(int Id);
        List<HR_Employee> GetEmployees();

       
    }
}
