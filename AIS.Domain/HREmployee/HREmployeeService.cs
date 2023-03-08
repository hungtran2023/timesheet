using AIS.Data;
using AIS.Data.EntityBase.StoredProcedures;
using AIS.Data.Model;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AIS.Domain.HREmployee
{
    public class HREmployeeService : Service<HR_Employee>, IHREmployeeService
    {
       
        public HREmployeeService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

       

        public HR_Employee GetEmployeeInfoById(int Id)
        {
            var result = this.FindAll().Where(e => e.PersonID == Id).FirstOrDefault();
            return result;
        }

        public List<HR_Employee> GetEmployees()
        {
            var result = this.FindAll().ToList();
            return result;
        }
    }
}
