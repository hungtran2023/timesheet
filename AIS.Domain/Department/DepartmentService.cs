using System;
using AIS.Data;
using System.Collections.Generic;
using System.Web.Mvc;
using AIS.Domain.Base;
using System.Linq;

namespace AIS.Domain.Department
{
    public class DepartmentService : Service<ATC_Department>, IDepartmentService
    {
        public DepartmentService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public List<SelectListItem> GetDepartmentList()
        {
            List<SelectListItem> listOfStatus = new List<SelectListItem>();
            listOfStatus.Add(new SelectListItem { Value= "", Text="" });
            foreach(var item in this.FindAll())
            {
                var department = new SelectListItem
                {
                    Value = item.DepartmentID.ToString(),
                    Text = item.Department.ToString()
                };
                listOfStatus.Add(department);
            }
            return listOfStatus.OrderBy(n => n.Text).ToList();
        }
    }
}
