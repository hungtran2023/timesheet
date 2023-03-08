using System;
using System.Collections.Generic;
using AIS.Data;
using System.Web.Mvc;
using AIS.Domain.Base;

namespace AIS.Domain.Department
{
    public interface IDepartmentService : IService<ATC_Department>
    {
        List<SelectListItem> GetDepartmentList();
    }
}
