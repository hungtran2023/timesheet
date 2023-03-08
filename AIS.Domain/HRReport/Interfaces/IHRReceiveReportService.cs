using System;
using System.Collections.Generic;
using AIS.Data;
using AIS.Domain.Base;
using System.Web.Mvc;

namespace AIS.Domain.HRReport
{
    public interface IHRReceiveReportService : IService<HR_ReceiveReport>
    {
        List<SelectListItem> GetAuthoriserList(int userId);
    }
}
