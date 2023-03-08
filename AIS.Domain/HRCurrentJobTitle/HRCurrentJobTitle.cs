using System;
using AIS.Data;
using System.Collections.Generic;
using System.Web.Mvc;
using AIS.Domain.Base;
using System.Linq;

namespace AIS.Domain.HRCurrentJobTitle
{
    public class HRCurrentJobTitle : Service<HR_CurrentJobtitle>, IHRCurrentJobTitle
    {
        public HRCurrentJobTitle(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public string Get(int staffId)
        {
            var result = FindAll().Where(t => t.StaffID == staffId).ToList();
            return result.Count > 0 ? result.First().JobTitle : "";
        }
    }
}
