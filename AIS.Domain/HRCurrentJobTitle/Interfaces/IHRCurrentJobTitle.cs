using System;
using System.Collections.Generic;
using AIS.Data;
using System.Web.Mvc;
using AIS.Domain.Base;

namespace AIS.Domain.HRCurrentJobTitle
{
    public interface IHRCurrentJobTitle : IService<HR_CurrentJobtitle>
    {
        String Get(int staffId);
    }
}
