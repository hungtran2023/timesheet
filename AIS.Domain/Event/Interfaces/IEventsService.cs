using AIS.Data;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace AIS.Domain.Event
{
    public interface IEventsService : IService<ATC_Events>
    {
        List<SelectListItem> GetEventsList();
    }
}
