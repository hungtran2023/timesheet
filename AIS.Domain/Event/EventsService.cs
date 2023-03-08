using AIS.Data;
using System;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Linq;
using AIS.Domain.Base;
using AIS.Domain.Common.Constants;

namespace AIS.Domain.Event
{
    public class EventsService : Service<ATC_Events>, IEventsService
    {
        public EventsService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public List<SelectListItem> GetEventsList()
        {
            List<SelectListItem> listOfEvent = new List<SelectListItem>();
            var exceptList = new List<int> { NumberConstants.Project, NumberConstants.PersonalTimeEvent, NumberConstants.GeneralAdmin, NumberConstants.Holiday };
            var listOfEvents = this.FindAll().ToList();
            foreach (var item in listOfEvents)
            {
                if (!exceptList.Contains(item.EventID))
                {
                    var eventItem = new SelectListItem
                    {
                        Value = item.EventID.ToString(),
                        Text = item.EventName.ToString()
                    };
                    listOfEvent.Add(eventItem);
                }
            }
            return listOfEvent.OrderBy(i => i.Text).ToList();
        }
    }
}
