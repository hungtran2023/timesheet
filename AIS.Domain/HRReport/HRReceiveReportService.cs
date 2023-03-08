using System;
using System.Collections.Generic;
using AIS.Data;
using System.Linq;
using System.Web.Mvc;
using AIS.Domain.Base;

namespace AIS.Domain.HRReport
{
    public class HRReceiveReportService : Service<HR_ReceiveReport > , IHRReceiveReportService
    {
        public HRReceiveReportService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }
       
        public List<SelectListItem> GetAuthoriserList(int userId)
        {
            List<SelectListItem> items = new List<SelectListItem>();
            var listData = this.FindAll().Where(t => t.UserID != userId).ToList().GroupBy(t => t.ObjId).Select(t => t.First());
            items.Add(new SelectListItem() { Value = "", Text = "" });
            foreach (var item in listData)
            {
                var temp = new SelectListItem()
                {
                    Value = item.UserID.ToString(),
                    Text = item.Fullname
                };
                items.Add(temp);
            }
            return items.OrderBy(n => n.Text).ToList();
        }
    }
}
