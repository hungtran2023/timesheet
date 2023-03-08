using System.Web.Mvc;
using System.Linq;
using AIS.Domain.TimeSheet;
using AIS.Domain.AnualLeaveDays;
using AIS.Domain.Common.Constants;
using AIS.Domain.AbsenceRequest;

namespace AIS.Controllers
{
    public class OverviewOfRequestHistoryController : BaseController
    {
        private readonly ITimeSheetService _timesheetService = Inject.Service<ITimeSheetService>();
        private readonly IAnualLeaveDaysService _anualLeaveDaysService = Inject.Service<IAnualLeaveDaysService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();
        public ActionResult OVerviewOfRequestHistory()
        {
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            ViewBag.StaffID = UserId;
            ViewBag.FullNameOfUser = UserFullName;
            ViewBag.PositionOfUser = UserPosition;
            ViewBag.CurrentRate = _anualLeaveDaysService.GetCurrentRate();
            ViewBag.BalanceDay = _anualLeaveDaysService.GetAnualLeaveBalance();
            ViewBag.BalanceLastYear = _anualLeaveDaysService.GetAnualLeaveBalanceLastYear();
            ViewBag.LeaveUntilDay = _anualLeaveDaysService.GetLeaveUntilDay();
            ViewBag.TotalHours = _anualLeaveDaysService.GetTotalHours();
            ViewBag.AnualLeaveCurrentYear = _anualLeaveDaysService.GetAnualLeaveCurrentYear();
            ViewBag.AnualLeaveReserved = _anualLeaveDaysService.GetAnualLeaveReserved();
            ViewBag.TotalUnauthorisedDays = serviceHandler.GetTotalUnauthorisedDays(UserId);
            ViewBag.BalanceHours = _anualLeaveDaysService.GetBalanceHours();
            return View();
        }

        // GET: OverviewOfRequestHistory/GetDaysOff/2016
        [HttpGet]
        public ActionResult GetDaysOff(int id)
        {
            var listOfTimeSheet = _timesheetService.GetDaysOff(id);
            if (listOfTimeSheet != null)
            {
                var data = listOfTimeSheet.ToList().Cast<dynamic>()
                .Where(timesheet => timesheet.StaffID == UserId
                        && timesheet.EventID != NumberConstants.PersonalTimeEvent
                        && timesheet.EventID != NumberConstants.GeneralAdmin && timesheet.EventID != NumberConstants.Project);
                var result = (from dates in data
                              select new
                              {
                                  day = dates.TDate.Day,
                                  month = dates.TDate.Month,
                                  isHoliday = dates.EventID == NumberConstants.Holiday
                              }).ToList();
                return Json(result.Distinct(), JsonRequestBehavior.AllowGet);
            }
            return Json(null, JsonRequestBehavior.AllowGet);
        }
    }
}