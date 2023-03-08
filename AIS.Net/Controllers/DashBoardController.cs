using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using AIS.Models;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.HRReport;
using AIS.Domain.AnualLeaveDays;
using AIS.Domain.TimeSheet;
using AIS.Domain.Event;
using AIS.Domain.HREmployee;
using AIS.Domain.Project;
using AIS.Domain.Holiday;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;
using AIS.Domain.Email;
using PagedList;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.Text;
using AIS.Domain.DashBoard.Interfaces;
using System.Web.UI.WebControls;
using AIS.Data.Model;

namespace AIS.Controllers
{
    public class DashBoardController : MenuBaseController
    {
        private readonly IDashBoardService _dashBoardService = Inject.Service<IDashBoardService>();

        public ActionResult DashBoard()
        {
            var total = _dashBoardService.GetTotalHoursWorkOfMonth(UserId, DateTime.Now.Month).FirstOrDefault();

            ViewBag.Total = total.TotalHours;

            ViewBag.TotalOT = total.TotalHourOverTime;

            var totalLeave = _dashBoardService.GetDataTotalViewLeave(UserId, DateTime.Now);

            var totalLeaveLast = totalLeave.LastOrDefault();

            var totalLeaveDueUntilLeavedue = totalLeave.Where(x => x.YearAN == DateTime.Now.Year).Select(x => x.Leavedue).Sum();

            var totalLeaveDueUntilmorhours = totalLeave.Where(x => x.YearAN == DateTime.Now.Year).Select(x => x.MoreHours).Sum();

            var totalAnnualleavehours = totalLeave.Where(x => x.YearAN == DateTime.Now.Year && x.ANType != "Reserved").Select(x => x.ApplicationBy).Sum();

            ViewBag.CurrentRate = Math.Round((totalLeaveLast.RateByYTD + totalLeaveLast.RateperYear) / 12, 2);

            ViewBag.LeaveDueUntil = Math.Round((totalLeaveDueUntilLeavedue + totalLeaveDueUntilmorhours), 2);

            var viewExpired = _dashBoardService.GetDataExpiredView(UserId).FirstOrDefault();


            var employees = _dashBoardService.GetAllEmoloyeeCurrent();

            string dateexpired = string.Empty;

            string displaytextBrought = "show";

            if (viewExpired != null)
            {
                dateexpired = viewExpired.ExpiredDay + "/" + viewExpired.ExpiredMonth;
            }

            if (!string.IsNullOrEmpty(dateexpired))
            {
                displaytextBrought = "None";
            }

            ViewBag.display = displaytextBrought;
            ViewBag.Annualleave = totalAnnualleavehours;
            ViewBag.Balance = Math.Round(totalLeaveLast.AfterExpired, 2);

            ViewBag.TotalEmployes = employees.Count();


            var messageContent = _dashBoardService.GetAllTopMesageContent().ToList();
            ViewBag.Content = messageContent;

            return View();
        }
        [HttpPost]
        public ActionResult LoadChartByDepartment(string department = "")
        {
            ViewBag.Department = _dashBoardService.GetAllDepartment();
            if (department == "")
            {
                department = "1";
            }
            var model = new object();
            var data = _dashBoardService.GetTreeViewData(int.Parse(department));
            var groupListChart = data.Select(x => new AIS.Data.Model.EntityChart { dateHireDisplay = x.dateHireDisplay, id = x.id, head = x.head, contents = x.contents, parentID = x.parentID, level = x.level }).ToList();
            List<AIS.Data.Model.EntityChart> list_tree = new List<AIS.Data.Model.EntityChart>();
            if (department == "1")
            {

                list_tree = groupListChart
                                .Where(c => c.level == 3)
                                .Select(c => new AIS.Data.Model.EntityChart()
                                {
                                    id = c.id,
                                    head = c.head,
                                    parentID = c.parentID,
                                    contents = c.contents,
                                    level = c.level,
                                    dateHireDisplay = c.dateHireDisplay,
                                    teamdescription = c.teamdescription,
                                    children = GetChildren(groupListChart, c.id)
                                })
                                .ToList();
            }
            else
            {
                list_tree = groupListChart
                            .Where(c => c.level == 2)
                            .Select(c => new AIS.Data.Model.EntityChart()
                            {
                                id = c.id,
                                head = c.head,
                                parentID = c.parentID,
                                contents = c.contents,
                                level = c.level,
                                dateHireDisplay = c.dateHireDisplay,
                                teamdescription = c.teamdescription,
                                children = GetChildren(groupListChart, c.id)
                            })
                            .ToList();
            }

            model = JsonConvert.SerializeObject(list_tree);
            ViewBag.OrgChart = model;
            return PartialView(@"_TreeChart", model);
        }
        public ActionResult TreeChart()
        {
            ViewBag.Department = _dashBoardService.GetAllDepartment();

            var model = new object();
            var data = _dashBoardService.GetTreeViewData(1);
            var groupListChart = data.Select(x => new AIS.Data.Model.EntityChart { dateHireDisplay = x.dateHireDisplay, id = x.id, head = x.head, contents = x.contents, parentID = x.parentID, level = x.level }).ToList();
            List<AIS.Data.Model.EntityChart> list_tree = new List<AIS.Data.Model.EntityChart>();

            list_tree = groupListChart
                            .Where(c => c.level == 3)
                            .Select(c => new AIS.Data.Model.EntityChart()
                            {
                                id = c.id,
                                head = c.head,
                                parentID = c.parentID,
                                contents = c.contents,
                                level = c.level,
                                dateHireDisplay = c.dateHireDisplay,
                                teamdescription = c.teamdescription,
                                children = GetChildren(groupListChart, c.id)
                            })
                            .ToList();

            model = JsonConvert.SerializeObject(list_tree);
            ViewBag.OrgChart = model;

            return View(model);
        }
        public static List<AIS.Data.Model.EntityChart> GetChildren(List<AIS.Data.Model.EntityChart> entitycharts, int parentId)
        {
            try
            {
                return entitycharts
                        .Where(c => c.parentID == parentId && c.id != parentId)
                        .Select(c => new AIS.Data.Model.EntityChart
                        {

                            id = c.id,
                            head = c.head,
                            parentID = c.parentID,
                            contents = c.contents,
                            level = c.level,
                            dateHireDisplay = c.dateHireDisplay,
                            teamdescription = c.teamdescription,
                            children = GetChildren(entitycharts, c.id)
                        })
                        .ToList();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);

            }
        }
        [HttpGet]
        public PartialViewResult BookingTimeSheet()
        {
        
            AIS.Data.Model.TimeSheetBookingModel model = new Data.Model.TimeSheetBookingModel();
          
            model.CurrentDay = DateTime.Now.Date.ToShortDateString();
            model.subTasks = _dashBoardService.GetAllProjectTypes(UserId, "");
            model.resultBooks = _dashBoardService.GetAllTimeSheetBookedView(UserId, DateTime.Now.Date).ToList();
            var totalHours = model.resultBooks.Sum(x => x.Hours).ToDouble();
            ViewBag.TotalHours = totalHours;
            return PartialView("~/Views/DashBoard/_ModalBookingTimeSheet.cshtml", model);
        }


        [HttpGet]
        public JsonResult GetTimeSheetLogData()
        {
            var resultbooks = _dashBoardService.GetAllTimeSheetBookedView(UserId, DateTime.Now.Date).ToList();
            List<AIS.Data.Model.TimeSheetBookedView> list = new List<Data.Model.TimeSheetBookedView>();
            list = resultbooks.ToList();

            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;
        }

        [HttpPost]
        public JsonResult DoBookingTimeSheet(TimeSheetBookingModel model)
        {        
            if(model.Note==null)
            {
                model.Note = String.Empty;
            }
         
           var flag = _dashBoardService.DoLogTimeSheetNormal(DateTime.Now.Date, UserId, model.SubtaskID, model.Hours, model.Note, model.IsGenerateAdmin);

            model.resultBooks = _dashBoardService.GetAllTimeSheetBookedView(UserId, DateTime.Now.Date).ToList();
            var totalHours = model.resultBooks.Sum(x => x.Hours).ToDouble();

            model.CurrentDay = DateTime.Now.Date.ToShortDateString();
            model.subTasks = _dashBoardService.GetAllProjectTypes(UserId, "");
            var resultbooks = _dashBoardService.GetAllTimeSheetBookedView(UserId, DateTime.Now.Date).ToList();
            model.resultBooks = resultbooks;
            ViewBag.TotalHours = totalHours;
            model.Hours = 0;
            model.Note = "";
            model.IsGenerateAdmin = false;
                    
            List<AIS.Data.Model.TimeSheetBookedView> list = new List<Data.Model.TimeSheetBookedView>();
            list = resultbooks.ToList();
          
         

            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;
           // return Json("~/Views/DashBoard/_ModalBookingTimeSheet.cshtml", model);
        }

        public JsonResult GetAPKAutoComplete(string term = "")
        {
          var projectTypes = _dashBoardService.GetAllProjectAutoComplete(UserId);
            var objAPK = projectTypes
                            .Where(c => c.Text.ToUpper()
                            .Contains(term.ToUpper()))
                            .Select(c => new { Name = c.Text, APK = c.Value })
                            .Distinct().ToList();
            return Json(objAPK, JsonRequestBehavior.AllowGet);
        }

       // [AcceptVerbs(HttpVerbs.Get)]
        public JsonResult LoadSubtaskValues(string apk)
        {
           var subTasks = _dashBoardService.GetAllProjectTypes(UserId, apk);
            var taskIds = subTasks.Select(m => new SelectListItem()
            {
                Text = m.Text,
                Value = m.Value.ToString(),
            });
            return Json(taskIds, JsonRequestBehavior.AllowGet);
        }

    }
}