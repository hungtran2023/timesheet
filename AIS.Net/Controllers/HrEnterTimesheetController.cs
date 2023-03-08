using System.Web.Mvc;
using AIS.Models;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.Event;
using AIS.Domain.Common.Constants;
using AIS.Domain.TimeSheet;
using AIS.Domain.HREmployee;
using System.Collections.Generic;
using System.Linq;
using AIS.Data.Model;
using System;
using AIS.Domain.Project;

namespace AIS.Controllers
{
    public class HREnterTimeSheetController : MenuBaseController
    {
        private readonly IEventsService eventsService = Inject.Service<IEventsService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();
        private readonly IHREmployeeService hREmployeeService = Inject.Service<IHREmployeeService>();
        private readonly IProjectService _projectService = Inject.Service<IProjectService>();
      //  GET: /HREnterTimeSheet/HREnterTimeSheet/252
        [HttpGet]
        public ActionResult HREnterTimeSheet(int id)
        {
            var employee = employeeService.FindById(id);
            var queryData = string.Format("{0}?id={1}", StringConstants.TimesheetURL, id);
            var currentWorkingHours = employeeService.CurrentWorkingHours(id);
            TimeSheet model = new TimeSheet();
            model.Hours = currentWorkingHours;
            ViewBag.EmployeeId = id;
            ViewBag.TimesheetUrl = queryData;
            ViewBag.WorkingHours = currentWorkingHours;
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            ViewBag.TimesheetListPageUrl = StringConstants.TimeSheetListURL;
            ViewBag.NameOfEmployee = employeeService.GetFullName(employee);
            ViewBag.AbsenceTypeList = eventsService.GetEventsList();
            return View(model);
        }

        // POST: /HREnterTimeSheet/Add
        [HttpPost]
        public ActionResult Add(TimeSheet timesheet)
        {
            TimeSheet temp = new TimeSheet();
            TimeSheetModel item = temp.MapToTimeSheetDTO(timesheet);
            var result = serviceHandler.AddTimesheetByHr(item);
            if (result == "")
            {
                return AjaxJsonResult(StringConstants.TimesheetAddMessage, null, true);
            }
            else
            {
                return AjaxJsonResult(result, null, false);
            }
        }

        public JsonResult GetFullNameAutoComplete(string term = "")
        {
            var employees = hREmployeeService.GetEmployees();
            var objAPK = employees
                            .Where(c => c.Fullname.ToUpper()
                            .Contains(term.ToUpper()))
                            .Select(c => new { Name = c.Fullname, statffID = c.PersonID})
                            .Distinct().ToList();
            return Json(objAPK, JsonRequestBehavior.AllowGet);
        }
        public ActionResult ResetTimeSheet()
        {
            ViewBag.PageSize = 100;
            var listemployee = GetChooseEmployeeList();
            ViewBag.EmployeeList = listemployee;
            TimeSheetResetModel model = new TimeSheetResetModel();
           
            return View(model);
        }
        [HttpGet]
        public JsonResult GetTimeSheetCleanUpData()
        {

            var lstCleanUpTimeSheet = _projectService.GetDataCleanUpTimeSheet();
            List<AIS.Data.Model.TimeSheetResetModel> list = new List<Data.Model.TimeSheetResetModel>();
            list = lstCleanUpTimeSheet.ToList();
            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;
        }
        [HttpPost]
        public JsonResult DoResetTimeSheet()
        {

            var lstCleanUpTimeSheets = _projectService.GetDataCleanUpTimeSheet();

            foreach (var item in lstCleanUpTimeSheets)
            {
                var flag = _projectService.DoResetTimeSheet(item.TDate, item.StaffID, item.AssignmentID,item.EventID);
            }          
            
            List<AIS.Data.Model.TimeSheetResetModel> list = new List<Data.Model.TimeSheetResetModel>();
            list = lstCleanUpTimeSheets.ToList();
            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;
        }
        public  List<SelectListItem> GetChooseEmployeeList()
        {
          
          
            List<SelectListItem> listemloyees = new List<SelectListItem>();
            var listemployee = hREmployeeService.GetEmployees();
            listemloyees.Add(new SelectListItem() { Value = "", Text = "" });
            foreach (var item in listemployee)
            {
                listemloyees.Add(new SelectListItem() { Value = item.PersonID.ToString(), Text = item.Fullname });
            }

            return listemloyees;
           
        }
    }
}