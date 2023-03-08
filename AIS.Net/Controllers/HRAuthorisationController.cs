using AIS.Domain;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;
using AIS.Domain.Department;
using AIS.Domain.Email;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.Event;
using AIS.Domain.HREmployee;
using AIS.Domain.HRReport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace AIS.Controllers
{
    public class HrAuthorisationController : MenuBaseController
    {
        private readonly IAbsenceRequestService _absenceRequestService = Inject.Service<IAbsenceRequestService>();
        private readonly IDepartmentService _departmentService = Inject.Service<IDepartmentService>();
        private readonly IHRReceiveReportService _hrReceiveReportService = Inject.Service<IHRReceiveReportService>();
        private readonly IEmailService _emailService = Inject.Service<IEmailService>();
        private readonly IEventsService _eventsService = Inject.Service<IEventsService>();
        private readonly IEmailTemplateService _emailTemplateService = Inject.Service<IEmailTemplateService>();
        private readonly IHREmployeeService _hrEmployeeService = Inject.Service<IHREmployeeService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();

        private IEnumerable<Object> ListRequests(int status, string name, string department)
        {
           return serviceHandler.GetListRequestsForHr(status, name, department);
        }

        public ActionResult HRAuthorisation()
        {
            ViewBag.PageSize = PageSize;
            ViewBag.DepartmentList = _departmentService.GetDepartmentList();
            ViewBag.StatusList = ListItemHelper.GetStatusList();
            ViewBag.AbsenceTypeList = _eventsService.GetEventsList();
            ViewBag.GetAuthoriserList = _hrReceiveReportService.GetAuthoriserList(UserId);
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            return View();
        }
        // GET: /HRAuthorisation/GetRequestList
        [HttpGet]
        public ActionResult GetRequestList(int? status , string name,string department)
        {
            status =  status ?? (int)AbsenceStatus.UnAuthorised;
            return Json(ListRequests((int)status, name, department), JsonRequestBehavior.AllowGet);
        }

        // POST: /HRAuthorisation/ApproveRequest
        [HttpPost]
        public ActionResult ApproveRequest(int id, int type, string note, int status,string name,string department)
        {
            var request = _absenceRequestService.FindById(id);
            var authoriseMessage = serviceHandler.AuthorizeRequestByHR(id, note, type, UserId);
            if (authoriseMessage == "")
            {
                EmailModel model = new EmailModel()
                {
                    ManagerId = UserId,
                    RequesterId = request.StaffId,
                    Note = note,
                    DateFrom = request.DateFrom,
                    DateTo = request.DateTo,
                    CC = request.Authoriser1_Id,
                    BCC = request.Authoriser2_Id,
                    Template = StringConstants.EmailInformApproveByAnotherAuthorizor
                };
                _emailService.SendMailForAbsenceRequest(model);
                return AjaxJsonResult(StringConstants.RequestApproveMessage, ListRequests(status, name, department), true);
            }
            else
            {
                return AjaxJsonResult(authoriseMessage, ListRequests(status, name, department), false);
            }
        }

        // POST: /HRAuthorisation/RejectRequest
        [HttpPost]
        public ActionResult RejectRequest(int id, int type,string note , int status, string name, string department)
        {
            serviceHandler.RejectRequestByHR(id, note, UserId, type);
            return AjaxJsonResult(StringConstants.RequestRejectMessage, ListRequests(status, name, department), true);
        }
    }
}