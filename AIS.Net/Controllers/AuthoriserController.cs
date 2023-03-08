using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.TimeSheet;
using AIS.Domain.HREmployee;
using AIS.Domain.Common.Helper;
using AIS.Domain.Common.Constants;

namespace AIS.Controllers
{
    public class AuthoriserController : BaseController
    {
        private readonly IEmailService _emailService = Inject.Service<IEmailService>();
        private readonly IAbsenceRequestService _absenceRequestService = Inject.Service<IAbsenceRequestService>();
        private readonly ITimeSheetService _timeSheet = Inject.Service<ITimeSheetService>();
        private readonly IEmailTemplateService _emailTemplateService = Inject.Service<IEmailTemplateService>();
        private readonly IHREmployeeService _hrEmployeeService = Inject.Service<IHREmployeeService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();
        private List<Object> OffRequestsOfStaffs
        {
            get
            {
                try
                {
                    return serviceHandler.ListRequestsForAuthoriser(UserId).ToList();
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        public ActionResult Authoriser()
        {
            ViewBag.PageSize = PageSize;
            ViewBag.Manager = UserFullName;
            ViewBag.ListOfMonths = ListItemHelper.Months();
            ViewBag.ListOfYears = ListItemHelper.Years();
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            return View();
        }

        // GET: /Authoriser/GetRequestsOfStaffs
        [HttpGet]
        public ActionResult GetRequestsOfStaffs()
        {
            return Json(OffRequestsOfStaffs, JsonRequestBehavior.AllowGet);
        }

        // POST: /Authoriser/GetDataForTeamCalendar
        [HttpPost]
        public ActionResult GetDataForTeamCalendar(int month, int year)
        {
            ViewBag.DaysInSelectMonth = DateTime.DaysInMonth(year, month);
            try
            {
                return Json(serviceHandler.GetRequestsForManager(UserId, month, year), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(null, JsonRequestBehavior.AllowGet);
            }
        }

        // POST: /Authoriser/DoRejectRequests
        [HttpPost]
        public ActionResult RejectRequests(List<int> RequestIds, String Note)
        {
            serviceHandler.RejectRequests(RequestIds, Note, UserId);
            return AjaxJsonResult(StringConstants.RequestAddMessage, OffRequestsOfStaffs, true);
        }

        // POST: /Authoriser/DoRejectRequest
        [HttpPost]
        public ActionResult RejectRequest(List<int> RequestIds, String Note)
        {
            serviceHandler.RejectRequests(RequestIds, Note, UserId);
            return AjaxJsonResult(StringConstants.RequestAddMessage, OffRequestsOfStaffs, true);
        }

        // POST: /Authoriser/DoApproveRequests
        [HttpPost]
        public ActionResult ApproveRequests(List<int> RequestIds, String Note)
        {
            var message = String.Empty;
            foreach (var item in RequestIds)
            {
                var request = _absenceRequestService.FindById(item);
                if (!request.isAuthoriser1Approved)
                {
                    message += serviceHandler.SendMailAndApprove1(request, Note, UserId);
                }
                else
                {
                    message += serviceHandler.SendMailAndApprove2(request, Note, UserId);
                }
            }
            if (message == String.Empty)
            {
                return AjaxJsonResult(StringConstants.RequestApproveMessage, OffRequestsOfStaffs, true);
            }
            else
            {
                return AjaxJsonResult(message, OffRequestsOfStaffs, false);
            }
        }

        // POST: /Authoriser/DoApproveRequest
        [HttpPost]
        public ActionResult ApproveRequest(List<int> RequestIds, String Note)
        {
            var requestId = RequestIds.First();
            var request = _absenceRequestService.FindById(requestId);
            var result = "";
            if (!request.isAuthoriser1Approved)
            {
                result = serviceHandler.SendMailAndApprove1(request, Note,UserId);
            }
            else
            {
                result = serviceHandler.SendMailAndApprove2(request, Note,UserId);
            }
            if (result == String.Empty)
            {
                return AjaxJsonResult(StringConstants.RequestApproveMessage, OffRequestsOfStaffs, true);
            }
            else
            {
                return AjaxJsonResult(result, OffRequestsOfStaffs, false);
            }
        }

    }
}