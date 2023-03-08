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
using AIS.Domain.Holiday;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;
using AIS.Domain.Email;

namespace AIS.Controllers
{

    public class HolidayBookingController : BaseController
    {
        private readonly IAbsenceRequestService _absenceRequestService = Inject.Service<IAbsenceRequestService>();
        private readonly IEmailService _emailService = Inject.Service<IEmailService>();
        private readonly IHRReceiveReportService _hrReceiveReportService = Inject.Service<IHRReceiveReportService>();
        private readonly IAnualLeaveDaysService _anualLeaveDaysService = Inject.Service<IAnualLeaveDaysService>();
        private readonly IEventsService _eventsService = Inject.Service<IEventsService>();
        private readonly IEmailTemplateService _emailTemplateService = Inject.Service<IEmailTemplateService>();
        private readonly IHREmployeeService _hrEmployeeService = Inject.Service<IHREmployeeService>();
        private readonly IHolidayService _holidaysService = Inject.Service<IHolidayService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();
        private List<Object> ListRequests
        {
            get
            {
                return serviceHandler.GetListRequests(UserId).ToList();
            }
        }

        public ActionResult HolidayBooking()
        {
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            ViewBag.PageSize = PageSize;
            ViewBag.FullNameOfUser = UserFullName;
            ViewBag.AbsenceTypeList = _eventsService.GetEventsList();
            ViewBag.GetAuthoriserList = _hrReceiveReportService.GetAuthoriserList(UserId);
            return View();
        }

        // Get: /HolidayBooking/GetRequestList
        [HttpGet]
        public ActionResult GetRequestList()
        {
            
            return Json(ListRequests, JsonRequestBehavior.AllowGet);
        }

        // POST: /HolidayBooking/Add
        [HttpPost]
        public ActionResult Add(AbsenseRequest request)
        {
            var balance = _anualLeaveDaysService.GetBalanceHours();
            request.isAuthorizedByHr = request.AbsenceType == (int)AbsenceType.AnnualHoliday ? false : true;
            var requestToAdd = request.ToModel(serviceHandler, balance, UserId);
            var isTimesheetAvailableMessage = serviceHandler.IsTimesheetAvailableForRequest(requestToAdd);
            if (isTimesheetAvailableMessage != String.Empty)
            {
                return  AjaxJsonResult(isTimesheetAvailableMessage, ListRequests, false);
            }
            var currentRequest = _absenceRequestService.Add(requestToAdd);
            if (currentRequest.Id != 0)
            {
                var emailTemplate = _emailTemplateService.GetEmailTemplateByType(StringConstants.EmailRequestApproval);
                var content = emailTemplate.Content;
                EmailModel model = new EmailModel()
                {
                    ManagerId = request.FirstAuthoriserId,
                    RequesterId = UserId,
                    Note = request.Note,
                    LinkAprroveRequestForManager = StringConstants.AuthoriserRedirectInEmailURL,
                    Template = StringConstants.EmailRequestApproval
                };
                _emailService.SendMailForAuthoriser(model);
                return AjaxJsonResult(StringConstants.RequestAddMessage, ListRequests, true);
            }
            return AjaxJsonResult(StringConstants.ErrorMessage, ListRequests, false);
        }

        // POST: /HolidayBooking/Update
        [HttpPost]
        public ActionResult Update(AbsenseRequest request)
        {
            var balance = _anualLeaveDaysService.GetBalanceHours();
            request.isAuthorizedByHr = request.AbsenceType == (int)AbsenceType.AnnualHoliday ? false : true;
            var requestToUpdate = request.ToModel(serviceHandler, balance, UserId);
            

            var isTimesheetAvailableMessage = serviceHandler.IsTimesheetAvailableForRequest(requestToUpdate);
            if (isTimesheetAvailableMessage != String.Empty)
            {
                return AjaxJsonResult(isTimesheetAvailableMessage, ListRequests, false);
            }
            else
            {
                try
                {
                    _absenceRequestService.Update(request.ToModel(serviceHandler, balance, UserId));
                    EmailModel model = new EmailModel()
                    {
                        ManagerId = request.FirstAuthoriserId,
                        RequesterId = UserId,
                        Note = request.Note,
                        LinkAprroveRequestForManager = StringConstants.AuthoriserRedirectInEmailURL,
                        Template = StringConstants.EmailRequestApproval
                    };
                    _emailService.SendMailForAuthoriser(model);
                    ModelState.Clear();
                    return AjaxJsonResult(StringConstants.RequestUpdateMessage, ListRequests, true);
                }
                catch (Exception ex)
                {

                    throw;
                }
               
            }
        }

        // POST: /HolidayBooking/Delete
        [HttpPost]
        public ActionResult Delete(String listOfRequestId)
        {
            String[] IdSplitter = listOfRequestId.TrimEnd(',').Split(',');
            int[] idArray = new int[IdSplitter.Count()];
            foreach (var item in IdSplitter)
            {
                idArray[Array.FindIndex(IdSplitter, i => i == item)] = Convert.ToInt32(item);
            }
            _absenceRequestService.DeleteAll(idArray);
            return AjaxJsonResult(StringConstants.RequestDeleteMessage, ListRequests, true);
        }
    }
}