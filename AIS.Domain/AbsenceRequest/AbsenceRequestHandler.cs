using AIS.Data;
using AIS.Data.Model;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;
using AIS.Domain.Email;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.Employee;
using AIS.Domain.Holiday;
using AIS.Domain.HREmployee;
using AIS.Domain.TimeSheet;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Domain.AbsenceRequest
{
    public class AbsenceRequestHandler
    {
        private readonly IEmployeeService _employeeService;
        private readonly IAbsenceRequestService _absenceRequestService;
        private readonly ITimeSheetService _timesheetService;
        private readonly IEmailService _emailService;
        private readonly IHolidayService _holidayService;
        private readonly IHREmployeeService _hrEmployeeService;
        DateTime[] holidaysArray;
        TimeSpan endWorkTime = new TimeSpan(17, 30, 0);
        TimeSpan startWorkTime = new TimeSpan(8, 0, 0);
        public AbsenceRequestHandler(IEmployeeService employeeService,
        IAbsenceRequestService absenceRequestService,
        ITimeSheetService timesheetService, IEmailService emailService, IHolidayService holidayService, IHREmployeeService hrEmployeeService)
        {
            var nextYear = DateTime.Now.Year + 1;
            _employeeService = employeeService;
            _absenceRequestService = absenceRequestService;
            _timesheetService = timesheetService;
            _holidayService = holidayService;
            _hrEmployeeService = hrEmployeeService;
            _emailService = emailService;
            holidaysArray = _holidayService.GetAll();

        }
        public int IsTimesheetAvailableThisYear(DateTime date, int staffId, double Hours, int Type, List<ATC_Timesheet> currentYearTimesheet)
        {
            var workingTime = _employeeService.GetWorkingTime(date, staffId);
            var timesheet = currentYearTimesheet.Where(t => t.TDate.Date == date.Date && t.StaffID == staffId);
            if (timesheet.Count() > 0)
            {
                if (timesheet.Where(t => t.EventID == Type).Count() > 0)
                {
                    return NumberConstants.Error;
                }
                double workingHours = 0;
                foreach (var item in timesheet)
                {
                    workingHours += (double)item.Hours;
                }
                workingHours += Hours;
                if (workingHours > workingTime.ToDouble())
                {
                    if (timesheet.First().EventID == NumberConstants.Holiday)
                    {
                        return NumberConstants.Ignore;
                    }
                    return NumberConstants.Error;
                }
            }
            return Hours <= workingTime.ToDouble() ? NumberConstants.Success : NumberConstants.Error;
        }

        public int IsTimesheetAvailableForNextYear(DateTime date, int staffId, double Hours, int Type, IEnumerable<dynamic> nextYearTimesheet)
        {
            var timesheet = nextYearTimesheet.Where(t => t.TDate.Date == date.Date && t.StaffID == staffId);
            var workingTime = _employeeService.GetWorkingTime(date, staffId);
            if (timesheet.Count() > 0)
            {
                if (timesheet.Where(t => t.EventID == Type).Count() > 0)
                {
                    return NumberConstants.Error;
                }
                double workingHours = 0;
                foreach (var item in timesheet)
                {
                    workingHours += (double)item.Hours;
                }
                workingHours += Hours;
                if (workingHours > workingTime.ToDouble())
                {
                    if (timesheet.First().EventID == NumberConstants.Holiday)
                    {
                        return NumberConstants.Ignore;
                    }
                    return NumberConstants.Error;
                }
            }
            return Hours <= workingTime.ToDouble() ? NumberConstants.Success : NumberConstants.Error;
        }
        

        public List<ATC_Timesheet> IsThisYearTimesheetFailAdd(TimeSheetModel model, Dictionary<DateTime, double> listOfDateAndTimeOff, out DateTime errorDate)
        {
            errorDate = DateTime.Now;
            var listOfTimesheet = new List<ATC_Timesheet>();
            var listThisYearDateTimeOff = listOfDateAndTimeOff.Where(t => t.Key.Year == DateTime.Now.Year);
            var currentYearTimesheet = _timesheetService.FindAll().Where(t => t.TDate.Date >= model.DateFrom.Date && t.TDate.Date <= model.DateTo.Date).ToList();
            foreach (var item in listThisYearDateTimeOff)
            {
                var timesheetChecker = IsTimesheetAvailableThisYear(item.Key, model.StaffId, item.Value, model.AbsenceType, currentYearTimesheet);
                if (timesheetChecker == NumberConstants.Ignore)
                {
                    continue;
                }
                else if (timesheetChecker == NumberConstants.Success)
                {
                    var timesheet = new ATC_Timesheet();
                    timesheet.TDate = item.Key;
                    timesheet.StaffID = model.StaffId;
                    timesheet.AssignmentID = 1;
                    timesheet.EventID = model.AbsenceType;
                    timesheet.Hours = (decimal)item.Value;
                    timesheet.Note = model.Note;
                    timesheet.OTNight = 0;
                    timesheet.OTNormal = 0;
                    timesheet.OverRate = 0;
                    timesheet.OverTime = 0;
                    listOfTimesheet.Add(timesheet);
                }
                else
                {
                    errorDate = item.Key;
                    return null;
                }
            }
            return listOfTimesheet;
        }

        public List<dynamic> IsLastYearTimesheetFailAdd(TimeSheetModel model, Dictionary<DateTime, double> listOfDateAndTimeOff, out DateTime dateFullTimesheet)
        {
            dateFullTimesheet = DateTime.Now;
            List<dynamic> listOfTimesheet = new List<dynamic>();
            var lastYear = DateTime.Now.Year - 1;
            var listNextYearDateTimeOff = listOfDateAndTimeOff.Where(t => t.Key.Year == lastYear);
            var lastYearTimesheet = ((IEnumerable<dynamic>)_timesheetService.GetListByType(StringConstants.TableTimesheet + lastYear)).Where(t => t.TDate.Date >= model.DateFrom.Date && t.TDate.Date <= model.DateTo.Date).ToList();
            foreach (var item in listNextYearDateTimeOff)
            {
                var timesheetChecker = IsTimesheetAvailableForNextYear(item.Key, model.StaffId, item.Value, model.AbsenceType, lastYearTimesheet);
                if (timesheetChecker == NumberConstants.Ignore)
                {
                    continue;
                }
                else if (timesheetChecker == NumberConstants.Success)
                {
                    dynamic timesheet = typeof(Entity).Assembly.CreateInstance(StringConstants.TableTimesheet + lastYear);
                    timesheet.TDate = item.Key;
                    timesheet.StaffID = model.StaffId;
                    timesheet.AssignmentID = 1;
                    timesheet.EventID = model.AbsenceType;
                    timesheet.Hours = (decimal)item.Value;
                    timesheet.Note = model.Note;
                    timesheet.OTNight = 0;
                    timesheet.OTNormal = 0;
                    timesheet.OverRate = 0;
                    timesheet.OverTime = 0;
                    listOfTimesheet.Add(timesheet);
                }
                else
                {
                    dateFullTimesheet = item.Key;
                    return null;
                }
            }
            return listOfTimesheet;
        }

        public List<dynamic> IsNextYearTimesheetFailAdd(TimeSheetModel model, Dictionary<DateTime, double> listOfDateAndTimeOff, out DateTime dateFullTimesheet)
        {
            dateFullTimesheet = DateTime.Now;
            List<dynamic> listOfTimesheet = new List<dynamic>();
            var nextYear = DateTime.Now.Year + 1;
            var listNextYearDateTimeOff = listOfDateAndTimeOff.Where(t => t.Key.Year == nextYear);
            var nextYearTimesheet = ((IEnumerable<dynamic>)_timesheetService.GetListByType(StringConstants.TableTimesheet + nextYear)).Where(t => t.TDate.Date >= model.DateFrom.Date && t.TDate.Date <= model.DateTo.Date).ToList();
            foreach (var item in listNextYearDateTimeOff)
            {
                var timesheetChecker = IsTimesheetAvailableForNextYear(item.Key, model.StaffId, item.Value, model.AbsenceType, nextYearTimesheet);
                if (timesheetChecker == NumberConstants.Ignore)
                {
                    continue;
                }
                else if (timesheetChecker == NumberConstants.Success)
                {
                    dynamic timesheet = typeof(Entity).Assembly.CreateInstance(StringConstants.TableTimesheet + nextYear);
                    timesheet.TDate = item.Key;
                    timesheet.StaffID = model.StaffId;
                    timesheet.AssignmentID = 1;
                    timesheet.EventID = model.AbsenceType;
                    timesheet.Hours = (decimal)item.Value;
                    timesheet.Note = model.Note;
                    timesheet.OTNight = 0;
                    timesheet.OTNormal = 0;
                    timesheet.OverRate = 0;
                    timesheet.OverTime = 0;
                    listOfTimesheet.Add(timesheet);
                }
                else
                {
                    dateFullTimesheet = item.Key;
                    return null;
                }
            }
            return listOfTimesheet;
        }

        public DateTime? InsertTimesheet(TimeSheetModel model, Dictionary<DateTime, double> listOfDateAndTimeOff)
        {
            var nextYear = DateTime.Now.Year + 1;
            var lastYear = DateTime.Now.Year - 1;

            if (model.DateFrom.Year == model.DateTo.Year && model.DateFrom.Year == DateTime.Now.Year)
            {
                DateTime dateFullTimesheet;
                var listOfCurrentYearTimesheet = IsThisYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfCurrentYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                _timesheetService.AddAll(listOfCurrentYearTimesheet.ToArray());

            }
            else if (model.DateFrom.Year == model.DateTo.Year && model.DateFrom.Year == lastYear)
            {
                DateTime dateFullTimesheet;
                var listOfNextYearTimesheet = IsLastYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfNextYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                _timesheetService.AddByType(listOfNextYearTimesheet.ToArray());

            }
            else if (model.DateFrom.Year == model.DateTo.Year && model.DateFrom.Year == nextYear)
            {
                DateTime dateFullTimesheet;
                var listOfNextYearTimesheet = IsNextYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfNextYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                _timesheetService.AddByType(listOfNextYearTimesheet.ToArray());

            }
            else if (model.DateFrom.Year == lastYear && model.DateTo.Year == DateTime.Now.Year)
            {
                DateTime dateFullTimesheet;
                var listOfCurrentYearTimesheet = IsThisYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfCurrentYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                var listOfNextYearTimesheet = IsLastYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfNextYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                _timesheetService.AddAll(listOfCurrentYearTimesheet.ToArray());
                _timesheetService.AddByType(listOfNextYearTimesheet.ToArray());
            }
            else if (model.DateFrom.Year == DateTime.Now.Year && model.DateTo.Year == nextYear)
            {
                DateTime dateFullTimesheet;
                var listOfCurrentYearTimesheet = IsThisYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfCurrentYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                var listOfNextYearTimesheet = IsNextYearTimesheetFailAdd(model, listOfDateAndTimeOff, out dateFullTimesheet);
                if (listOfNextYearTimesheet == null)
                {
                    return dateFullTimesheet;
                }
                _timesheetService.AddAll(listOfCurrentYearTimesheet.ToArray());
                _timesheetService.AddByType(listOfNextYearTimesheet.ToArray());
            }
            return null;
        }

        public double GetTotalUnauthorisedDays(int staffId)
        {
            var getToTalOffRequest = _absenceRequestService.FindAll().Where(
                    request => (request.Status == (int)AbsenceStatus.New
                    || request.Status == (int)AbsenceStatus.InProgress)
                    && request.StaffId == staffId);
            if (getToTalOffRequest.Count() != 0)
            {
                double totalUnauthorisedHours = 0;
                foreach (var item in getToTalOffRequest)
                {
                    totalUnauthorisedHours += TotalWorkingHours(item.DateFrom, item.DateTo, staffId);
                }
                return totalUnauthorisedHours;
            }
            return 0;
        }

        public double TotalWorkingHours(DateTime StartDate, DateTime EndDate, int StaffId)
        {
            double result = 0;
            var listOffHours = GetListOfOffHours(StartDate, EndDate, StaffId);
            foreach (var item in listOffHours)
            {
                result += item.Value;
            }
            return result;
        }

        public String CheckValidForTimesheet(ATC_AbsenceRequests request)
        {
            var staff = _employeeService.FindById(request.StaffId);
            var leaveDate = staff.LeaveDate;
            var isNoWorkday = DateTimeHelper.BusinessDaysUntil(request.DateFrom, request.DateTo, holidaysArray) == 0;
            if (isNoWorkday)
            {
                return StringConstants.RequestErrorNoWorkDays;
            }
            if (leaveDate != null && request.DateTo > leaveDate.Value)
            {
                return StringConstants.RequestErrorGreaterLeaveDate;
            }
            if (request.Id == 0 && _absenceRequestService.IsRequestMade(request.DateFrom.Date, request.StaffId))
            {
                return String.Format(StringConstants.RequestErrorAlreadyMade, request.DateFrom.Date.ToString(StringConstants.DateOnlyFormat));
            }
            if (request.Id != 0 && _absenceRequestService.IsRequestMade(request.DateFrom.Date, request.StaffId, request.Id))
            {
                return String.Format(StringConstants.RequestErrorAlreadyMade, request.DateFrom.Date.ToString(StringConstants.DateOnlyFormat));
            }
            return string.Empty;
        }

        public void RejectRequests(List<int> RequestIds, String Note, int UserId)
        {
            var listOfRequests = _absenceRequestService.FindAll().Where(requests => RequestIds.Contains(requests.Id));
            var listUpdatedRequest = new List<ATC_AbsenceRequests>();
            foreach (var item in listOfRequests)
            {
                item.Status = (int)AbsenceStatus.Rejected;
                if (item.Authoriser1_Id == UserId)
                {
                    EmailModel model = new EmailModel()
                    {
                        ManagerId = (int)item.Authoriser1_Id,
                        RequesterId = item.StaffId,
                        Note = Note,
                        DateFrom = item.DateFrom,
                        DateTo = item.DateTo,
                        Template = StringConstants.EmailReject,
                    };
                    _emailService.SendMailForAbsenceRequest(model);
                }
                if (item.Authoriser2_Id == UserId)
                {
                    EmailModel model = new EmailModel()
                    {
                        ManagerId = (int)item.Authoriser2_Id,
                        RequesterId = item.StaffId,
                        Note = Note,
                        DateFrom = item.DateFrom,
                        DateTo = item.DateTo,
                        Template = StringConstants.EmailReject,
                        CC = item.Authoriser1_Id,
                    };
                    _emailService.SendMailForAbsenceRequest(model);
                }
                listUpdatedRequest.Add(item);
            }
            _absenceRequestService.UpdateAll(listUpdatedRequest);
        }

        public void RejectRequestByHR(int id, String Note, int UserId, int type)
        {
            var request = _absenceRequestService.FindAll().Where(requests => requests.Id == id).First();
            request.Status = (int)AbsenceStatus.Rejected;
            request.HrNote = Note;
            request.Type = type;
            _absenceRequestService.Update(request);
            EmailModel model = new EmailModel()
            {
                ManagerId = UserId,
                RequesterId = request.StaffId,
                Note = Note,
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                Template = StringConstants.EmailReject,
                CC = request.Authoriser1_Id,
                BCC = request.Authoriser2_Id
            };
            _emailService.SendMailForAbsenceRequest(model);
        }

        public IEnumerable<Object> GetListRequests(int staffId)
        {
            var currentYear = DateTime.Now.Year;
            var eagerList = _absenceRequestService.ListRequestView().ToList();
            var listdata = from t in eagerList.Where(t =>
                          t.StaffId == staffId
                          &&
                          ((t.DateTo.Year < currentYear && t.Status != (int)AbsenceStatus.Authorised && t.Status != (int)AbsenceStatus.Rejected)
                          ||
                          (t.DateTo.Year >= currentYear))).Reverse().ToList()
                           select new
                           {
                               Id = t.Id,
                               Category = t.ATC_Events.EventName,
                               FirstDay = t.DateFrom.ToString(StringConstants.DateTimeFormat),
                               LastDay = t.DateTo.ToString(StringConstants.DateTimeFormat),
                               Total = GetDayOff(t.DateFrom, t.DateTo, t.StaffId),
                               Status = AbsenceTypeHelper.ConvertAbsenceStatus((AbsenceStatus)t.Status, t.DateTo),
                               AbsenceType = t.Type,
                               StartDate = t.DateFrom.ToString(StringConstants.DateOnlyFormat),
                               EndDate = t.DateTo.ToString(StringConstants.DateOnlyFormat),
                               StartTime = t.DateFrom.ToString(StringConstants.HourOnlyFormat),
                               EndTime = t.DateTo.ToString(StringConstants.HourOnlyFormat),
                               FirstAuthoriserId = t.Authoriser1_Id,
                               SecondAuthoriserId = t.Authoriser2_Id,
                               Note = t.Note,
                               isAuthorizedByHr = t.isAuthorisedByHr
                           };
            return listdata;
        }

        public IEnumerable<Object> ListRequestsForAuthoriser(int managerId)
        {
            var eagerList = _absenceRequestService.ListRequestView().ToList();
            var tempRequests = eagerList.Where(
                request =>
                ((request.Authoriser1_Id == managerId && (request.Status == (int)AbsenceStatus.New))
                || (request.Status == (int)AbsenceStatus.InProgress && request.Authoriser2_Id == managerId && request.isAuthoriser1Approved && request.isAuthoriser2Approved != true))).Reverse().ToList();
            return from requests in tempRequests
                   select new
                   {
                       Id = requests.Id,
                       FullName = _employeeService.GetFullName(requests.StaffId),
                       Type = requests.ATC_Events.EventName,
                       DateFrom = requests.DateFrom.ToString(StringConstants.DateTimeFormat),
                       DateTo = requests.DateTo.ToString(StringConstants.DateTimeFormat),
                       Total = GetDayOff(requests.DateFrom, requests.DateTo, requests.StaffId),
                       Note = requests.Note
                   };
        }

        public IEnumerable<Object> GetListRequestsForHr(int status, string name = "", string departmentId = "")
        {
            var index = 0;
            
            var eagerList = _absenceRequestService.ListRequestView().Where(t => t.isAuthorisedByHr == true);
            if (!string.IsNullOrEmpty(name))
            {
                name = name.ToLower();
                eagerList = eagerList.Where(t => (t.Staff.PersonalInfo.FirstName + " " + t.Staff.PersonalInfo.LastName).ToLower().Contains(name));
            }

            if (!string.IsNullOrEmpty(departmentId))
            {
                try
                {
                    eagerList = eagerList.Where(t => t.Staff.DepartmentID == int.Parse(departmentId));
                }
                catch (Exception)
                {
                    return null;
                }
            }
            var query = new List<ATC_AbsenceRequests>();

            if (status == -1)
            {
                query = eagerList.Where(t =>
                              (t.isAuthoriser1Approved && t.Authoriser2_Id == null && t.isHrApproved != true && t.Status != (int)AbsenceStatus.Rejected)
                            || (t.isAuthoriser2Approved == true && t.isHrApproved != true)
                            || t.Status == (int)AbsenceStatus.Rejected || t.Status == (int)AbsenceStatus.Authorised)
                            .Reverse().OrderBy(t => t.Status).ToList();
            }
            else
            {
                if (!Enum.IsDefined(typeof(AbsenceStatus), status))
                {
                    return null;
                }

                var enumStatus = (AbsenceStatus)status;

                if (AbsenceStatus.UnAuthorised == enumStatus)
                {
                    query = eagerList.Where(t =>
                    t.Status != (int)AbsenceStatus.Rejected &&
                             ((t.isAuthoriser1Approved && t.Authoriser2_Id == null && t.isHrApproved != true)
                             || (t.isAuthoriser2Approved == true && t.isHrApproved != true))
                             ).Reverse().ToList();

                }
                else
                if (AbsenceStatus.Authorised == enumStatus)
                {
                    query = eagerList.Where(t => t.Status == (int)AbsenceStatus.Authorised && t.DateTo > DateTime.Now).Reverse().ToList();
                }
                else
                if (AbsenceStatus.Rejected == enumStatus)
                {
                    query = eagerList.Where(t => t.Status == (int)AbsenceStatus.Rejected).Reverse().ToList();
                }
                else
                if (AbsenceStatus.Taken == enumStatus)
                {
                    query = eagerList.Where(t => t.Status == (int)AbsenceStatus.Authorised && t.DateTo <= DateTime.Now).Reverse().ToList();
                }
            }
            var listdata = from t in query
                           select new
                           {
                               ID = index++,
                               RequestId = t.Id,
                               StaffId = t.StaffId,
                               FullName = _employeeService.GetFullName(t.Staff),
                               FirstDay = t.DateFrom.ToString(StringConstants.DateTimeFormat),
                               LastDay = t.DateTo.ToString(StringConstants.DateTimeFormat),
                               DepartmentId = t.Staff.DepartmentID,
                               Type = t.ATC_Events.EventName,
                               Status = AbsenceTypeHelper.ConvertAbsenceStatus((AbsenceStatus)t.Status, t.DateTo),
                               StatusId = AbsenceTypeHelper.GetStatusId((AbsenceStatus)t.Status, t.DateTo),
                               AbsenceType = t.Type,
                               StartDate = t.DateFrom.ToString(StringConstants.DateOnlyFormat),
                               EndDate = t.DateTo.ToString(StringConstants.DateOnlyFormat),
                               StartTime = t.DateFrom.ToString(StringConstants.HourOnlyFormat),
                               EndTime = t.DateTo.ToString(StringConstants.HourOnlyFormat),
                               FirstAuthoriserId = t.Authoriser1_Id,
                               SecondAuthoriserId = t.Authoriser2_Id,
                               Note = t.Note,
                               isAuthorizedByHr = t.isAuthorisedByHr
                           };
            return listdata;
        }
        private void SendInprogressMailToRequester(ATC_AbsenceRequests request, String Note, int UserId)
        {
            EmailModel model = new EmailModel()
            {
                ManagerId = UserId,
                RequesterId = request.StaffId,
                Note = Note,
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                Template = StringConstants.EmailInformApproveByAuthorizor1
            };
            _emailService.SendMailForAbsenceRequest(model);
        }
        public String SendMailAndApprove1(ATC_AbsenceRequests request, String Note, int UserId)
        {
            EmailModel model;
            if (request.Authoriser2_Id != null)
            {
                SendInprogressMailToRequester(request, Note, UserId);
                model = new EmailModel()
                {
                    ManagerId = (int)request.Authoriser2_Id,
                    RequesterId = request.StaffId,
                    Note = Note,
                    Template = StringConstants.EmailRequestApproval
                };
                _emailService.SendMailForAuthoriser(model);
                _absenceRequestService.InProgressRequest(request.Id, Note, UserId);
            }
            else
            if (request.isAuthorisedByHr)
            {
                SendInprogressMailToRequester(request, Note, UserId);
                _absenceRequestService.InProgressRequest(request.Id, Note, UserId);
            }
            else
            {
                var result = AuthorizeRequest(request.Id, Note, UserId);
                if (result != "")
                {
                    return result;
                }
                model = new EmailModel()
                {
                    ManagerId = UserId,
                    RequesterId = request.StaffId,
                    Note = Note,
                    DateFrom = request.DateFrom,
                    DateTo = request.DateTo,
                    Template = StringConstants.EmailInformApproveByAnotherAuthorizor
                };
                _emailService.SendMailForAbsenceRequest(model);
            }
            request.isAuthoriser1Approved = true;
            _absenceRequestService.Update(request);
            return String.Empty;
        }

        public String SendMailAndApprove2(ATC_AbsenceRequests request, string Note, int UserId)
        {
            var requester = _hrEmployeeService.GetEmployeeInfoById(request.StaffId);
            if (request.isAuthorisedByHr)
            {
                SendInprogressMailToRequester(request, Note, UserId);
            }
            else
            {
                var result = AuthorizeRequest(request.Id, Note, UserId);
                if (result != "")
                {
                    return result;
                }
                EmailModel model = new EmailModel()
                {
                    ManagerId = UserId,
                    RequesterId = request.StaffId,
                    Note = Note,
                    DateFrom = request.DateFrom,
                    DateTo = request.DateTo,
                    CC = request.Authoriser1_Id,
                    Template = StringConstants.EmailInformApproveByAnotherAuthorizor
                };
                _emailService.SendMailForAbsenceRequest(model);
            }
            request.isAuthoriser2Approved = true;
            _absenceRequestService.Update(request);
            return String.Empty;
        }

        public String AuthorizeRequest(int Id, String Note, int ManagerId)
        {
            var result = string.Empty;
            var request = _absenceRequestService.FindById(Id);
            var nextYear = DateTime.Now.Year + 1;
            var model = new TimeSheetModel()
            {
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                AbsenceType = request.Type,
                Note = String.Format(StringConstants.TimesheetAutomaticallyInserted, Id),
                StaffId = request.StaffId
            };
            var listOfDateAndTimeOff = GetListOfOffHours(model.DateFrom, model.DateTo, model.StaffId);
            var isInserted = InsertTimesheet(model, listOfDateAndTimeOff);
            if (isInserted != null)
            {
                return FailedAuthorisationMessage(isInserted.Value, request, ManagerId);
            }
            if (request.Authoriser1_Id == ManagerId)
            {
                request.isAuthoriser1Approved = true;
                request.Authoriser1Note = Note;
            }
            if (request.Authoriser2_Id == ManagerId)
            {
                request.isAuthoriser2Approved = true;
                request.Authoriser2Note = Note;
            }
            request.Status = (int)AbsenceStatus.Authorised;
            _absenceRequestService.Update(request);
            return String.Empty;
        }

        public String AddTimesheetByHr(TimeSheetModel model)
        {
            var result = String.Empty;
            var listUpdatedRequest = new List<ATC_AbsenceRequests>();
            var nextYear = DateTime.Now.Year + 1;
            var listOfDateAndTimeOff = DateTimeHelper.FillHoursAndDate(model.DateFrom, model.DateTo, model.Hours);
            result = this.CheckDateOffInRangeHR(model);
            var isInserted = InsertTimesheet(model, listOfDateAndTimeOff);
            if (isInserted != null)
            {
                return String.Format(StringConstants.TimesheetErrorFull, isInserted.Value.ToString(StringConstants.DateOnlyFormat));
            }
            return result;
        }
        public String AuthorizeRequestByHR(int Id, String Note, int type, int userId)
        {
            var result = string.Empty;
            var request = _absenceRequestService.FindById(Id);
            var model = new TimeSheetModel()
            {
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                AbsenceType = type,
                Note = String.Format(StringConstants.TimesheetAutomaticallyInserted, Id),
                StaffId = request.StaffId
            };
            var listOfDateAndTimeOff = GetListOfOffHours(model.DateFrom, model.DateTo, model.StaffId);
            var isInserted = InsertTimesheet(model, listOfDateAndTimeOff);
            if (isInserted != null)
            {
                return FailedAuthorisationMessageHR(isInserted.Value, request, userId);
            }
            request.Type = type;
            request.isHrApproved = true;
            request.HrNote = Note;
            request.Status = (int)AbsenceStatus.Authorised;
            _absenceRequestService.Update(request);
            return String.Empty;
        }

        public IEnumerable<object> GetRequestsForManager(int managerId, int month, int year)
        {
            String temp = year.ToString();
            var result = new List<TeamCalendar>();
            temp = year == DateTime.Now.Year ? String.Empty : temp;
            var currentYearTimesheet = _timesheetService.ListOfTimesheet(temp).Cast<dynamic>().ToList();
            var listOfAllAbsenceRequest = _absenceRequestService.FindAll().ToList();
            if (currentYearTimesheet != null)
            {
                var listOfStaffs = _employeeService.StaffsForManager(managerId);
                if (listOfStaffs.Count() == 0)
                {
                    while (result.ToArray().Count() < NumberConstants.DefaultNumberOfRowOnTable)
                    {
                        var teamCalendar = new TeamCalendar();
                        result.Add(teamCalendar);
                    }
                    return result;
                }
                foreach (var item in listOfStaffs)
                {
                    var teamCalendar = new TeamCalendar();
                    var listEventNotInclude = new List<int>() { 1, 2, 3 };
                    var timesheetList = currentYearTimesheet
                        .Where(
                            timesheet => timesheet.StaffID == item.StaffID
                          && timesheet.TDate.Month == month
                          && !listEventNotInclude.Contains(timesheet.EventID)
                        );
                    var listOfRequests = listOfAllAbsenceRequest.Where(request =>
                        request.StaffId == item.StaffID
                        && request.Status != (int)AbsenceStatus.Authorised && request.Status != (int)AbsenceStatus.Rejected
                        && request.DateFrom.Month == month && request.DateTo.Month == month && request.DateFrom.Year == year && request.DateTo.Year == year
                    );
                    teamCalendar.FullName = _employeeService.GetFullName(item);
                    foreach (var request in listOfRequests)
                    {
                        var dictionaryOffHours = GetListOfOffHours(request.DateFrom, request.DateTo, request.StaffId);
                        foreach (var dateHours in dictionaryOffHours)
                        {
                            var teamDay = new TeamDay();
                            teamDay.WorkingHours = Math.Round(Convert.ToDecimal(dateHours.Value), 1);
                            teamDay.Status = AbsenceTypeHelper.ConvertAbsenceStatus((AbsenceStatus)request.Status);
                            var day = dateHours.Key.Day;
                            var dateNumber = day.ToString();
                            if (day < 10)
                            {
                                dateNumber = StringConstants.Zero + dateNumber;
                            }
                            typeof(TeamCalendar).GetProperty("Day" + dateNumber).SetValue(teamCalendar, teamDay);
                        }
                    }

                    foreach (var timesheet in timesheetList)
                    {
                        var teamDay = new TeamDay();
                        teamDay.WorkingHours = Math.Round(timesheet.Hours, 1);
                        teamDay.Status = StringConstants.AuthorisedStatus;
                        if (timesheet.TDate < DateTime.Now)
                        {
                            teamDay.Status = StringConstants.TakenStatus;
                        }
                        teamDay.Status = timesheet.EventID == 5 ? "Holiday" : teamDay.Status;
                        var day = timesheet.TDate.Day;
                        var dateNumber = day.ToString();
                        if (day < 10)
                        {
                            dateNumber = StringConstants.Zero + day;
                        }
                        typeof(TeamCalendar).GetProperty("Day" + dateNumber).SetValue(teamCalendar, teamDay);
                    }
                    result.Add(teamCalendar);
                }
                while (result.ToArray().Count() < NumberConstants.DefaultNumberOfRowOnTable)
                {
                    var teamCalendar = new TeamCalendar();
                    result.Add(teamCalendar);
                }
                return result;
            }
            return null;
        }

        public IEnumerable<object> GetRequestAjax(int managerId, int month, int year)
        {
                 
           var name = _employeeService.GetFullName(managerId);
            String temp = year.ToString();
            var result = new List<EmployeeModel>();
            //temp = year == DateTime.Now.Year ? String.Empty : temp;
            //var currentYearTimesheet = _timesheetService.ListOfTimesheet(temp).Cast<dynamic>().ToList();
            //var listOfAllAbsenceRequest = _absenceRequestService.FindAll().ToList();

            if (name != String.Empty)
            {
                EmployeeModel employee = new EmployeeModel();

                employee.FullName = name;
                result.Add(employee);

                return result;
            }

          
            return null;
        }

        public String IsTimesheetAvailableForRequest(ATC_AbsenceRequests request)
        {
            var requestErrorMessage = CheckValidForTimesheet(request);
            if (requestErrorMessage != String.Empty)
            {
                return requestErrorMessage;
            }
            var nextYear = DateTime.Now.Year + 1;
            var dictionaryOffHours = GetListOfOffHours(request.DateFrom, request.DateTo, request.StaffId);
            if (request.DateFrom.Year == request.DateTo.Year && request.DateFrom.Year == DateTime.Now.Year)
            {
                List<ATC_Timesheet> listOfTimesheet = new List<ATC_Timesheet>();
                var currentYearTimesheet = _timesheetService.FindAll().Where(t => t.TDate.Date >= request.DateFrom.Date && t.TDate.Date <= request.DateTo.Date).ToList();
                foreach (var item in dictionaryOffHours)
                {
                    if (IsTimesheetAvailableThisYear(item.Key, request.StaffId, item.Value, request.Type, currentYearTimesheet) == NumberConstants.Error)
                    {
                        return FailedAuthorisationMessageTimesheet(item.Key);
                    }
                }
            }
            if (request.DateFrom.Year == request.DateTo.Year && request.DateFrom.Year == nextYear)
            {
                try
                {
                    List<dynamic> listOfTimesheet = new List<dynamic>();
                    var nextYearTimesheet = ((IEnumerable<dynamic>)_timesheetService.GetListByType(StringConstants.TableTimesheet + nextYear)).Where(t => t.TDate.Date >= request.DateFrom.Date && t.TDate.Date <= request.DateTo.Date).ToList();
                    foreach (var item in dictionaryOffHours)
                    {
                        if (IsTimesheetAvailableForNextYear(item.Key, request.StaffId, item.Value, request.Type, nextYearTimesheet) == NumberConstants.Error)
                        {
                            return FailedAuthorisationMessageTimesheet(item.Key);
                        }
                    }
                }
                catch (Exception)
                {
                    return StringConstants.TimesheetNoTable;
                }
            }
            else
            {
                try
                {
                    var currentYearTimesheet = _timesheetService.FindAll().Where(t => t.TDate.Date >= request.DateFrom.Date && t.TDate.Date <= request.DateTo.Date).ToList();
                    var nextYearTimesheet = ((IEnumerable<dynamic>)_timesheetService.GetListByType(StringConstants.TableTimesheet + nextYear)).Where(t => t.TDate.Date >= request.DateFrom.Date && t.TDate.Date <= request.DateTo.Date).ToList();
                    foreach (var item in dictionaryOffHours)
                    {
                        if (item.Key.Year == DateTime.Now.Year && IsTimesheetAvailableThisYear(item.Key, request.StaffId, item.Value, request.Type, currentYearTimesheet) == NumberConstants.Error)
                        {
                            return FailedAuthorisationMessageTimesheet(item.Key);
                        }
                        if (item.Key.Year != DateTime.Now.Year && IsTimesheetAvailableForNextYear(item.Key, request.StaffId, item.Value, request.Type, nextYearTimesheet) == NumberConstants.Error)
                        {
                            return FailedAuthorisationMessageTimesheet(item.Key);
                        }
                    }
                }
                catch (Exception)
                {
                    return StringConstants.TimesheetNoTable;
                }
            }
            return "";
        }

        public Dictionary<DateTime, Double> GetListOfOffHours(DateTime StartDate, DateTime EndDate, int StaffId)
        {
            var results = new Dictionary<DateTime, Double>();
            if (StartDate.Date == EndDate.Date)
            {
                var endBreakTime = StartDate.Date.AddHours(13);
                var startBreakTime = StartDate.Date.AddHours(12);
                if (!StartDate.isWeekdays())
                {
                    return results;
                }
                if (StartDate > startBreakTime && StartDate < endBreakTime)
                {
                    StartDate = endBreakTime;
                }
                if (EndDate > startBreakTime && EndDate < endBreakTime)
                {
                    EndDate = startBreakTime;
                }
                if (EndDate >= endBreakTime && StartDate <= startBreakTime)
                {
                    EndDate = EndDate.AddHours(-1);
                }
                results.Add(StartDate.Date, (EndDate - StartDate).TotalHours);
                return results;
            }
            var listOfSalaryStatus = _employeeService.GetSalaryStatusBetween(StartDate, EndDate, StaffId).ToList();
            if (StartDate.isWeekdays())
            {
                var startDateWorkingHours = listOfSalaryStatus.First().ATC_WorkingHours.Hours.ToDouble();
                double startWorkDateHours = DateTimeHelper.ConvertMinutesToRoundHours(DateTimeHelper.GetStartDateOffMinutes(StartDate, EndDate, startDateWorkingHours));
                results.Add(StartDate.Date, startWorkDateHours);
            }
            if (EndDate.isWeekdays())
            {
                double endWorkDateHours = DateTimeHelper.ConvertMinutesToRoundHours(DateTimeHelper.GetEndDateOffMinutes(EndDate));
                results.Add(EndDate.Date, endWorkDateHours);
            }
            DateTimeHelper.FillHoursAndDate(listOfSalaryStatus, StartDate.AddDays(1), EndDate.AddDays(-1), ref results);
            var holidays = from days in holidaysArray where days.Date >= StartDate.Date && days.Date.Date <= EndDate select days;
            foreach (var days in holidays)
            {
                var key = days.Date;
                if (results.ContainsKey(key))
                {
                    results.Remove(key);
                }
            }
            return results;
        }

        private String FailedAuthorisationMessage(DateTime date, ATC_AbsenceRequests request, int ManagerId, String message = null)
        {
            if (String.IsNullOrEmpty(message))
            {
                message = String.Format(StringConstants.TimesheetErrorFull, date.ToString(StringConstants.DateOnlyFormat));
            }
            message = String.Format(StringConstants.RequestOf, _employeeService.GetFullName(request.StaffId)) + message;
            EmailModel model = new EmailModel()
            {
                ManagerId = ManagerId,
                RequesterId = request.StaffId,
                Note = message,
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                Template = StringConstants.EmailReject
            };
            if (request.Authoriser1_Id == ManagerId)
            {
                request.isAuthoriser1Approved = false;
                request.Authoriser1Note = message;
            }
            if (request.Authoriser2_Id == ManagerId)
            {
                request.isAuthoriser2Approved = false;
                request.Authoriser2Note = message;
                model.CC = request.Authoriser1_Id;
            }
            request.Status = (int)AbsenceStatus.Rejected;
            _absenceRequestService.Update(request);
            _emailService.SendMailForAbsenceRequest(model);
            return message;
        }

        private string CheckDateOffInRange(ATC_AbsenceRequests request, int managerId)
        {
            var currentStaff = _employeeService.FindById(request.StaffId);
            var joinDate = currentStaff.JoinDate;
            var leaveDate = currentStaff.LeaveDate;
            if (joinDate > request.DateFrom)
            {
                return FailedAuthorisationMessage(joinDate, request, managerId, StringConstants.RequestErrorLessJoinDate);
            }
            if (leaveDate != null && leaveDate.Value < request.DateTo)
            {
                return FailedAuthorisationMessage(joinDate, request, managerId, StringConstants.RequestErrorGreaterLeaveDate);
            }
            return String.Empty;
        }

        private string CheckDateOffInRangeHR(TimeSheetModel model)
        {
            var currentStaff = _employeeService.FindById(model.StaffId);
            var joinDate = currentStaff.JoinDate;
            var leaveDate = currentStaff.LeaveDate;
            if (joinDate > model.DateFrom)
            {
                return StringConstants.RequestErrorLessJoinDate;
            }
            if (leaveDate != null && leaveDate.Value < model.DateTo)
            {
                return StringConstants.RequestErrorGreaterLeaveDate;
            }
            return String.Empty;
        }

        private String GetDayOff(DateTime StartDate, DateTime EndDate, int StaffId)
        {
            var listOfSalaryStatus = _employeeService.GetSalaryStatusBetween(StartDate, EndDate, StaffId);
            var isWorkingHoursChanged = listOfSalaryStatus.Count() == 1 ? false : true;
            DateTime validStartTime = StartDate.Date.AddTicks(startWorkTime.Ticks);
            DateTime validEndTime = EndDate.Date.AddTicks(endWorkTime.Ticks);
            TimeSpan checkvalidStartTime = StartDate.Subtract(validStartTime);
            var holidays = _holidayService.FindAll();
            var holidaysArray = _holidayService.GetAll();
            if (checkvalidStartTime.TotalMinutes < 0)
            {
                return StringConstants.RequestErrorStartTimeInvalid;
            }
            TimeSpan checkvalidEndTime = validEndTime.Subtract(EndDate);
            if (checkvalidEndTime.TotalMinutes < 0)
            {
                return StringConstants.RequestErrorEndTimeInvalid;
            }

            int totalDayBuffer = 0;
            var validStartWorkday = listOfSalaryStatus.First().ATC_WorkingHours.Hours.ToDouble();
            double startDateOffMinute = 0;
            if (DateTimeHelper.isWorkday(StartDate, holidaysArray))
            {
                startDateOffMinute = DateTimeHelper.GetStartDateOffMinutes(StartDate, EndDate, validStartWorkday);
            }
            double endDateOffMinute = 0;
            if (DateTimeHelper.isWorkday(EndDate, holidaysArray))
            {
                endDateOffMinute = DateTimeHelper.GetEndDateOffMinutes(EndDate);
            }
            var validStartWorkingMinutes = Convert.ToDouble(validStartWorkday * 60);
            if (StartDate.Date == EndDate.Date)
            {
                if (startDateOffMinute == validStartWorkingMinutes)
                {
                    return "1 Day";
                }
                return DateTimeHelper.FormatMinutesToTimeString(startDateOffMinute);
            }
            if (StartDate.Date != EndDate.Date && startDateOffMinute == validStartWorkingMinutes)
            {
                totalDayBuffer++;
                startDateOffMinute = 0;
            }
            var validEndWorkday = listOfSalaryStatus.Last().ATC_WorkingHours.Hours;
            var validEndWorkingMinutes = Convert.ToDouble(validEndWorkday * 60);
            if (endDateOffMinute == validEndWorkingMinutes && DateTimeHelper.isWorkday(EndDate, holidaysArray))
            {
                totalDayBuffer++;
                endDateOffMinute = 0;
            }
            if (isWorkingHoursChanged)
            {
                double totalMinutes = startDateOffMinute + endDateOffMinute;
                int totalDays = DateTimeHelper.BusinessDaysUntil(StartDate.AddDays(1), EndDate.AddDays(-1), holidaysArray);
                totalDays += totalDayBuffer;
                String result = "";
                if (totalDays > 1)
                {
                    result += String.Format("{0} Days ", totalDays);
                }
                else if (totalDays == 1)
                {
                    result += String.Format("{0} Day ", totalDays);
                }
                result += DateTimeHelper.FormatMinutesToTimeString(totalMinutes);
                return result;
            }
            else
            {
                double totalMinutesPerDay = Convert.ToDouble(listOfSalaryStatus.First().ATC_WorkingHours.Hours) * 60;
                double totalMinutes = startDateOffMinute + endDateOffMinute;
                if (totalMinutes > totalMinutesPerDay)
                {
                    totalDayBuffer += Convert.ToInt32(Math.Truncate(totalMinutes / totalMinutesPerDay));
                    totalMinutes = totalMinutes % totalMinutesPerDay;
                }
                int totalDays = DateTimeHelper.BusinessDaysUntil(StartDate.AddDays(1), EndDate.AddDays(-1), holidaysArray);
                totalDays += totalDayBuffer;
                String result = "";
                if (totalDays > 1)
                {
                    result += String.Format("{0} Days ", totalDays);
                }
                else if (totalDays == 1)
                {
                    result += String.Format("{0} Day ", totalDays);
                }
                result += DateTimeHelper.FormatMinutesToTimeString(totalMinutes);
                return result;
            }
        }

        private String FailedAuthorisationMessageHR(DateTime date, ATC_AbsenceRequests request, int UserId)
        {
            var message = String.Format(StringConstants.TimesheetErrorFull, date.ToString(StringConstants.DateOnlyFormat));
            message = String.Format(StringConstants.RequestOf, _employeeService.GetFullName(request.StaffId)) + message;
            request.Status = (int)AbsenceStatus.Rejected;
            _absenceRequestService.Update(request);
            EmailModel model = new EmailModel()
            {
                ManagerId = UserId,
                RequesterId = request.StaffId,
                Note = message,
                DateFrom = request.DateFrom,
                DateTo = request.DateTo,
                Template = StringConstants.EmailReject,
                CC = request.Authoriser1_Id,
                BCC = request.Authoriser2_Id
            };
            _emailService.SendMailForAbsenceRequest(model);
            return message;
        }

        private String FailedAuthorisationMessageTimesheet(DateTime date)
        {
            var message = String.Format(StringConstants.TimesheetErrorFull, date.ToString(StringConstants.DateOnlyFormat));
            return message;
        }
    }
}
