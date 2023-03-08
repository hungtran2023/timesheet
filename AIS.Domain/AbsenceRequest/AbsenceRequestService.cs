using System;
using System.Collections.Generic;
using System.Linq;
using AIS.Data;
using AIS.Domain.Base;
using AIS.Domain.Holiday;
using AIS.Domain.Employee;
using AIS.Domain.TimeSheet;
using AIS.Domain.Email;
using AIS.Domain.HREmployee;
using AIS.Domain.Common.Helper;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Constants;
using AIS.Domain.Email.Interfaces;
using System.Data.Entity;

namespace AIS.Domain.AbsenceRequest
{
    public class AbsenceRequestService : Service<ATC_AbsenceRequests>, IAbsenceRequestService
    {

        public AbsenceRequestService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }
        public override ATC_AbsenceRequests Add(ATC_AbsenceRequests target)
        {
            target.DateSubmitted = DateTime.Now;
            return base.Add(target);
        }
        public IEnumerable<ATC_AbsenceRequests> ListRequestView() {
            return   this.FindAll().AsQueryable().Include(t => t.ATC_Events).Include(t => t.Staff.PersonalInfo);
        }
        public void InProgressRequest(int Id, String Note, int ManagerId)
        {
            var request = FindById(Id);
            if (request.Authoriser1_Id == ManagerId)
            {
                request.isAuthoriser1Approved = true;
                request.Authoriser1Note = Note;
            }
            else if (request.Authoriser2_Id == ManagerId)
            {
                request.isAuthoriser2Approved = true;
                request.Authoriser2Note = Note;
            }
            else
            {
                request.isHrApproved = true;
                request.HrNote = Note;
            }
            request.Status = (int)AbsenceStatus.InProgress;
            this.Update(request);
        }
        public bool IsRequestMade(DateTime date, int staffId)
        {
            var target = from t in FindAll()
                         where t.DateFrom.Date <= date.Date && t.DateTo.Date >= date.Date && t.StaffId == staffId && t.Status != (int)AbsenceStatus.Rejected
                         select t;
            return target.Count() > 0;
        }

        public bool IsRequestMade(DateTime date, int staffId, int requestId)
        {
            var target = from t in FindAll()
                         where t.DateFrom.Date <= date.Date && t.DateTo.Date >= date.Date && t.Id != requestId && t.StaffId == staffId && t.Status != (int)AbsenceStatus.Rejected
                         select t;
            return target.Count() > 0;
        }
    }
}
