using System;
using System.Collections.Generic;
using AIS.Data;
using AIS.Domain.Base;
using AIS.Domain.TimeSheet;

namespace AIS.Domain.AbsenceRequest
{
    public interface IAbsenceRequestService : IService<ATC_AbsenceRequests>
    {
        IEnumerable<ATC_AbsenceRequests> ListRequestView();
        void InProgressRequest(int Id, String Note, int ManagerId);
        bool IsRequestMade(DateTime date, int staffId);
        bool IsRequestMade(DateTime date, int staffId, int requestId);
    }
}
