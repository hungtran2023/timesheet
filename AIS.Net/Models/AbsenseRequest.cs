using System;
using AIS.Data;
using System.ComponentModel.DataAnnotations;
using AIS.Domain.AbsenceRequest;

namespace AIS.Models
{

    public class AbsenseRequest
    {
        public int? Id { get; set; }

        [Required]
        [Display(Name =("Type Of Absence"))]
        public int AbsenceType { get; set; }

        [Required]
        [Display(Name =("From Date"))]
        public String StartDate { get; set; }

        [Required]
        [Display(Name = ("To Date"))]
        public String EndDate { get; set; }

        [Required]
        [Display(Name = ("From"))]
        public String StartTime { get; set; }

        [Required]
        [Display(Name = ("To"))]
        public String EndTime { get; set; }

        [Required]
        [Display(Name = ("Authoriser 1"))]
        public int FirstAuthoriserId { get; set; }

        [Required]
        [Display(Name = ("Authoriser 2"))]
        public int? SecondAuthoriserId { get; set; }

        [Display(Name =("Note"))]
        public String Note { get; set; }

        [Display(Name =("Authorized by Hr"))]
        public bool isAuthorizedByHr { get; set; }

        private DateTime GetDate(String date, String time) {
            var tmpDate = DateTime.Now;
            var timeSpliter = time.Split(':');
            if (timeSpliter[0].Length > 1)
            {
                DateTime.TryParseExact(String.Format("{0} {1}", date, time), "dd/MM/yyyy HH:mm", null, System.Globalization.DateTimeStyles.None, out tmpDate);
            }
            else
            {
                DateTime.TryParseExact(String.Format("{0} {1}", date, time), "dd/MM/yyyy H:mm", null, System.Globalization.DateTimeStyles.None, out tmpDate);
            }
            return tmpDate;
        }

        public ATC_AbsenceRequests ToModel(int StaffId)
        {
            var temp = new ATC_AbsenceRequests();
            if (Id != null)
            {
                temp.Id = (int)Id;
            }
            temp.StaffId = StaffId;
            temp.Authoriser1_Id = FirstAuthoriserId;
            temp.Authoriser2_Id = SecondAuthoriserId == 0 ? null : SecondAuthoriserId;
            temp.DateFrom = GetDate(StartDate, StartTime);
            temp.DateTo = GetDate(EndDate, EndTime);
            temp.Note = Note;
            temp.Status = 0;
            temp.Type = AbsenceType;
            temp.isAuthorisedByHr = isAuthorizedByHr;
            return temp;
        }
        public ATC_AbsenceRequests ToModel(AbsenceRequestHandler serviceHandler , double balance, int StaffId)
        {
            var temp = new ATC_AbsenceRequests();
            if (Id != null)
            {
                temp.Id = (int)Id;
            }
            temp.StaffId = StaffId;
            temp.Authoriser1_Id = FirstAuthoriserId;
            temp.Authoriser2_Id = SecondAuthoriserId == 0 ? null : SecondAuthoriserId;
            temp.DateFrom = GetDate(StartDate, StartTime);
            temp.DateTo = GetDate(EndDate, EndTime);
            temp.Note = Note;
            temp.Status = 0;
            temp.Type = AbsenceType;
            temp.isAuthorisedByHr = isAuthorizedByHr;
            if (serviceHandler.TotalWorkingHours(temp.DateFrom , temp.DateTo , StaffId) > balance)
            {
                temp.isAuthorisedByHr = true;
            }
            return temp;
        }
    }
}