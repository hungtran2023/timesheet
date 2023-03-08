using AIS.Domain.TimeSheet;
using System;
using System.ComponentModel.DataAnnotations;

namespace AIS.Models
{
    public class TimeSheet
    {
        public int StaffId { get; set; }

        [Required]
        [Display(Name = ("Type Of Absence"))]
        public int AbsenceType { get; set; }

        [Required]
        [Display(Name = ("From Date"))]
        public String StartDate { get; set; }

        [Required]
        [Display(Name = ("To Date"))]
        public String EndDate { get; set; }

        [Required]
        [Display(Name = ("Hours"))]
        public decimal Hours { get; set; }

        [Display(Name = ("Note"))]
        public String Note { get; set; }

        public TimeSheetModel MapToTimeSheetDTO(TimeSheet model)
        {
            TimeSheetModel result = new TimeSheetModel() {
                StaffId = model.StaffId,
                AbsenceType = model.AbsenceType,
                DateFrom = DateTime.ParseExact(model.StartDate,"dd/MM/yyyy", null),
                DateTo = DateTime.ParseExact(model.EndDate, "dd/MM/yyyy", null),
                Hours = Convert.ToDecimal(model.Hours),
                Note = model.Note
            };
            return result;
        }
    }
}