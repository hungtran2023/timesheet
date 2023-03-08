using AIS.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace AIS.Domain.TimeSheet
{
    public class TimeSheetModel
    {
        public int StaffId { get; set; }
        public int AbsenceType { get; set; }
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public decimal Hours { get; set; }
        public String Note { get; set; }

        public TimeSheetModel()
        {
        }

        public TimeSheetModel(int staffId , DateTime dateFrom , DateTime dateTo)
        {
            StaffId = staffId;
            DateFrom = dateFrom;
            DateTo = dateTo;
        }
    }
}