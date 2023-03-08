using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class TimeSheetResetModel
    {
     
        public String TDate { get; set; }
        public int StaffID { get; set; }

        public int AssignmentID { get; set; }

        public int EventID { get; set; }

        public decimal Hours { get; set; }

        public decimal OverTime { get; set; }

    }
}
