//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AIS.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class ATC_Timesheet2001
    {
        public System.DateTime TDate { get; set; }
        public int StaffID { get; set; }
        public int AssignmentID { get; set; }
        public int EventID { get; set; }
        public Nullable<decimal> Hours { get; set; }
        public Nullable<decimal> OverTime { get; set; }
        public Nullable<decimal> OverRate { get; set; }
        public string Note { get; set; }
        public decimal OTNight { get; set; }
        public decimal OTNormal { get; set; }
    }
}
