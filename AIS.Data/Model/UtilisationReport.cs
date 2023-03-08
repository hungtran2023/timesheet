using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class UtilisationReport
    {
        public List<GroupNameReport> ReportGroups { get; set; }  
        
    }

    public class GroupNameReport
    {
        public string GroupName { get; set; }
        public string Department { get; set; }
        public double Billablehrs { get; set; }

        public double OT { get; set; }

        public double EstTraining { get; set; }

        public double BDMdowntime { get; set; }

        public double Projectdowntime { get; set; }

        public double TotalProjectHours { get; set; }

        public double Atlasproject { get; set; }

        public double GA { get; set; }

        public double Nonprojectdowntime { get; set; }

        public double TotalNonprojects { get; set; }

        public double Availablehours { get; set; }

        public double BillableUtilization { get; set; }

        public double Utilization { get; set; }
    }

}
