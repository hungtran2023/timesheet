using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public  class  ProjectClosingModel
    {
        public string ProjectKey { get; set; }
        public string ProjectName { get; set; }
        public int ManagerID { get; set; }
        public string Manager { get; set; }
        public string Department { get; set; }
        public string  BDM { get; set; }
        public string ProjStatus { get; set; }
        public string Server { get; set; }

        public string statusData { get; set; }

        public Decimal Proposal { get; set; }

        public Decimal Awarded { get; set; }
        public Decimal Invoice { get; set; }

    }

    public class ProjectArchiveModel
    {
        public string myLegacyDataString;
        [DisplayName("APK:")]
        public string ProjectKey { get; set; }
        [DisplayName("Project Name:")]
        public string ProjectName { get; set; }

        public string Note
        {
            get; set; 
        }

        [DisplayName("Archive Date:")]
        public string ArchiveDate { get; set; }


        public string ProjStatus { get; set; }
        [DisplayName("Server Path:")]
        public string ServerPath { get; set; }


    }
}
