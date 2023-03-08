using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class EmployeeModel
    {
        public int PersonID { get; set; }

        public string IDNumber { get; set; }
        
        public string FullName { get; set; }
        
        public string FirstName { get; set; }
        
        public string LastName { get; set; }

        public string Country { get; set; }

        public string EmailAddress { get; set; }    

        public string Phone { get; set; }

        public string MobilePhone { get; set; }

        public string  JobTitle { get; set; }

        public string Department { get; set; }

        public string CardNumber { get; set; }


    }

    public class EmployeeDetail
    {
        public int PersonID { get; set; }

        public string StaffID { get; set; }

        public string Fullname { get; set; }

        public string Birthday { get; set; }

        public string StartDate { get; set; }

        public string JobTitle { get; set; }

        public string Department { get; set; }

        public string ReportTo { get; set; }

        public string CSOLevel { get; set; }

        public int CompanyID { get; set; }
        public string CharCode { get; set; }

        public string Record { get; set; }


    }
}
