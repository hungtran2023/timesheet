using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class TreeViewEmployeeModel
    {
        public string head { get; set; }
        public int id { get; set; }
        public string contents { get; set; }
        public int level { get; set; }
        public string dateHireDisplay { get; set; }

        public string teamdescription { get; set; }
        public int parentID { get; set; }

        public int departmentID { get; set; }

    }
    public class EntityChart
    {
        public string head { get; set; }
        public int id { get; set; }
        public string contents { get; set; }
        public int level { get; set; }

        public string teamdescription { get; set; }
        public string dateHireDisplay { get; set; }
        public int parentID { get; set; }

        public List<EntityChart> children { get; set; }


    }
    public class ListDepartments
    {
        public int DepartmentID { get; set; }

        public string Department { get; set; }

    }
}
