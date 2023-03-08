using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace AIS.Data.Model
{
    public class TimeSheetBookingModel
    {
        public string APK { get; set; }

        public string NameProject { get; set; }

        public int AssignmentID { get; set; }

        public string CurrentDay { get; set; }

        public double Hours { get; set; }

        public string Note { get; set; }

        public int SubtaskID { get; set; }

        public bool IsGenerateAdmin { get; set; }


        public IEnumerable<SelectListItem> subTasks { get; set; }

        public List<TimeSheetBookedView> resultBooks { get; set; }

    }
    public class SubTask
    {
        public int SubtaskID { get; set; }

        public string Name { get; set; }
    }
    public class ProjectTypes
    {
        public int Id { get; set; }

        public string Name { get; set; }
    }

    public class ListProjectAutoComplete
    {
        public string APK { get; set; }

        public string ProjectName { get; set; }
    }
    public class TimeSheetBookedView
    {

        public string Name { get; set; }

        public string Note { get; set; }

        public decimal Hours { get; set; }
    }

}
