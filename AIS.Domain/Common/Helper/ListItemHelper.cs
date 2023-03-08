using System.Collections.Generic;
using System.Web.Mvc;
using System.Globalization;
using System.Linq;
using System;
using AIS.Domain.Common.Enum;

namespace AIS.Domain.Common.Helper
{
    public class ListItemHelper
    {
        public static List<SelectListItem> GetStatusList()
        {
            List<SelectListItem> listOfStatus = new List<SelectListItem>();
            listOfStatus.Add(new SelectListItem() { Value = "-1", Text = "" });
            listOfStatus.Add(new SelectListItem() { Value = ((int)AbsenceStatus.UnAuthorised).ToString(), Text = AbsenceTypeHelper.ConvertAbsenceStatus(AbsenceStatus.UnAuthorised) });
            listOfStatus.Add(new SelectListItem() { Value = ((int)AbsenceStatus.Authorised).ToString(), Text = AbsenceTypeHelper.ConvertAbsenceStatus(AbsenceStatus.Authorised) });
            listOfStatus.Add(new SelectListItem() { Value = ((int)AbsenceStatus.Rejected).ToString(), Text = AbsenceTypeHelper.ConvertAbsenceStatus(AbsenceStatus.Rejected) });
            listOfStatus.Add(new SelectListItem() { Value = ((int)AbsenceStatus.Taken).ToString(), Text = AbsenceTypeHelper.ConvertAbsenceStatus(AbsenceStatus.Taken) });
            return listOfStatus;
        }

        public static List<SelectListItem> GetChooseProjectList()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();
            listprojects.Add(new SelectListItem() { Value = "-1", Text = "" });
            listprojects.Add(new SelectListItem() { Value = "1", Text = "no data input for 3 months" });
            listprojects.Add(new SelectListItem() { Value = "2", Text ="your projects" });
            listprojects.Add(new SelectListItem() { Value = "3", Text = "all projects" });
        

            return listprojects;
        }

        public static List<SelectListItem> GetChooseEmployeeList()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();
            listprojects.Add(new SelectListItem() { Value = "", Text = "All" });
            listprojects.Add(new SelectListItem() { Value = "ATL", Text = "Atlas Employee" });
            listprojects.Add(new SelectListItem() { Value = "BPO", Text = "BPO Statff" });          
            return listprojects;
        }

        public static List<SelectListItem> GetFillterSearchEmployeeAtlas()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();          
            listprojects.Add(new SelectListItem() { Value = "1", Text = "By Fullname" });      
            listprojects.Add(new SelectListItem() { Value = "2", Text = "By Jobtitle" });
            listprojects.Add(new SelectListItem() { Value = "3", Text = "By StaffID" });
            listprojects.Add(new SelectListItem() { Value = "4", Text = "By Department" });


            return listprojects;
        }

        public static List<SelectListItem> GetChooseProjectStatusList()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();
            listprojects.Add(new SelectListItem() { Value = "", Text = "" });
            listprojects.Add(new SelectListItem() { Value = "beArchive", Text = "To be archived" });
            listprojects.Add(new SelectListItem() { Value = "NotYet", Text = "Not yet" });
            listprojects.Add(new SelectListItem() { Value = "Archive", Text = "Archived" });
            listprojects.Add(new SelectListItem() { Value = "OneArchive", Text = "Multi-Part Archived" });

            return listprojects;
        }

        public static IEnumerable<SelectListItem> Months()
        {
            return DateTimeFormatInfo
                    .InvariantInfo
                    .MonthNames
                    .Where(month => !String.IsNullOrEmpty(month))
                    .Select((monthName, index) => new SelectListItem
                    {
                        Value = (index + 1).ToString(),
                        Text = monthName
                    });
        }

        public static List<SelectListItem> Years()
        {
            var listOfYear = new List<SelectListItem>();
            var currentYear = DateTime.Now.Year + 1;
            for(var i = 2000; i < currentYear + 1; i++){
                var item = new SelectListItem
                {
                    Value = i.ToString(),
                    Text = i.ToString()
                };
                listOfYear.Add(item);
            }
            return listOfYear;
        }

      
    }
}