using AIS.Data;
using AIS.Data.Model;
using AIS.Domain.Base;
using AIS.Domain.Project;
using System;
using System.Collections.Generic;
using System.Linq;
using AIS.Data.StoredProcedures;
using AIS.Data.EntityBase.StoredProcedures;
using AIS.Domain.DashBoard.Interfaces;
using System.Web.Mvc;

namespace AIS.Domain.DashBoard
{
    public class DashBoardService : IDashBoardService
    {
        private IAtlasStoredProcedures _dashboardService;

        public DashBoardService(IAtlasStoredProcedures dashBoardService)
        {

            _dashboardService = dashBoardService;
        }

        public bool DeleteMessageContent(string title, string Description, DateTime updateDate, int Id)
        {
            return _dashboardService.DeleteMessageContent(title, Description,updateDate,Id);
        }

        public bool DoLogTimeSheetNormal(DateTime tdate, int statffId, int assignmentID, double hours, string note, bool isGenerateAdmin)
        {
           return _dashboardService.DoLogTimeSheetNormal(tdate, statffId, assignmentID, hours, note, isGenerateAdmin);  
        }

        public IEnumerable<SelectListItem> GetAllDepartment()
        {
            List<SelectListItem> listdepartments = new List<SelectListItem>();
            var listitems = _dashboardService.GetAllDepartment();
            listdepartments.Add(new SelectListItem() { Value = "", Text = "" });
            foreach (var item in listitems)
            {
                listdepartments.Add(new SelectListItem() { Value = item.DepartmentID.ToString(), Text = item.Department });
            }

            return listdepartments;
        }

        public IEnumerable<EmployeeModel> GetAllEmoloyeeCurrent()
        {
           return _dashboardService.GetAllEmoloyeeCurrent();
        }

        public IEnumerable<SelectListItem> GetAllProjectAutoComplete(int staffId)
        {
            List<SelectListItem> listdepartments = new List<SelectListItem>();
            var listitems = _dashboardService.GetAllProjectAutoComplete(staffId);
            listdepartments.Add(new SelectListItem() { Value = "", Text = "" });
            foreach (var item in listitems)
            {
                listdepartments.Add(new SelectListItem() { Value = item.APK.ToString(), Text = item.ProjectName });
            }

            return listdepartments;
        }

        public IEnumerable<SelectListItem> GetAllProjectTypes(int staffID,string apk)
        {

            List<SelectListItem> listdepartments = new List<SelectListItem>();
            var listitems = _dashboardService.GetAllProjectTypes(staffID,apk);
            listdepartments.Add(new SelectListItem() { Value = "-1", Text = "Select a Task" });
            foreach (var item in listitems)
            {
                listdepartments.Add(new SelectListItem() { Value = item.Id.ToString(), Text = item.Name });
            }

            return listdepartments;
        }

        public IEnumerable<TimeSheetBookedView> GetAllTimeSheetBookedView(int staffId, DateTime date)
        {
            return _dashboardService.GetAllTimeSheetBookedView(staffId, date);
        }

        public IEnumerable<MessageContentModel> GetAllTopMesageContent()
        {
            return _dashboardService.GetAllTopMessageContent();
        }

        public IEnumerable<TimeSheetExpiredViewLeave> GetDataExpiredView(int staffID)
        {
            return _dashboardService.GetDataExpiredView(staffID);
        }

        public IEnumerable<TimeSheetViewLeaveModel> GetDataTotalViewLeave(int staffID, DateTime dt)
        {
            return _dashboardService.GetDataTotalViewLeave(staffID, dt);
        }

        public IEnumerable<TimeSheetHoursModel> GetTotalHoursWorkOfMonth(int staffID, int month)
        {
            return _dashboardService.GetTotalHoursWorkOfMonth(staffID, month);
        }

        public IEnumerable<TreeViewEmployeeModel> GetTreeViewData(int departemtnId)
        {
            return _dashboardService.GetTreeViewData(departemtnId);
        }

        public bool InsertMessageContent(string title, string Description, DateTime updateDate,int Id)
        {
            return _dashboardService.InsertMessageContent(title, Description, updateDate,Id);
        }

        public bool UpdateMessageContent(string title, string Description, DateTime updateDate, int Id)
        {
            return _dashboardService.UpdateMessageContent(title,Description,DateTime.Now,Id);
        }

        public IEnumerable<AtlasStaff> GetAllStaffAtlas()
        {
            return _dashboardService.GetAllStaffAtlas();
        }
    }
}
