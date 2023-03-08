using AIS.Data.Model;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace AIS.Domain.DashBoard.Interfaces
{
    public interface IDashBoardService
    {
        IEnumerable<TimeSheetHoursModel> GetTotalHoursWorkOfMonth(int staffID, int month);

        IEnumerable<TimeSheetViewLeaveModel> GetDataTotalViewLeave(int staffID, DateTime dt);

        IEnumerable<TimeSheetExpiredViewLeave> GetDataExpiredView(int staffID);

        IEnumerable<EmployeeModel> GetAllEmoloyeeCurrent();


        IEnumerable<MessageContentModel> GetAllTopMesageContent();

        Boolean InsertMessageContent(string title, string Description, DateTime updateDate, int Id);

        Boolean DeleteMessageContent(string title, string Description, DateTime updateDate, int Id);

        Boolean UpdateMessageContent(string title, string Description, DateTime updateDate, int Id);

        IEnumerable<TreeViewEmployeeModel> GetTreeViewData(int departemtnId);
        IEnumerable<SelectListItem> GetAllDepartment();

        IEnumerable<SelectListItem> GetAllProjectTypes(int staffID, string apk);


        IEnumerable<SelectListItem> GetAllProjectAutoComplete(int staffId);


        Boolean DoLogTimeSheetNormal(DateTime tdate, int statffId, int assignmentID, double hours, string note, bool isGenerateAdmin);

        IEnumerable<TimeSheetBookedView> GetAllTimeSheetBookedView(int staffId, DateTime date);

        IEnumerable<AtlasStaff> GetAllStaffAtlas();
    }
}
