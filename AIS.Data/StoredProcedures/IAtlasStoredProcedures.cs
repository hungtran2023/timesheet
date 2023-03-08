using AIS.Data.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.EntityBase.StoredProcedures
{
    public interface IAtlasStoredProcedures
    {
        #region Archiving
        IEnumerable<ProjectClosingModel> GetProjectClosingList();

        Boolean UpdateProjectStatus(string ProjectId, int ProjectStatus, decimal ProposalValue, decimal AwardedValue, int StaffID, int statusarchived);
        IEnumerable<ProjectClosingModel> GetProjectClosingValuesList(string ListprojectIDs);

        IEnumerable<ProjectArchiveModel> GetProjectArchivingList(string projectStatus);

        IEnumerable<ProjectArchiveModel> GetProjectArchivingValuesList(string projectID);

        IEnumerable<ProjectClosingModel> GetProjectNoInputClosingList();

        Boolean DoArchivingProject(string APK, string ServerPath, int StaffID, string ArchiveNote, int isDelete, int ArchivePart);

        Boolean UpdateProjectStatuToClose(string APK);

        Boolean GetPermissionFunctionByUserId(int userId,int functionID);

        IEnumerable<TimeSheetResetModel> GetDataCleanUpTimeSheet();

        Boolean DoResetTimeSheet(string TDate, int StaffID, int AssignmentID,int EventID);
        #endregion
        //Get Inforamtion Data for DashBoard
        #region Dashboard
        IEnumerable<TimeSheetHoursModel> GetTotalHoursWorkOfMonth(int staffID, int month);

        IEnumerable<TimeSheetViewLeaveModel> GetDataTotalViewLeave(int staffID, DateTime dt);

        IEnumerable<TimeSheetExpiredViewLeave> GetDataExpiredView(int staffID);
        #endregion

        #region SumaryReport
        IEnumerable<TimsheetReportProjectSector> GetTotalHourProjectSectors(int staffID, string SECTORTYPE);
        IEnumerable<TimsheetReportProjectServiceType> GetTotalHourProjectServices(int staffID, string SERVICETYPE);

        IEnumerable<ReportSumaryProjectEmployee> GetReportSumarayHourProjectEmployee(string SectorType, string SeriveCode);

        IEnumerable<ListSectors> GetAllSector();
        IEnumerable<ListServices> GetAllServices();
        #endregion

        #region Employee
        IEnumerable<EmployeeModel> GetAllEmoloyeeCurrent();
        IEnumerable<EmployeeDetail> GetAllEmoloyeeDetail();
        #endregion
        #region MessageContent
        IEnumerable<MessageContentModel> GetAllTopMessageContent();

        Boolean InsertMessageContent(string title, string Description,DateTime updateDate,int Id);

        Boolean DeleteMessageContent(string title, string Description, DateTime updateDate, int Id);

        Boolean UpdateMessageContent(string title, string Description, DateTime updateDate, int Id);

        #endregion

        #region TreeView
        IEnumerable<TreeViewEmployeeModel> GetTreeViewData(int departemtnId);
        IEnumerable<ListDepartments> GetAllDepartment();
        #endregion
        #region TimeSheetBooking
        IEnumerable<ProjectTypes> GetAllProjectTypes(int staffID,string apk);

        IEnumerable<ListProjectAutoComplete> GetAllProjectAutoComplete(int staffId);

        Boolean DoLogTimeSheetNormal(DateTime tdate, int statffId, int assignmentID, double hours, string note, bool isGenerateAdmin);


        IEnumerable<TimeSheetBookedView> GetAllTimeSheetBookedView(int staffId, DateTime date);
        #endregion

        #region AtlasStaff

        IEnumerable<AtlasStaff> GetAllStaffAtlas();
        #endregion


    }

}
