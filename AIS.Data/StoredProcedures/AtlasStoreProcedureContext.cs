using AIS.Data.EntityBase.StoredProcedures;
using AIS.Data.Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Diagnostics;

namespace AIS.Data.StoredProcedures
{
    public partial class AtlasStoreProcedureContext : IAtlasStoredProcedures
    {

        #region Process Archiving
        public IEnumerable<ProjectArchiveModel> GetProjectArchivingList(string projectStatus)
        {
            using (var ctx = new LeaveManagementContext())
            {
                var listArchivings = ctx.Database.SqlQuery<ProjectArchiveModel>("exec GetListOfProjectArchive @projectStatus",
                      new SqlParameter("projectStatus", projectStatus)).ToList();

                return listArchivings;
            }

        }

        public IEnumerable<ProjectClosingModel> GetProjectClosingValuesList(string ListprojectIDs)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listArchivings = ctx.Database.SqlQuery<ProjectClosingModel>("exec SelectedAPKsToClose @listprojectIDs",
                        new SqlParameter("listprojectIDs", ListprojectIDs)).ToList();
                    return listArchivings;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ProjectClosingModel> GetProjectClosingList()
        {
            using (var ctx = new LeaveManagementContext())
            {
                var listArchivings = ctx.Database.SqlQuery<ProjectClosingModel>("exec GetListOfLiveProjects").ToList();

                return listArchivings;
            }
        }

        public bool UpdateProjectStatus(string ProjectId, int ProjectStatus, decimal ProposalValue, decimal AwardedValue, int StaffID,int statusarchived)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertProjectTracking @ProjTrackerID, @ProjectID,@ProjStatus," +
                        "@ProposalValue,@AwardedValue,@StaffID,@statusArchived",
                     new SqlParameter("ProjTrackerID", -1),
                      new SqlParameter("ProjectID", ProjectId),
                      new SqlParameter("ProjStatus", ProjectStatus),
                      new SqlParameter("ProposalValue", ProposalValue),
                      new SqlParameter("AwardedValue", AwardedValue),
                      new SqlParameter("StaffID", StaffID),
                        new SqlParameter("statusArchived", statusarchived)
                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public IEnumerable<ProjectArchiveModel> GetProjectArchivingValuesList(string projectID)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listArchivings = ctx.Database.SqlQuery<ProjectArchiveModel>("exec GetAProjectArchive @ProjectID",
                        new SqlParameter("ProjectID", projectID)).ToList();
                    return listArchivings;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool DoArchivingProject(string APK, string ServerPath, int StaffID, string ArchiveNote, int isDelete,int ArchivePart)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertProjectArchive @ProArchiveID,@APK,@ServerPath,@StaffID," +
                        "@ArchiveNote,@isDelete,@ArchivePart",
                      new SqlParameter("ProArchiveID", -1),
                     new SqlParameter("APK", APK),
                      new SqlParameter("ServerPath", @ServerPath),
                      new SqlParameter("StaffID", StaffID),
                      new SqlParameter("ArchiveNote", ArchiveNote),
                        new SqlParameter("isDelete", isDelete),
                         new SqlParameter("ArchivePart", ArchivePart)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        #endregion

        #region Sumary Report
        public IEnumerable<TimeSheetHoursModel> GetTotalHoursWorkOfMonth(int staffID, int month)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var total = ctx.Database.SqlQuery<TimeSheetHoursModel>("exec Get_TotalHoursByStaffID @monthcurrent,@staffID",
                        new SqlParameter("monthcurrent", month),
                         new SqlParameter("staffID", staffID)).ToList();
                    return total;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<TimeSheetViewLeaveModel> GetDataTotalViewLeave(int staffID, DateTime dt)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listViewLeaves = ctx.Database.SqlQuery<TimeSheetViewLeaveModel>("exec GetDurationAnnualLeave_2018 @StaffID,@dateTo",
                        new SqlParameter("StaffID", staffID),
                         new SqlParameter("dateTo", dt)).ToList();
                    return listViewLeaves;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<TimeSheetExpiredViewLeave> GetDataExpiredView(int staffID)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var ExpiredView = ctx.Database.SqlQuery<TimeSheetExpiredViewLeave>("exec SP_EXPIRED_DAY_VIEWLEAVE @StaffID",
                        new SqlParameter("StaffID", staffID)
                       ).ToList();
                    return ExpiredView;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<TimsheetReportProjectSector> GetTotalHourProjectSectors(int staffID,string SECTORTYPE)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var ExpiredView = ctx.Database.SqlQuery<TimsheetReportProjectSector>("exec SP_GET_REPORT_SECTOR_HOURPERCENT @STATTID ,@SECTORTYPE",
                        new SqlParameter("STATTID", staffID),
                         new SqlParameter("SECTORTYPE", SECTORTYPE)
                       ).ToList();
                    return ExpiredView;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<TimsheetReportProjectServiceType> GetTotalHourProjectServices(int staffID,string SERVICETYPE)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var ExpiredView = ctx.Database.SqlQuery<TimsheetReportProjectServiceType>("exec SP_GET_REPORT_SERVICECODE_HOURPERCENT @STATTID,@SERVICETYPE",
                        new SqlParameter("@STATTID", staffID),
                        new SqlParameter("@SERVICETYPE", SERVICETYPE)
                       ).ToList();
                    return ExpiredView;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ReportSumaryProjectEmployee> GetReportSumarayHourProjectEmployee(string SectorType, string SeriveCode)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var summaryReport = ctx.Database.SqlQuery<ReportSumaryProjectEmployee>("exec SP_GET_REPORT_SERVICECODE_HOURPERCENT @SectorType ,@SeriveCode",
                        new SqlParameter("@SectorType", SectorType),
                       new SqlParameter("@SeriveCode", SeriveCode)
                       ).ToList();
                    return summaryReport;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ListSectors> GetAllSector()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listsectors = ctx.Database.SqlQuery<ListSectors>("exec SP_GET_ListSectorCombox"
                       ).ToList();
                    return listsectors;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ListServices> GetAllServices()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    try
                    {
                        var listservices = ctx.Database.SqlQuery<ListServices>("exec SP_GET_ListServiceTypeCombox"
                           ).ToList();
                        return listservices;
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }


        #endregion

        #region Employee 

        public IEnumerable<EmployeeModel> GetAllEmoloyeeCurrent()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var employees = ctx.Database.SqlQuery<EmployeeModel>("exec GETALLEmployee"
                       ).ToList();
                    return employees;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        public IEnumerable<EmployeeDetail> GetAllEmoloyeeDetail()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var employees = ctx.Database.SqlQuery<EmployeeDetail>("exec GetAllEmployeeDetails"
                        
                       ).ToList();
                    return employees;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        #endregion

        #region MessageContent
        public IEnumerable<MessageContentModel> GetAllTopMessageContent()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var messageContents = ctx.Database.SqlQuery<MessageContentModel>("exec GetMessageContent"
                       ).ToList();
                    return messageContents;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool InsertMessageContent(string title, string Description, DateTime updateDate,int Id)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertUpdateMessageContent @Title, @Description,@UpdateDate,@action,@Id",
                       
                     new SqlParameter("Title", title),
                      new SqlParameter("Description", Description),
                      new SqlParameter("UpdateDate", updateDate),
                      new SqlParameter("action", "insert"),
                       new SqlParameter("Id", Id)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public bool DeleteMessageContent(string title, string Description, DateTime updateDate, int Id)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertUpdateMessageContent @Title, @Description,@UpdateDate,@action,@Id",
                      new SqlParameter("Title", title),
                      new SqlParameter("Description", Description),
                      new SqlParameter("UpdateDate", updateDate),
                      new SqlParameter("action", "delete"),
                       new SqlParameter("Id", Id)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public bool UpdateMessageContent(string title, string Description, DateTime updateDate, int Id)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertUpdateMessageContent @Title, @Description,@UpdateDate,@action,@Id",

                     new SqlParameter("Title", title),
                      new SqlParameter("Description", Description),
                      new SqlParameter("UpdateDate", updateDate),
                      new SqlParameter("action", "update"),
                       new SqlParameter("Id", Id)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public IEnumerable<TreeViewEmployeeModel> GetTreeViewData(int departemtnId)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listtreeview = ctx.Database.SqlQuery<TreeViewEmployeeModel>("exec GETLEVELOfDeparment @departmentId",
                        new SqlParameter("departmentId", departemtnId)).ToList();
                    return listtreeview;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ListDepartments> GetAllDepartment()
        {

            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listdepartment= ctx.Database.SqlQuery<ListDepartments>("exec GetListDepartermentCombox"
                    ).ToList();
                    return listdepartment;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ProjectTypes> GetAllProjectTypes(int staffID, string apk)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var projectTypes = ctx.Database.SqlQuery<ProjectTypes>("exec GetProjectAssigmentTimeSheet @staffID,@apk",
                   new SqlParameter("staffID", staffID),
                    new SqlParameter("apk", apk)).ToList();
                    return projectTypes;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ListProjectAutoComplete> GetAllProjectAutoComplete(int staffId)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listprojects = ctx.Database.SqlQuery<ListProjectAutoComplete>("exec GetProjectsAutocomplete @staffId",
                      new SqlParameter("staffId", staffId)).ToList();
                    return listprojects;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool DoLogTimeSheetNormal(DateTime tdate, int statffId, int assignmentID, double hours, string note, bool isGenerateAdmin)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec InsertTimeSheetNormal @date, @staffId ,@assignmentID,@hours,@note,@isGenerateAdmin",

                     new SqlParameter("date", tdate),
                      new SqlParameter("staffId", statffId),
                      new SqlParameter("assignmentID", assignmentID),
                      new SqlParameter("hours", hours),
                       new SqlParameter("note", note),
                       new SqlParameter("isGenerateAdmin", isGenerateAdmin)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public IEnumerable<TimeSheetBookedView> GetAllTimeSheetBookedView(int staffId, DateTime date)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listimeSheetsbooked = ctx.Database.SqlQuery<TimeSheetBookedView>("exec GetTimeSheetLogOnDateByStaffID @stattId,@date",
                    new SqlParameter("date", date),
                      new SqlParameter("stattId", staffId)).ToList();
                    return listimeSheetsbooked;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public IEnumerable<ProjectClosingModel> GetProjectNoInputClosingList()
        {
            using (var ctx = new LeaveManagementContext())
            {
                var listArchivings = ctx.Database.SqlQuery<ProjectClosingModel>("exec GetListOfLiveProjectsShouldBeClose").ToList();

                return listArchivings;
            }
        }

        public bool UpdateProjectStatuToClose(string APK)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec UpdateProjectToClosed @apk",
                    
                     new SqlParameter("apk", APK)
                    
                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public IEnumerable<TimeSheetResetModel> GetDataCleanUpTimeSheet()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var listimeSheetsCleanUps = ctx.Database.SqlQuery<TimeSheetResetModel>("exec GetDataTimeSheetCleanUp"
                  ).ToList();
                    return listimeSheetsCleanUps;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool DoResetTimeSheet(string TDate, int StaffID, int AssignmentID, int EventID)
        {
            bool flag = false;
            try
            {
                using (var ctx = new LeaveManagementContext())
                {
                    var insert = ctx.Database.ExecuteSqlCommand("exec CleanUpTimeSheetForEmployee @date,@staffID,@assignmentID,@eventID",

                     new SqlParameter("date", TDate),
                    new SqlParameter("staffID", StaffID),
                      new SqlParameter("assignmentID", AssignmentID),
                       new SqlParameter("eventID", EventID)

                     );

                    flag = true;
                }
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public IEnumerable<AtlasStaff> GetAllStaffAtlas()
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var atlasStaffs = ctx.Database.SqlQuery<AtlasStaff>("exec GetAllStaffAtlas"
                  ).ToList();
                    return atlasStaffs;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool GetPermissionFunctionByUserId(int userId, int functionID)
        {
            using (var ctx = new LeaveManagementContext())
            {
                try
                {
                    var result = ctx.Database.SqlQuery<bool>("exec GetPermissionFunctionByUserId @userId,@functionId",
                    new SqlParameter("userId", userId),
                      new SqlParameter("functionId", functionID)).FirstOrDefault();
                    return result;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

      

        #endregion
    }
}
