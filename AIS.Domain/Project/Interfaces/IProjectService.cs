using AIS.Data.Model;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Web.Mvc;

namespace AIS.Domain.Project
{
    public interface IProjectService 
    {
        //HR_Employee GetEmployeeInfoById(int Id);
        IEnumerable<ProjectClosingModel> GetProjectClosingList();

        IEnumerable<ProjectArchiveModel> GetProjectArchivingList(string projectStatus);
        Boolean UpdateProjectStatus(string ProjectId, int ProjectStatus, decimal ProposalValue, decimal AwardedValue, int StaffID, int statusarchived);

        IEnumerable<ProjectClosingModel> GetProjectValueClosingList(string listProjectIds);

        IEnumerable<ProjectArchiveModel> GetProjectValueArchivingList(string ProjectId);

         IEnumerable<ProjectClosingModel> GetProjectNoInputClosingList();

        Boolean DoArchivingProject(string APK, string ServerPath, int StaffID, string ArchiveNote, int isDelete, int ArchivePart);

        Boolean UpdateProjectStatuToClose(string APK);

        IEnumerable<TimsheetReportProjectSector> GetTotalHourProjectSectors(int staffID, string SECTORTYPE);
        IEnumerable<TimsheetReportProjectServiceType> GetTotalHourProjectServices(int staffID, string SERVICETYPE);
        IEnumerable<ReportSumaryProjectEmployee> GetReportSumarayHourProjectEmployee(string SectorType, string SeriveCode);

        IEnumerable<SelectListItem> GetAllSector();

        IEnumerable<SelectListItem> GetAllServices();


        IEnumerable<TimeSheetResetModel> GetDataCleanUpTimeSheet();

        Boolean DoResetTimeSheet(string TDate, int StaffID, int AssignmentID, int EventID);

        Boolean GetPermissionFunctionByUserId(int userId, int functionID);
        IEnumerable<EmployeeDetail> GetAllEmoloyeeDetail();
    }
}
