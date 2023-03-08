using AIS.Data;
using AIS.Data.Model;
using AIS.Domain.Base;
using AIS.Domain.Project;
using System;
using System.Collections.Generic;
using System.Linq;
using AIS.Data.StoredProcedures;
using AIS.Data.EntityBase.StoredProcedures;
using System.Web.Mvc;

namespace AIS.Domain.Project
{
    public class ProjectService :  IProjectService
    {
        private IAtlasStoredProcedures _projectService;
        public ProjectService(IAtlasStoredProcedures projectService)
        { 
        
            _projectService = projectService;
        }

        public bool DoArchivingProject(string APK, string ServerPath, int StaffID, string ArchiveNote, int isDelete, int ArchivePart)
        {
            return _projectService.DoArchivingProject(APK, ServerPath, StaffID, ArchiveNote, isDelete,ArchivePart);
        }

        public bool DoResetTimeSheet(string TDate, int StaffID, int AssignmentID, int EventID)
        {
           return _projectService.DoResetTimeSheet(TDate, StaffID, AssignmentID,EventID);
        }

        public IEnumerable<EmployeeDetail> GetAllEmoloyeeDetail()
        {
            return _projectService.GetAllEmoloyeeDetail();
        }

        public IEnumerable<SelectListItem> GetAllSector()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();
            var listitems = _projectService.GetAllSector();
            listprojects.Add(new SelectListItem() { Value = "", Text = "" });
            foreach (var item in listitems)
            {
                listprojects.Add(new SelectListItem() { Value = item.SectorCode, Text = item.SectorName });
            }

            return listprojects;
        }

        public IEnumerable<SelectListItem> GetAllServices()
        {
            List<SelectListItem> listprojects = new List<SelectListItem>();
            var listitems= _projectService.GetAllServices();
            listprojects.Add(new SelectListItem() { Value = "", Text = ""});
            foreach (var item in listitems)
            {
                listprojects.Add(new SelectListItem() { Value = item.ServiceCode, Text = item.SeriveName });
            }

            return listprojects;
        }

        public IEnumerable<TimeSheetResetModel> GetDataCleanUpTimeSheet()
        {
            return _projectService.GetDataCleanUpTimeSheet();
        }

        public bool GetPermissionFunctionByUserId(int userId, int functionID)
        {
            return _projectService.GetPermissionFunctionByUserId(userId, functionID);
        }

        public IEnumerable<ProjectArchiveModel> GetProjectArchivingList(string projectStatus)
        {
            return _projectService.GetProjectArchivingList(projectStatus);
        }

        public IEnumerable<ProjectClosingModel> GetProjectClosingList()
        {
            
            return _projectService.GetProjectClosingList();
        }

        public IEnumerable<ProjectClosingModel> GetProjectNoInputClosingList()
        {
            return _projectService.GetProjectNoInputClosingList();
        }

        public IEnumerable<ProjectArchiveModel> GetProjectValueArchivingList(string ProjectId)
        {
            return _projectService.GetProjectArchivingValuesList(ProjectId);
        }

        public IEnumerable<ProjectClosingModel> GetProjectValueClosingList(string listProjectIds)
        {
            return _projectService.GetProjectClosingValuesList(listProjectIds);
        }

        public IEnumerable<ReportSumaryProjectEmployee> GetReportSumarayHourProjectEmployee(string SectorType, string SeriveCode)
        {
            return _projectService.GetReportSumarayHourProjectEmployee(SectorType, SeriveCode);
        }

        public IEnumerable<TimsheetReportProjectSector> GetTotalHourProjectSectors(int staffID, string SECTORTYPE)
        {
            return _projectService.GetTotalHourProjectSectors(staffID,SECTORTYPE);
        }

        public IEnumerable<TimsheetReportProjectServiceType> GetTotalHourProjectServices(int staffID,string SERVICETYPE)
        {
            return _projectService.GetTotalHourProjectServices(staffID,SERVICETYPE);
        }

        public bool UpdateProjectStatus(string ProjectId, int ProjectStatus, decimal ProposalValue, decimal AwardedValue, int StaffID,int statusArchived)
        {
            return _projectService.UpdateProjectStatus(ProjectId, ProjectStatus, ProposalValue,AwardedValue, StaffID, statusArchived);
        }

        public bool UpdateProjectStatuToClose(string APK)
        {
            return _projectService.UpdateProjectStatuToClose(APK);
        }
    }
}
