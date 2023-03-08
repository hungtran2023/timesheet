using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using AIS.Models;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.HRReport;
using AIS.Domain.AnualLeaveDays;
using AIS.Domain.TimeSheet;
using AIS.Domain.Event;
using AIS.Domain.HREmployee;
using AIS.Domain.Project;
using AIS.Domain.Holiday;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Enum;
using AIS.Domain.Common.Helper;
using AIS.Domain.Email;
using PagedList;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.Text;
using AIS.Data;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Web;
using System.IO;

namespace AIS.Controllers
{
    public class ProjectArchivingController : MenuBaseController

    {
      
        private readonly IProjectService _projectService = Inject.Service<IProjectService>();
        private readonly IEmailService _emailService = Inject.Service<IEmailService>();
        private readonly AbsenceRequestHandler serviceHandler = Inject.Service<AbsenceRequestHandler>();
       // private readonly int UserId = 2272;
      private readonly int PageSize = 100;
       

        public ActionResult ProjectClosing()
        {
            
            ViewBag.PageSize = PageSize;
            ViewBag.LoginPageUrl = StringConstants.ProjectClosingEmailRedirect;
            ViewBag.ProjectChooseList = ListItemHelper.GetChooseProjectList();
            return View();

        }

        public ActionResult ProjectArchiving()
        {

      
            ViewBag.PageSize = PageSize;
            ViewBag.LoginPageUrl = StringConstants.MessageURL;
            ViewBag.ProjectStatusChooseList = ListItemHelper.GetChooseProjectStatusList();
            return View();

        }

        [HttpGet]
        public JsonResult GetClosingData(string data = "")
        {

            var lstProjectArchiving = _projectService.GetProjectClosingList();
            List<AIS.Data.Model.ProjectClosingModel> list = new List<Data.Model.ProjectClosingModel>();

            if (data == "")
            {
                list = lstProjectArchiving.ToList();
            }
            if (data == StringConstants.noinputdataProjects)
            {
                list = _projectService.GetProjectNoInputClosingList().ToList();
            }

            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;
        }

        [HttpGet]
        public JsonResult GetArchivingData()
        {

            var lstProjectArchiving = _projectService.GetProjectArchivingList("");
            List<AIS.Data.Model.ProjectArchiveModel> list = new List<Data.Model.ProjectArchiveModel>();
            list = lstProjectArchiving.ToList();
            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;
        }

        [HttpPost]
        public ActionResult GetClosingValuesData(List<string> ProjectIds)
        {
            StringBuilder listProjectIds = new StringBuilder();

            foreach (var apk in ProjectIds)
            {
                if (ProjectIds.Count > 1)
                {

                    listProjectIds.Append(apk.ToString());
                    listProjectIds.Append(",");
                }
                else
                {
                    listProjectIds.Append(apk.ToString());
                }
            }

            var lstProjectArchiving = _projectService.GetProjectValueClosingList(listProjectIds.ToString());
            List<AIS.Data.Model.ProjectClosingModel> list = new List<Data.Model.ProjectClosingModel>();
            list = lstProjectArchiving.ToList();
            ViewBag.ListProjectCloses = listProjectIds.ToString();
            return PartialView("~/Views/ProjectArchiving/_ModalClosingChoose.cshtml", list);
        }

        [HttpPost]
        public ActionResult GetArchivingValuesData(string ProjectId)
        {

            var checkPersionFunction = _projectService.GetPermissionFunctionByUserId(UserId, 91);
            ViewBag.PermissionUpdate = checkPersionFunction;

            var ProjectArchiving = _projectService.GetProjectValueArchivingList(ProjectId).FirstOrDefault();
            AIS.Data.Model.ProjectArchiveModel model = new Data.Model.ProjectArchiveModel();
            if (ProjectArchiving.Note != null && ProjectArchiving.Note.Contains("<br />"))
            {
                model.Note = ProjectArchiving.Note.Replace("<br />", "\r\n");
            }
            else
            {
                model.Note = ProjectArchiving.Note;
            }
            model.ProjectKey = ProjectArchiving.ProjectKey;
            model.ServerPath = ProjectArchiving.ServerPath;
            model.ArchiveDate = ProjectArchiving.ArchiveDate;
            model.ProjectName = ProjectArchiving.ProjectName;
            model.ProjStatus = ProjectArchiving.ProjStatus;
            return PartialView("~/Views/ProjectArchiving/_ModalArchivingChoose.cshtml", model);
        }

        [HttpPost]
        public JsonResult UpdateProjectStatus(List<string> ProjectIds,string projectCloseIds)
        {

            StringBuilder listProjectIds = new StringBuilder();

            StringBuilder listProjectCloseIds = new StringBuilder();

            StringBuilder listProjectIdleft7Char = new StringBuilder();

            StringBuilder listProjectIdleft3Char = new StringBuilder();
            if (ProjectIds != null)
            {
                foreach (var apk in ProjectIds)
                {
                    if (ProjectIds.Count > 1)
                    {
                        listProjectIds.Append(apk.ToString());
                        listProjectIds.Append(",");

                        var apkleft = apk.Substring(0, 10);
                        var apk3left = apk.Substring(0, 3);
                        listProjectIdleft7Char.Append(apkleft.ToString());
                        listProjectIdleft7Char.Append(",");
                        listProjectIdleft3Char.Append(apk3left.ToString());
                        listProjectIdleft3Char.Append(",");
                    }
                    else
                    {
                        var apkleft = apk.Substring(0, 10);
                        var apk3left = apk.Substring(0, 3);
                        listProjectIdleft7Char.Append(apkleft.ToString());
                        listProjectIds.Append(apk.ToString());
                        listProjectIdleft3Char.Append(apk3left.ToString());
                    }
                }
            }

            //foreach (var apkclose in projectCloseIds)
            //{
            //    if (projectCloseIds.Count > 1)
            //    {
            //        listProjectCloseIds.Append(apkclose.ToString());
            //        listProjectCloseIds.Append(",");
            //    }
            //    else
            //    {
            //        listProjectCloseIds.Append(apkclose.ToString());
            //    }
            //}

            var lstProjectArchiving = _projectService.GetProjectValueClosingList(listProjectIds.ToString());

            foreach (var item in lstProjectArchiving)
            {
                var update = _projectService.UpdateProjectStatus(item.ProjectKey, StringConstants.ProjectStatusClose, item.Proposal, item.Awarded, UserId,1);
            }

            var lstProjectClose = _projectService.GetProjectValueClosingList(projectCloseIds.ToString());

            foreach (var item in lstProjectClose)
            {
               // var update = _projectService.UpdateProjectStatuToClose(item.ProjectKey);
                var update = _projectService.UpdateProjectStatus(item.ProjectKey, StringConstants.ProjectStatusClose, item.Proposal, item.Awarded, UserId, 0);
            }


            string s = DateTime.Now.ToString("HH:mm:ss");

            List<EmailModel> emailmodel = new List<EmailModel>();

            foreach (var item in lstProjectArchiving)
            {
                EmailModel model = new EmailModel()
                {
                    ManagerId = UserId,
                    Template = StringConstants.EmailProjectArchiving,
                    Note = listProjectIds.ToString(),
                    APK7Character = item.ProjectKey.Substring(0,10),
                    ServerPath = item.Server,
                    TimeClose = s + " - " + DateTime.Now.ToString("dd/MM/yyyy")
                    

                };
                emailmodel.Add(model);
            }                      

            foreach (var model in emailmodel)
            {
                _emailService.SendMailForProjectArchiving(model);
            }
         
            var lstProjectArchivingCurrent = _projectService.GetProjectClosingList();
            List<AIS.Data.Model.ProjectClosingModel> list = new List<Data.Model.ProjectClosingModel>();
            list = lstProjectArchivingCurrent.ToList();
            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;

        }

        public JsonResult GetListProjectNoInput()
        {
            List<AIS.Data.Model.ProjectClosingModel> list = new List<Data.Model.ProjectClosingModel>();
            list = _projectService.GetProjectNoInputClosingList().ToList();
            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;
        }

        [HttpPost]
        public JsonResult ChangeYourProjectStatus(string projectChoose)
        {

            var lstProjectArchiving = _projectService.GetProjectClosingList();
            List<AIS.Data.Model.ProjectClosingModel> list = new List<Data.Model.ProjectClosingModel>();
            list = lstProjectArchiving.ToList();

            if (projectChoose == StringConstants.yourProjects)
            {
                list = list.Where(x => x.ManagerID == UserId).ToList();
            }
            if (projectChoose == StringConstants.noinputdataProjects)
            {
                list = _projectService.GetProjectNoInputClosingList().ToList();
            }

            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;

        }

        [HttpPost]
        public JsonResult ChangeYourProjectStatusArchiving(string projectChoose)
        {

            var lstProjectArchiving = _projectService.GetProjectArchivingList(projectChoose);
            List<AIS.Data.Model.ProjectArchiveModel> list = new List<Data.Model.ProjectArchiveModel>();
            list = lstProjectArchiving.ToList();
            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;
        }

        [HttpPost, ValidateInput(false)]

        public ActionResult DoArchivingProject(string ProjectKey, string Note)
        {
            string NoteFromat = Note.Replace("\r\n", "<br />");

            var count = NoteFromat.Split(new[] { "<br /><br />" }, StringSplitOptions.None).Length - 1;

            if (count == 1)
            {
                NoteFromat = NoteFromat.Replace("<br />", "");
            }

            var ProjectArchiving = _projectService.GetProjectValueArchivingList(ProjectKey).FirstOrDefault();


            var lstProjectArchiving = _projectService.GetProjectValueClosingList(ProjectKey.ToString());

            if (ProjectArchiving != null)
            {
                int statusArchived = 1;
                if(ProjectArchiving.ProjStatus== "Mutipart" || ProjectArchiving.ProjStatus==null)
                {
                    statusArchived = 2;
                } 
                else
                {
                    statusArchived = 1;
                }    
                var update = _projectService.DoArchivingProject(ProjectArchiving.ProjectKey, ProjectArchiving.ServerPath==null?String.Empty: ProjectArchiving.ServerPath, UserId, NoteFromat, 0, statusArchived);
            }

            return RedirectToAction("./ProjectArchiving", "ProjectArchiving");

        }
        [HttpPost, ValidateInput(false)]

        public ActionResult RemoveArchivingProject(string ProjectKey, string Note)
        {


            string NoteFromat = Note.Replace("\r\n", "<br />");

            var count = NoteFromat.Split(new[] { "<br /><br />" }, StringSplitOptions.None).Length - 1;

            if (count == 1)
            {
                NoteFromat = NoteFromat.Replace("<br />", "");
            }

            var ProjectArchiving = _projectService.GetProjectValueArchivingList(ProjectKey).FirstOrDefault();


            var lstProjectArchiving = _projectService.GetProjectValueClosingList(ProjectKey.ToString());

            if (ProjectArchiving != null)
            {
                int statusArchived = 1;
                if (ProjectArchiving.ProjStatus == "Mutipart")
                {
                    statusArchived = 2;
                }
                var update = _projectService.DoArchivingProject(ProjectArchiving.ProjectKey, ProjectArchiving.ServerPath, UserId, NoteFromat, 1, statusArchived);
            }

            return RedirectToAction("./ProjectArchiving", "ProjectArchiving");

        }
        [HttpPost]
        public ActionResult GetDataLoading(int month, int year)
        {
            ViewBag.DaysInSelectMonth = DateTime.DaysInMonth(year, month);          
            try
            {
                return Json(serviceHandler.GetRequestAjax(UserId, month,year), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(null, JsonRequestBehavior.AllowGet);
            }


        }

        public ActionResult ImportFromExcel()
        {
            return View();
        }

        //dong bo archive : import file excel 
        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase postedFile)
        {
            string filePath = string.Empty;
            string path = Server.MapPath("~/Uploads/");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            filePath = path + Path.GetFileName(postedFile.FileName);
            string extension = Path.GetExtension(postedFile.FileName);
            postedFile.SaveAs(filePath);


            string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + filePath + "';Extended Properties='Excel 12.0 Xml; HDR=YES;IMEX=1'";
            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

            //Sheet Name
            excelConnection.Open();
            string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
            excelConnection.Close();
            //End

            OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);

            excelConnection.Open();

            OleDbDataReader dReader;
            dReader = cmd.ExecuteReader();
            SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["strConnectDataString"].ConnectionString);

            //Give your Destination table name
            sqlBulk.DestinationTableName = "ArchiveDataTest";

            //Mappings
            sqlBulk.ColumnMappings.Add("ProjectKey", "ProjectKey");
            sqlBulk.ColumnMappings.Add("ProjectName", "ProjectName");
            sqlBulk.ColumnMappings.Add("ProjStatus", "ProjStatus");
            sqlBulk.ColumnMappings.Add("ServerPath", "ServerPath");
            sqlBulk.ColumnMappings.Add("Note", "Note");
  

            sqlBulk.WriteToServer(dReader);
            excelConnection.Close();

           ViewBag.Result = "Successfully Imported";
        
             return View();
    }


    }
}