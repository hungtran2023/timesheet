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
using AIS.Data.Model;
using ProjectService = AIS.Domain.Project.ProjectService;
using System.Web.Util;
using System.EnterpriseServices;
using System.IO;
using AIS.Domain.DashBoard.Interfaces;
using System.Web;
using static System.Net.WebRequestMethods;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.Web.Helpers;
namespace AIS.Controllers
{
    public class UtilizationController : Controller
    {
        // GET: Utilization
        public ActionResult Index()
        {
           var report = new AIS.Data.Model.UtilisationReport();

            var dataGrooups = new AIS.Data.Model.GroupNameReport();
            dataGrooups.Department = "Architects";
            dataGrooups.GroupName = "Architect";            
            dataGrooups.Billablehrs = 1232;
            dataGrooups.OT = 0;
            dataGrooups.EstTraining = 0;
            dataGrooups.BDMdowntime = 0;
            dataGrooups.Projectdowntime = 36;
            dataGrooups.TotalProjectHours = dataGrooups.Billablehrs + dataGrooups.OT + dataGrooups.EstTraining + dataGrooups.BDMdowntime + dataGrooups.Projectdowntime;
            dataGrooups.Atlasproject = 79;
            dataGrooups.GA = 5;
            dataGrooups.Nonprojectdowntime = 240;
            dataGrooups.TotalNonprojects = dataGrooups.Atlasproject + dataGrooups.GA + dataGrooups.Nonprojectdowntime;
            dataGrooups.Availablehours = dataGrooups.TotalProjectHours + dataGrooups.TotalNonprojects - dataGrooups.OT;


            dataGrooups.BillableUtilization = Math.Round(((dataGrooups.Billablehrs + dataGrooups.OT) / dataGrooups.Availablehours) * 100, 2);

            dataGrooups.Utilization = Math.Round((dataGrooups.TotalProjectHours / dataGrooups.Availablehours) * 100, 2);
            report.ReportGroups = new List<GroupNameReport>();
            report.ReportGroups.Add(dataGrooups);

            var dataGrooups2 = new AIS.Data.Model.GroupNameReport();
            dataGrooups2.Department = "Architects";
            dataGrooups2.GroupName = "Architectural Technician";
            dataGrooups2.Billablehrs = 627;
            dataGrooups2.OT = 0;
            dataGrooups2.EstTraining = 0;
            dataGrooups2.BDMdowntime = 0;
            dataGrooups2.Projectdowntime = 0;
            dataGrooups2.TotalProjectHours = dataGrooups2.Billablehrs + dataGrooups2.OT + dataGrooups2.EstTraining + dataGrooups2.BDMdowntime + dataGrooups2.Projectdowntime;
            dataGrooups2.Atlasproject = 0;
            dataGrooups2.GA = 0;
            dataGrooups2.Nonprojectdowntime = 149;
            dataGrooups2.TotalNonprojects = dataGrooups2.Atlasproject + dataGrooups2.GA + dataGrooups2.Nonprojectdowntime;
            dataGrooups2.Availablehours = dataGrooups2.TotalProjectHours + dataGrooups2.TotalNonprojects - dataGrooups2.OT;

            dataGrooups2.BillableUtilization = Math.Round(((dataGrooups2.Billablehrs + dataGrooups2.OT)/ dataGrooups2.Availablehours)*100,2);

            dataGrooups2.Utilization =  Math.Round((dataGrooups2.TotalProjectHours / dataGrooups2.Availablehours)*100,2);
            report.ReportGroups.Add(dataGrooups2);

            ViewBag.UtilizationReport = report;
            return View();
        }
    }
}