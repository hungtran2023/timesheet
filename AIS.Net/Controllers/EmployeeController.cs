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
    public class EmployeeController : Controller

    {

        private readonly IDashBoardService _dashBoardService = Inject.Service<IDashBoardService>();

        private readonly IProjectService _projectService = Inject.Service<IProjectService>();
        private readonly IEmailService _emailService = Inject.Service<IEmailService>();
        private readonly IHREmployeeService hREmployeeService = Inject.Service<IHREmployeeService>();

       private readonly int UserId = 2272;
        private readonly int PageSize = 100;

        private const int _avatarWidth = 105;  // ToDo - Change the size of the stored avatar image
        private const int _avatarHeight = 110;
        private const int _avatarScreenWidth = 500;
        private const string _tempFolder = "/Temp";
        private const string _mapTempFolder = "~" + _tempFolder;
        private const string _avatarPath = "/Data/photos";
        private readonly string[] _imageFileExtensions = { ".jpg", ".png", ".gif", ".jpeg" };
        [HttpGet]
        public ActionResult ProjectRecord(int id = 1416)
        {
            var report = new AIS.Data.Model.ReportProjectHours();

            var headerInformation = new AIS.Data.Model.InformationPerson();

            var listHourSectors = _projectService.GetTotalHourProjectSectors(id, "");

            var listHourServices = _projectService.GetTotalHourProjectServices(id, "");

            if (listHourSectors.Count() > 0)
            {

                var headerItem = listHourSectors.FirstOrDefault();

                if (headerItem != null)
                {
                    headerInformation.FullName = headerItem.FullName;
                    headerInformation.JOBTITLE = headerItem.JOBTITLE;
                    headerInformation.STATTIDDISPLAY = headerItem.STATTIDDISPLAY;
                    headerInformation.STARDATE = headerItem.STARDATE;
                    headerInformation.BIRTHDAY = headerItem.BIRTHDAY;
                    headerInformation.LASTDATE = headerItem.LASTDATE;


                }

                report.information = headerInformation;
                report.totalHourSectors = listHourSectors.ToList();
                report.totalHourServices = listHourServices.ToList();
            }

            return View(report);


        }

        [HttpGet]
        public ActionResult SumaryReportEmployee()
        {
            ViewBag.projectsector = "A";
            ViewBag.projectservice = "0";
            ViewBag.ProjectSectors= _projectService.GetAllSector();
            ViewBag.ProjectServices = _projectService.GetAllServices();

            var report = new AIS.Data.Model.ReportProjectHours();

            var headerInformation = new AIS.Data.Model.InformationPerson();

            var listHourSectors = _projectService.GetTotalHourProjectSectors(-1, "A").OrderByDescending(x=>x.TOTALHOURSECTOR);

            var listHourServices = _projectService.GetTotalHourProjectServices(-1, "").OrderByDescending(x => x.TOTALHOURSERVICECODE); ;

            var sumaryreport = new AIS.Data.Model.ReportSumaryProjects();

            var listSumarayresuts = (from a in listHourSectors
                              join b in listHourServices on a.STATTIDDISPLAY equals b.STATTIDDISPLAY
                              select new ReportSumaryProjectEmployee
                              {
                                  Information = new InformationPerson()
                                  {
                                      FullName = a.FullName,
                                      BIRTHDAY = a.BIRTHDAY,
                                      JOBTITLE = a.JOBTITLE,
                                      DEPARTERMENT = b.DEPARTERMENT,
                                      STARDATE = a.STARDATE,
                                      REPORTTO = a.REPORTTO,
                                      STATTIDDISPLAY = a.STATTIDDISPLAY,
                                      LASTDATE = a.LASTDATE,
                                  },
                                  ListSector = listHourSectors.Where(z => z.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList(),
                                  ListService = listHourServices.Where(y => y.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList()


                              }).GroupBy(x => x.Information.STATTIDDISPLAY).Select(m => new ReportSumaryProjectEmployee()
                              {
                                  Information = m.FirstOrDefault().Information,
                                  ListSector = m.FirstOrDefault().ListSector,
                                  ListService = m.FirstOrDefault().ListService,

                              }).ToList();


            return View(listSumarayresuts);
        }
        
        [HttpPost]
        public ActionResult ChangeProjectSector(string projectsector="", string projectservice="")
        {
            ViewBag.projectsector = projectsector;
            ViewBag.projectservice = projectservice;
            var lstProjectArchiving = _projectService.GetProjectClosingList();
            ViewBag.ProjectSectors = _projectService.GetAllSector();
            ViewBag.ProjectServices = _projectService.GetAllServices();

            var report = new AIS.Data.Model.ReportProjectHours();

            var headerInformation = new AIS.Data.Model.InformationPerson();

            var listHourSectors = _projectService.GetTotalHourProjectSectors(-1, projectsector).OrderByDescending(x => x.TOTALHOURSECTOR);

            var listHourServices = _projectService.GetTotalHourProjectServices(-1, projectservice);

            var sumaryreport = new AIS.Data.Model.ReportSumaryProjects();

            var listSumarayresuts = (from a in listHourSectors
                                     join b in listHourServices on a.STATTIDDISPLAY equals b.STATTIDDISPLAY
                                     select new ReportSumaryProjectEmployee
                                     {
                                         Information = new InformationPerson()
                                         {
                                             FullName = a.FullName,
                                             BIRTHDAY = a.BIRTHDAY,
                                             JOBTITLE = a.JOBTITLE,
                                             DEPARTERMENT = b.DEPARTERMENT,
                                             STARDATE = a.STARDATE,
                                             REPORTTO = a.REPORTTO,
                                             STATTIDDISPLAY = a.STATTIDDISPLAY,
                                             LASTDATE = a.LASTDATE,
                                         },
                                         ListSector = listHourSectors.Where(z => z.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList(),
                                         ListService = listHourServices.Where(y => y.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList()


                                     }).GroupBy(x => x.Information.STATTIDDISPLAY).Select(m => new ReportSumaryProjectEmployee()
                                     {
                                         Information = m.FirstOrDefault().Information,
                                         ListSector = m.FirstOrDefault().ListSector,
                                         ListService = m.FirstOrDefault().ListService,

                                     }).ToList();

            var list = listSumarayresuts;

          
            return PartialView("_SumarySectorEmployee", list);
                
        }
        [HttpPost]
        public ActionResult ChangeProjectService(string projectsector = "", string projectservice = "")
        {
            ViewBag.projectsector = projectsector;
            ViewBag.projectservice = projectservice;
            var lstProjectArchiving = _projectService.GetProjectClosingList();
            ViewBag.ProjectSectors = _projectService.GetAllSector();
            ViewBag.ProjectServices = _projectService.GetAllServices();

            var report = new AIS.Data.Model.ReportProjectHours();

            var headerInformation = new AIS.Data.Model.InformationPerson();

            var listHourSectors = _projectService.GetTotalHourProjectSectors(-1, projectsector).OrderByDescending(x => x.TOTALHOURSECTOR);

            var listHourServices = _projectService.GetTotalHourProjectServices(-1, projectservice);

            var sumaryreport = new AIS.Data.Model.ReportSumaryProjects();

            var listSumarayresuts = (from a in listHourSectors
                                     join b in listHourServices on a.STATTIDDISPLAY equals b.STATTIDDISPLAY
                                     select new ReportSumaryProjectEmployee
                                     {
                                         Information = new InformationPerson()
                                         {
                                             FullName = a.FullName,
                                             BIRTHDAY = a.BIRTHDAY,
                                             JOBTITLE = a.JOBTITLE,
                                             DEPARTERMENT = b.DEPARTERMENT,
                                             STARDATE = a.STARDATE,
                                             REPORTTO = a.REPORTTO,
                                             STATTIDDISPLAY = a.STATTIDDISPLAY,
                                             LASTDATE = a.LASTDATE,
                                         },
                                         ListSector = listHourSectors.Where(z => z.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList(),
                                         ListService = listHourServices.Where(y => y.STATTIDDISPLAY == a.STATTIDDISPLAY).ToList()


                                     }).GroupBy(x => x.Information.STATTIDDISPLAY).Select(m => new ReportSumaryProjectEmployee()
                                     {
                                         Information = m.FirstOrDefault().Information,
                                         ListSector = m.FirstOrDefault().ListSector,
                                         ListService = m.FirstOrDefault().ListService,

                                     }).ToList();

            var list = listSumarayresuts;


            return PartialView("_SumarySectorEmployee", list);

        }
        public ActionResult AtlasStaff()
        {
            ViewBag.LoginPageUrl = StringConstants.AtlasStaffRedirect;
            var listemployees = _dashBoardService.GetAllStaffAtlas().ToList();
            DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/Data/photos"));
            FileInfo[] finfos = di.GetFiles("*.jpg", SearchOption.TopDirectoryOnly);

         
                foreach (var item in listemployees)
                {
                    foreach (FileInfo fi in finfos)
                    {
                        if (Path.GetFileNameWithoutExtension(fi.Name).ToUpper()==item.UserName.ToUpper())
                            {
                                item.PathPhoto = fi.Name;
                    
                            }                  
                }
            }

            foreach (var item in listemployees)
            {
                if(item.PathPhoto==null || item.PathPhoto==String.Empty)
                {
                    if (item.Gender == "F")
                    {
                        item.PathPhoto =  "female.jpg";
                       
                    }
                    else
                    {
                        item.PathPhoto = "male.jpg";
                        
                    }

                }
            }
            ViewBag.FilterSearch = ListItemHelper.GetFillterSearchEmployeeAtlas();
            ViewBag.userId = UserId;
            return View(listemployees);

        }

        private void CleanUpTempFolder(int hoursOld)
        {
            try
            {
                var currentUtcNow = DateTime.UtcNow;
                var serverPath = HttpContext.Server.MapPath(_mapTempFolder);
                if (!Directory.Exists(serverPath)) return;
                var fileEntries = Directory.GetFiles(serverPath);
                foreach (var fileEntry in fileEntries)
                {
                    var fileCreationTime = System.IO.File.GetCreationTimeUtc(fileEntry);
                    var res = currentUtcNow - fileCreationTime;
                    if (res.TotalHours > hoursOld)
                    {
                        System.IO.File.Delete(fileEntry);
                    }
                }
            }
            catch
            {
                // Deliberately empty.
            }
        }
        private static string SaveTemporaryAvatarFileImage(HttpPostedFileBase file, string serverPath, string fileName)
        {
            var img = new WebImage(file.InputStream);
            var ratio = img.Height / (double)img.Width;
            img.Resize(_avatarScreenWidth, (int)(_avatarScreenWidth * ratio));

            var fullFileName = Path.Combine(serverPath, fileName);
           var FileName= Path.GetFileNameWithoutExtension(file.FileName);
            FileName = Path.ChangeExtension(fullFileName, ".JPG");
           


            FileName = Path.GetFileNameWithoutExtension(file.FileName);

            FileName = FileName + ".JPG";
            fullFileName = Path.Combine(serverPath, FileName);

            if (System.IO.File.Exists(fullFileName))
            {
                System.IO.File.Delete(fullFileName);
            }

            img.Save(fullFileName,"JPG",false);
            return Path.GetFileName(img.FileName);
        }
        private string GetTempSavedFilePath(HttpPostedFileBase file)
        {
            // Define destination 
            var serverPath = HttpContext.Server.MapPath(_mapTempFolder); //Request.MapPath(_tempFolder);
            if (Directory.Exists(serverPath) == false)
            {
                Directory.CreateDirectory(serverPath);
            }

            // Generate unique file name
            var fileName = Path.GetFileName(file.FileName);
            //fileName = Path.GetFileNameWithoutExtension(file.FileName);
            //fileName= Path.ChangeExtension(fileName, ".JPG");
           
            fileName = SaveTemporaryAvatarFileImage(file, serverPath, fileName);
           

            fileName = Path.GetFileNameWithoutExtension(file.FileName);

            fileName = fileName + ".JPG";
            // Clean up old files after every save
            CleanUpTempFolder(1);
            return Path.Combine(_mapTempFolder, fileName);
        }
        private bool IsImage(HttpPostedFileBase file)
        {
            if (file == null) return false;
            return file.ContentType.Contains("image") ||
                _imageFileExtensions.Any(item => file.FileName.EndsWith(item, StringComparison.OrdinalIgnoreCase));
        }

        [ValidateAntiForgeryToken]
        public ActionResult _Upload(IEnumerable<HttpPostedFileBase> files)
        {
            if (files == null || !files.Any())
                return Json(new { success = false, errorMessage = "No file uploaded." });

            var file = files.FirstOrDefault();  // get ONE only
            if (file == null || !IsImage(file))
                return Json(new { success = false, errorMessage = "File is of wrong format." });

            if (file.ContentLength <= 0)
                return Json(new { success = false, errorMessage = "File cannot be zero length." });

            var webPath = GetTempSavedFilePath(file);
            var  fileName = Path.GetFileNameWithoutExtension(file.FileName);

            fileName = fileName + ".JPG";
            var index = Request.Url.OriginalString.IndexOf("Employee");
            var str = Request.Url.OriginalString.Remove(index);
            var filenamepath = str + "Temp" + "/"+ fileName;

            return Json(new { success = true, fileName = filenamepath }); // success
        }
        [HttpGet]
        public ActionResult Upload()
        {
            return View();
        }

        [HttpGet]
        public ActionResult _Upload()
        {
            return PartialView();
        }

        [HttpGet]
        public ActionResult UpLoadImages()
        {
            ViewBag.userId = UserId; //_UpLoadAvatarProfile
            return PartialView("_Upload");
        }

        [HttpPost]
        public ActionResult Save(string t, string l, string h, string w, string fileName)
        {
            try
            {
             //   var employee = employeeService.FindById(UserId);

                //if (employee != null)
                //{
                //    string fullPath = string.Empty;

                //    string returnImagePath = string.Empty;                

                //    string Extension;

                //    string imageName;

                //    string imageSavePath;


                //    var usertImage = _dashBoardService.GetAllStaffAtlas().Where(x => x.PersonID == UserId).FirstOrDefault();
                //    // Get file from temporary folder, ...
                //    var fn = Path.Combine(Server.MapPath(_mapTempFolder), Path.GetFileName(fileName));

                //    // ... get the image, ...
                //    var img = new WebImage(fn);

                //    // ... calculate its new dimensions, ...
                //    var height = Convert.ToInt32(h.Replace("-", "").Replace("px", ""));
                //    var width = Convert.ToInt32(w.Replace("-", "").Replace("px", ""));

                //    // ... scale it, ...
                //   img.Resize(width, height);

                //    // ... crop the part the user selected, ...
                //    var top = Convert.ToInt32(t.Replace("-", "").Replace("px", ""));
                //    var left = Convert.ToInt32(l.Replace("-", "").Replace("px", ""));
                //    var bottom = img.Height  -_avatarHeight- top;
                //    var right = img.Width  -_avatarHeight -left;
                //    //var bottom = img.Height - top - _avatarHeight;
                //    //var right = img.Width - left - _avatarWidth;


                //    // ... check for validity of calculations, ...
                //    //if (bottom < 0 || right < 0)
                //    //{
                //    //    // If you reach this point, your avatar sizes in here and in the CSS file are different.
                //    //    // Check _avatarHeight and _avatarWidth in this file
                //    //    // and height and width for #preview-pane .preview-container in site.avatar.css
                //    //    throw new ArgumentException("Definitions of dimensions of the cropping window do not match. Talk to the developer who customized the sample code :)");
                //    //}

                //    img.Crop(top, left, bottom, right);

                //    // ... delete the temporary file,...
                //    System.IO.File.Delete(fn);

                //    // ... and save the new one.

                 
                //   // var newFileName = Path.Combine(_avatarPath, imageName + Extension);


                   
                //    Extension = ".JPG";

                //    imageName = usertImage.UserName;
                //    var path = AppDomain.CurrentDomain.BaseDirectory;
                //    imageSavePath = path + "\\Data\\photos\\" + imageName + Extension;


                //    returnImagePath = path + "\\Data\\photos\\" + imageName + Extension;

                //    var fInfo = new FileInfo(returnImagePath);

                //    fullPath = Request.MapPath("~/Data/photos/" + imageName + Extension);



                //    if (System.IO.File.Exists(fullPath))
                //    {
                //        System.IO.File.Delete(fullPath);
                //    }


                //    img.Save(fullPath);

                  
                //}
                return RedirectToAction("AtlasStaff", "Employee");
            }
            catch (Exception ex)
            {
                return Json(new { success = false, errorMessage = "Unable to upload file.\nERRORINFO: " + ex.Message });
            }
        }

        [HttpPost]
        public ActionResult SaveImageProfile(HttpPostedFileBase filename)
        {

            //if (filename != null && filename.ContentLength > 0)
            //{
            //    var employee = employeeService.FindById(UserId);
            //    if (employee != null)
            //    {
            //        string fullPath = string.Empty;

            //        string returnImagePath = string.Empty;

            //        string fileName;

            //        string Extension;

            //        string imageName;

            //        string imageSavePath;

            //    //    string userId = UserId;

            //        var usertImage = _dashBoardService.GetAllStaffAtlas().Where(x => x.PersonID == UserId).FirstOrDefault();
            //        string fname = filename.FileName;

            //        //   System.Drawing.Image i = resizeImage(filename, new Size(100, 100));

                 

            //        System.Drawing.Bitmap bmpPostedImage = new System.Drawing.Bitmap(filename.InputStream);

            //        System.Drawing.Image objImage = ScaleImage(bmpPostedImage, 130);

            //        if (filename.ContentLength > 0)
            //        {
                      
            //            fileName = Path.GetFileNameWithoutExtension(filename.FileName);
            //            Extension = Path.GetExtension(filename.FileName);
            //            Extension = ".JPG";

            //            imageName = usertImage.UserName;
            //            var path = AppDomain.CurrentDomain.BaseDirectory;
            //            imageSavePath = path + "\\Data\\photos\\" + imageName + Extension;


            //            returnImagePath = path + "\\Data\\photos\\" + imageName + Extension;

            //            var fInfo = new FileInfo(returnImagePath);

            //            fullPath = Request.MapPath("~/Data/photos/" + imageName + Extension);



            //            if (System.IO.File.Exists(fullPath))
            //            {
            //                System.IO.File.Delete(fullPath);
            //            }

            //           // filename.SaveAs(fullPath);
            //            objImage.Save(fullPath);
            //        }

            //    }

            //}
            return RedirectToAction("AtlasStaff", "Employee");
            //   return View("_Position", ViewData["DashBoardWorkRoleForHR"]);
        }
        public static System.Drawing.Image ScaleImage(System.Drawing.Image image, int maxHeight)
        {
            var ratio = (double)maxHeight / image.Height;
            //var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);
            var newImage = new Bitmap(100, newHeight);
            using (var g = Graphics.FromImage(newImage))
            {
                g.DrawImage(image, 0, 0, 100, newHeight);
            }
            return newImage;
        }

        public ActionResult EmployeeDetail()
        {
            ViewBag.PageSize = PageSize;
            ViewBag.EmployeeChooseList = ListItemHelper.GetChooseEmployeeList();
            return View();
        }
        [HttpGet]
        public JsonResult GetEmployeeData(string data = "")
        {

            var lstEmployeeDetails = _projectService.GetAllEmoloyeeDetail();
            List<AIS.Data.Model.EmployeeDetail> list = new List<Data.Model.EmployeeDetail>();

            if (data == "")
            {
                list = lstEmployeeDetails.ToList();
            }
            if (data == StringConstants.AltasEmployee)
            {
                list = lstEmployeeDetails.ToList();
            }

            JsonResult json = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            json.MaxJsonLength = int.MaxValue;

            return json;
        }

        [HttpPost]
        public JsonResult ChangeEmployeeStaff(string employeeChoose)
        {

            var lstProjectArchiving = _projectService.GetAllEmoloyeeDetail();
            List<AIS.Data.Model.EmployeeDetail> list = new List<Data.Model.EmployeeDetail>();
            list = lstProjectArchiving.ToList();

            if (employeeChoose == StringConstants.AltasEmployee)
            {
                list = list.Where(x => x.CharCode == StringConstants.AltasEmployee).ToList();
            }
            if (employeeChoose == StringConstants.TPEmployee)
            {
                list = list.Where(x => x.CharCode == StringConstants.TPEmployee).ToList();
            }

            JsonResult data = Json(new { data = list }, JsonRequestBehavior.AllowGet);
            data.MaxJsonLength = int.MaxValue;
            return data;

        }
        public ActionResult EmployeeProfile()
        {
            return View();
        }
    }
}