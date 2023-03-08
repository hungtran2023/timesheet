using AIS.Data;
using AIS.Data.Model;
using AIS.Domain.DashBoard.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using static System.Net.Mime.MediaTypeNames;

namespace AIS.Controllers
{
    public class MessageContentController : MenuBaseController
    {
        // GET: MessageContent
        private readonly IDashBoardService _dashBoardService = Inject.Service<IDashBoardService>();
        public ActionResult Index()
        {
            var messageContent = _dashBoardService.GetAllTopMesageContent();
            return View(messageContent);
        }

        public ActionResult Create()
        {
            return View();
        }
        [HttpPost, ValidateInput(false)]
        public ActionResult Create(MessageContentModel messageContentModel)
        {
            var insert = _dashBoardService.InsertMessageContent(messageContentModel.Title, messageContentModel.Description, DateTime.Now, 0);
            return View("Index");
        }

        [HttpPost]
        public ActionResult DeleteConfirm(int Id)
        {
            var insert = _dashBoardService.DeleteMessageContent("", "", DateTime.Now, Id);
            return RedirectToAction("Index");
        }
        public ActionResult Delete(int Id)
        {
            var messageContent = _dashBoardService.GetAllTopMesageContent().ToList().Where(x => x.id == Id).FirstOrDefault();
            return View(messageContent);
        }

        public ActionResult Edit(int Id)
        {
            var messageContent = _dashBoardService.GetAllTopMesageContent().Where(x => x.id == Id).FirstOrDefault();
            return View(messageContent);
        }

        public ActionResult Detail(int Id)
        {
            var messageContent = _dashBoardService.GetAllTopMesageContent().Where(x => x.id == Id).FirstOrDefault();
            return View(messageContent);
        }
        [HttpPost, ValidateInput(false)]
        public ActionResult Update(MessageContentModel messageContentModel)
        {
            var insert = _dashBoardService.UpdateMessageContent(messageContentModel.Title, messageContentModel.Description, DateTime.Now, messageContentModel.id);
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult UploadFileImage(HttpPostedFileBase file, string UserId)
        {
          
            string fullPath = string.Empty;

            string returnImagePath = string.Empty;

            string fileName;

            string Extension;

            string imageName;

            string imageSavePath;

            string userId = UserId;

            var usertImage = _dashBoardService.GetAllStaffAtlas().Where(x => x.PersonID == int.Parse(userId)).FirstOrDefault();


            string fname = file.FileName;

            if (file.ContentLength > 0)
            {
                fileName = Path.GetFileNameWithoutExtension(file.FileName);
                Extension = Path.GetExtension(file.FileName);
                Extension = ".JPG";

                imageName = usertImage.UserName;
                var path = AppDomain.CurrentDomain.BaseDirectory;
                imageSavePath = path + "\\Data\\photos\\" + imageName + Extension;


                returnImagePath = path + "\\Data\\photos\\" + imageName + Extension;

                var fInfo = new FileInfo(returnImagePath);

                fullPath = Request.MapPath("~/Data/photos/" + imageName + Extension);



                if (System.IO.File.Exists(fullPath))
                {
                    System.IO.File.Delete(fullPath);
                }

                file.SaveAs(fullPath);
            }

            string url = "http://vnhcmcode/Timesheet/management/staff/employeeProfile.asp?id="+ UserId;

         

            return Redirect(url);


        }
    



        [HttpGet]
        public ActionResult UploadFile(string UserId)
        {

            //return Json("chamara", JsonRequestBehavior.AllowGet);
            return View();
        }
    }
}