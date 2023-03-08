using AIS.Domain.Common.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace AIS.Controllers
{
    public class ErrorController : Controller
    {
        public ViewResult Index()
        {
            ViewBag.Status = Response.StatusCode = 200;
            return View(StringConstants.Error);
        }

        public ViewResult NotFound()
        {
            ViewBag.Status = Response.StatusCode = 404;
            return View(StringConstants.Error);
        }

        public ViewResult InternalServerError()
        {
            ViewBag.Status = Response.StatusCode = 500;
            return View(StringConstants.Error);
        }

        public ViewResult Forbidden()
        {
            ViewBag.Status = Response.StatusCode = 403;
            return View(StringConstants.Error);
        }
    }
}