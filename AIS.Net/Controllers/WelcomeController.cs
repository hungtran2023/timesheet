using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace AIS.Controllers
{
    public class WelcomeController :MenuBaseController
    {
        // GET: Welcome
        public ActionResult Index()
        {
            return View();
        }
    }
}