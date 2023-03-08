using AIS.Domain.Menu;
using System.Web.Mvc;
using AIS.Domain.HRCurrentJobTitle;

namespace AIS.Controllers
{
    public class MenuBaseController : BaseController
    {
        IMenuService _menuService = Inject.Service<IMenuService>();
        IHRCurrentJobTitle _hrCurrentJobTitle = Inject.Service<IHRCurrentJobTitle>();
        protected override void OnActionExecuted(ActionExecutedContext filterContext)
        {            
            var staff = employeeService.FindById(UserId);
            ViewBag.Job = _hrCurrentJobTitle.Get(UserId);
            ViewBag.UserName = staff.PersonalInfo.ATC_Users.UserName;
            ViewBag.Name = employeeService.GetFullName(staff);
            ViewBag.Menu = _menuService.GetMenu(UserId);
            base.OnActionExecuted(filterContext);
        }
    }
}