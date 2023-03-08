using System;
using System.Web.Mvc;
using AIS.Data;
using AIS.Domain.Employee;
using AIS.Domain.Preference;
using System.Collections.Generic;
using AIS.Domain.HRCurrentJobTitle;

namespace AIS.Controllers
{
    public class BaseController : Controller
    {
        protected readonly IEmployeeService employeeService = Inject.Service<IEmployeeService>();

        protected int PageSize
        {
            get
            {
                IPreferenceService preferenceService = Inject.Service<IPreferenceService>();
                return preferenceService.GetRowOfPage(UserId);
            }
        }

        protected int UserId
        {
            get
            {
                try
                {
                    return Convert.ToInt16(Session["USERID"].ToString());
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }

        protected string UserFullName
        {
            get
            {
                try
                {
                    return employeeService.GetFullName(LoginUser);
                }
                catch (Exception)
                {
                    return String.Empty;
                }
            }
        }

        protected string UserPosition
        {
            get
            {
                try
                {
                    return Inject.Service<IHRCurrentJobTitle>().Get(UserId);
                }
                catch (Exception)
                {
                    return "";
                }
            }
        }

        protected ATC_Employees LoginUser
        {
            get
            {
                var returnValue = Session["LoginUser"] as ATC_Employees;
                if (returnValue == null || (returnValue != null && returnValue.StaffID != UserId))
                {
                 
                    returnValue = employeeService.FindById(UserId);
                }
                Session["LoginUser"] = returnValue;
                return returnValue;
            }
        }
        protected ActionResult AjaxJsonResult(String message, IEnumerable<Object> data, bool isSuccess) {
            return Json(new { message = message, data = data, isSuccess = isSuccess });
        }
    }
}