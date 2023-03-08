using System;
using System.Configuration;
using System.Web;
using System.Web.Mvc;
using MSDN;
using AIS.Domain.Common.Constants;

namespace AIS.ActionFilters
{
    public class LegacyAuthorize : AuthorizeAttribute
    {
        private HttpCookie cookie;
        private string CONNECTION_STRING = StringConstants.SessionDSN;
        private string TIME_OUT = StringConstants.SessionTimeOut;
        private string AIS_SITE = StringConstants.AISSiteUrl;
        private int DEFAULT_TIMEOUT = NumberConstants.SessionTimeOut;
        private ISessionPersistence sessionPersistence = new SessionPersistence();

        private int SessionExpiration
        {
            get
            {
                if (ConfigurationSettings.AppSettings[TIME_OUT] != null)
                    return Convert.ToInt32(ConfigurationSettings.AppSettings[TIME_OUT]);
                else
                    return DEFAULT_TIMEOUT;
            }
        }
        private string dsn
        {
            get
            {
                return ConfigurationSettings.AppSettings[CONNECTION_STRING];
            }
        }

        public override void OnAuthorization(AuthorizationContext filterContext)
        {

            SetSession();

            if (string.IsNullOrEmpty(HttpContext.Current.Session[StringConstants.UserId] as string))
            {
                HandleUnauthorizedRequest(filterContext);
            }
        }

        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            //HttpContext.Current.Session[StringConstants.UserId] = "252";
            if (filterContext.HttpContext.Request.IsAjaxRequest())
            {
                filterContext.Result = new JsonResult { Data = StringConstants.SessionExpired };
            }else if (filterContext.HttpContext.Request.QueryString["fromEmail"] == "true")
            {
                var authoriserUrl = StringConstants.AuthoriserFromEmailURL;
                var baseUrl = StringConstants.BaseUrl;
                authoriserUrl = authoriserUrl.Replace("/", StringConstants.FlashCode);
                var targetUrl = baseUrl + "?ReturnUrl=" + authoriserUrl;
                filterContext.HttpContext.Response.Redirect(targetUrl);
            }
            else if (filterContext.HttpContext.Request.QueryString["data"] == "1")
            {
                var authoriserUrl = StringConstants.ProjectClosingEmail;
                var baseUrl = StringConstants.BaseUrl;
                authoriserUrl = authoriserUrl.Replace("/", StringConstants.FlashCode);
                var targetUrl = baseUrl + "?ReturnUrl=" + authoriserUrl + "?data=1";
                filterContext.HttpContext.Response.Redirect(targetUrl);
            }
            else if (filterContext.HttpContext.Request.RequestContext.RouteData.Values["controller"].ToString() == "Employee")
            {
              
                var authoriserUrl = StringConstants.AtlasStaffRedirect;
                var baseUrl = StringConstants.BaseUrl;
                authoriserUrl = authoriserUrl.Replace("/", StringConstants.FlashCode);
                var targetUrl = baseUrl + "?ReturnUrl=" + authoriserUrl;
                filterContext.HttpContext.Response.Redirect(targetUrl);
            }
            else
            {
                base.HandleUnauthorizedRequest(filterContext);
            }
        }

        private void SetSession()
        {

            HttpContext.Current.Session[StringConstants.UserId] = "2272";
            return;
            ////////////////////////////////////////////////////////////

            mySession session = new mySession();
            cookie = HttpContext.Current.Request.Cookies[sessionPersistence.SessionID];

            if (cookie == null)
            {
                HttpContext.Current.Session[StringConstants.UserId] = null;
                return;
            }

            session = sessionPersistence.LoadSession(HttpUtility.UrlDecode(cookie.Value).ToLower().Trim(), dsn, SessionExpiration);
         
            if (session != null)
            {

                HttpContext.Current.Session[StringConstants.UserId] = session[StringConstants.UserId];
                HttpContext.Current.Session[StringConstants.CurrentRate] = session[StringConstants.CurrentRate];
                HttpContext.Current.Session[StringConstants.BalanceDay] = session[StringConstants.BalanceDay];
                HttpContext.Current.Session[StringConstants.BalanceLastYear] = session[StringConstants.BalanceLastYear];
                HttpContext.Current.Session[StringConstants.LeaveUntilDay] = session[StringConstants.LeaveUntilDay];
                HttpContext.Current.Session[StringConstants.TotalHours] = session[StringConstants.TotalHours];
                HttpContext.Current.Session[StringConstants.AnnualLeaveCurrentYear] = session[StringConstants.AnnualLeaveCurrentYear];
                HttpContext.Current.Session[StringConstants.AnnualLeaveReserved] = session[StringConstants.AnnualLeaveReserved];
                HttpContext.Current.Session[StringConstants.BalanceHours] = session[StringConstants.BalanceHours];
                if(session[StringConstants.UserId] != String.Empty)
                {
                    sessionPersistence.SaveSession(HttpUtility.UrlDecode(cookie.Value).ToLower().Trim(), dsn, session, false);
                }
                HttpContext.Current.Response.SetCookie(new HttpCookie(sessionPersistence.SessionID, cookie.Value));
            }
        }
    }
}