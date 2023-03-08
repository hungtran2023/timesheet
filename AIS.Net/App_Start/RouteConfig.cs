using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace AIS
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");


            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "ProjectArchiving", action = "ImportFromExcel", id = UrlParameter.Optional }
            //);
            //   routes.MapRoute(
            //   name: "Default",
            //   url: "{controller}/{action}/{id}",
            //   defaults: new { controller = "MessageContent", action = "UploadFile", id = UrlParameter.Optional }
            //);

            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "HREnterTimesheet", action = "ResetTimeSheet", id = UrlParameter.Optional }
            //);

            //         routes.MapRoute(
            //    name: "Dashboard",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "Utilization", action = "Index", id = UrlParameter.Optional }

            routes.MapRoute(
            name: "Dashboard",
            url: "{controller}/{action}/{id}",
            defaults: new { controller = "Employee", action = "EmployeeDetail", id = UrlParameter.Optional }
        );

            //);

            //            routes.MapRoute(
            // name: "Dashboard",
            // url: "{controller}/{action}/{id}",
            // defaults: new { controller = "Employee", action = "SumaryReportEmployee", id = UrlParameter.Optional }
            //);
            //    routes.MapRoute(
            //    name: "Dashboard",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "Employee", action = "AtlasStaff", id = UrlParameter.Optional }
            //);

            //    routes.MapRoute(
            //    name: "Dashboard",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "DashBoard", action = "BookingTimeSheet", id = UrlParameter.Optional }
            //);

            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "DashBoard", action = "DashBoard", id = UrlParameter.Optional }
            // );


            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "ProjectArchiving", action = "ProjectClosing", id = UrlParameter.Optional }
            //);
            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "ProjectArchiving", action = "ProjectArchiving", id = UrlParameter.Optional }
            //);

            //    routes.MapRoute(
            //    name: "Default",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "Authoriser", action = "Authoriser", id = UrlParameter.Optional }
            //);

            //      routes.MapRoute(
            //   name: "Default1",
            //   url: "{controller}/{action}/{id}",
            //   defaults: new { controller = "HRAuthorisation", action = "HRAuthorisation", id = UrlParameter.Optional }
            //);

            //     routes.MapRoute(
            //    name: "Alias",
            //    url: "{controller}/{action}/{id}",
            //    defaults: new { controller = "AIS/Project", action = UrlParameter.Optional, id = UrlParameter.Optional }
            //  //);

          //  routes.MapRoute(
          //    name: "Default",
          //    url: "{controller}/{action}/{id}",
          //    defaults: new { controller = "HolidayBooking", action = "HolidayBooking", id = UrlParameter.Optional }
          //);
        }
    }
}
