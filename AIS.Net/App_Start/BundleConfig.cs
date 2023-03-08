using System.Web;
using System.Web.Optimization;

namespace AIS
{
    public class BundleConfig
    {
        // For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // Use the development version of Modernizr to develop with and learn from. Then, when you're
            // ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      "~/Scripts/bootstrap.js",
                      "~/Scripts/bootstrap-datepicker.js",
                      "~/Scripts/bootstrap-table.min.js",
                      "~/Scripts/formValidation.min.js",
                      "~/Scripts/bootstrap-datepicker.js",
                      "~/Scripts/bootstrap.min.js", // This one for formValidation
                      "~/Scripts/validator.min.js",
                      "~/Scripts/respond.js",
                      "~/Scripts/moment.js"));

            bundles.Add(new ScriptBundle("~/bundles/menu").Include(
                    "~/Scripts/ais/menu.js"));

            bundles.Add(new ScriptBundle("~/bundles/app").Include(
                    "~/Scripts/ais/app.js"));

            bundles.Add(new ScriptBundle("~/bundles/holiday-booking").Include(
                    "~/Scripts/ais/holiday-booking.js"));

            bundles.Add(new ScriptBundle("~/bundles/overview-holiday").Include(
                   "~/Scripts/ais/overview-of-holiday.js"));

            bundles.Add(new ScriptBundle("~/bundles/authoriser").Include(
                    "~/Scripts/ais/authoriser.js"));

            bundles.Add(new ScriptBundle("~/bundles/hr-authoriser").Include(
                    "~/Scripts/ais/hr-authorisation.js"));

            bundles.Add(new ScriptBundle("~/bundles/hr-timesheet").Include(
                    "~/Scripts/ais/hr-timesheet.js"));


            bundles.Add(new ScriptBundle("~/bundles/avatar").Include(
                   "~/Scripts/Emloyee/site.avatar.js"));

            bundles.Add(new ScriptBundle("~/bundles/jcrop").Include(
                      "~/Scripts/Emloyee/jquery.Jcrop.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryform").Include(
                      "~/Scripts/Emloyee/jquery.form.js"));


            bundles.Add(new StyleBundle("~/Content/jcrop").Include(
                  "~/Content/avatar/jquery.Jcrop.css"));

            bundles.Add(new StyleBundle("~/Content/avatar").Include(
                      "~/Content/avatar/site.avatar.css"));




            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/style/bootstrap.min.css",
                      "~/Content/style/font-awesome.min.css",
                      "~/Content/style/bootstrap-table.min.css",
                      "~/Content/style/formValidation.min.css",
                      "~/Content/style/datepicker.css",
                      "~/Content/style/less/datepicker-custom.css",
                      "~/Content/style/less/base.css",
                      "~/Content/style/less/error.css",
                      "~/Content/style/less/layout.css",
                      "~/Content/style/less/navigation.css",
                      "~/Content/style/less/table.css",
                      "~/Content/style/less/holiday-booking.css",
                      "~/Content/style/less/overview-request-history.css",
                      "~/Content/style/less/authoriser.css",
                      "~/Content/style/less/hr.css",
                      "~/Content/style/less/hr-authorisation.css",
                      "~/Content/style/less/hr-timesheet.css"));
        }


    }
}
