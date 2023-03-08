using System;
using System.Linq;
using SimpleInjector;
using SimpleInjector.Integration.Web;
using AIS.Data;
using AIS.Domain.Base;
using AIS.Data.StoredProcedures;
using AIS.Domain.AnualLeaveDays;
using AIS.Domain.Menu;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.Email;
using AIS.Domain.Employee;
using AIS.Domain.Function;
using AIS.Domain.Holiday;
using AIS.Domain.HREmployee;
using AIS.Domain.HRReport;
using AIS.Domain.Event;
using AIS.Domain.Department;
using AIS.Domain.AbsenceRequest;
using AIS.Domain.TimeSheet;
using AIS.Domain.Preference;
using AIS.Domain.Project;
using AIS.Domain.HRCurrentJobTitle;
using AIS.Data.EntityBase.StoredProcedures;
using AIS.Domain.DashBoard.Interfaces;
using AIS.Domain.DashBoard;

namespace AIS
{
    public class Inject
    {
        private static Container _container;
        public static void Register()
        {
            var container = new Container();
            var webLifestyle = new WebRequestLifestyle();
            container.RegisterPerWebRequest<LeaveManagementContext>(() => new LeaveManagementContext());
         
            container.RegisterPerWebRequest<IUnitOfWork>(() => new UnitOfWork(container.GetInstance<LeaveManagementContext>()));
         

            container.Register<IHRCurrentJobTitle>(() => new HRCurrentJobTitle(container.GetInstance<IUnitOfWork>()));
            container.Register<IEmployeeService>(() => new EmployeeService(container.GetInstance<IUnitOfWork>()));
            container.Register<IEventsService>(() => new EventsService(container.GetInstance<IUnitOfWork>()));
            container.Register<IFunctionService>(() => new FunctionService(container.GetInstance<IUnitOfWork>()));
            container.Register<IHolidayService>(() => new HolidayService(container.GetInstance<IUnitOfWork>()));
            container.Register<IHREmployeeService>(() => new HREmployeeService(container.GetInstance<IUnitOfWork>()));
            container.Register<IHRReceiveReportService>(() => new HRReceiveReportService(container.GetInstance<IUnitOfWork>()));
            container.Register<IDepartmentService>(() => new DepartmentService(container.GetInstance<IUnitOfWork>()));
            container.Register<IAbsenceRequestService>(() => new AbsenceRequestService(container.GetInstance<IUnitOfWork>()));
            container.Register<ITimeSheetService>(() =>new TimeSheetService(container.GetInstance<IUnitOfWork>()));
            container.Register<IPreferenceService>(() => new PreferenceService(container.GetInstance<IUnitOfWork>()));
            container.Register<IEmailTemplateService>(() => new EmailTemplateService(container.GetInstance<IUnitOfWork>()));          
            container.Register<IMenuService>( () => new MenuService(container.GetInstance<IEmployeeService>(),container.GetInstance<IFunctionService>()));
            container.Register<IEmailService>( () => new EmailService(container.GetInstance<IHREmployeeService>(), container.GetInstance<IEmailTemplateService>()));


            container.Register<AbsenceRequestHandler>(() => new AbsenceRequestHandler(
                container.GetInstance<IEmployeeService>(),
                container.GetInstance<IAbsenceRequestService>(),
                container.GetInstance<ITimeSheetService>(),
                container.GetInstance<IEmailService>(),
                container.GetInstance<IHolidayService>(),
                container.GetInstance<IHREmployeeService>()
            ));
       
            container.Register<IAnualLeaveDaysService>(() => new AnualLeaveDaysService());
            //Register Service to call Store procedure
            container.Register<IAtlasStoredProcedures, AtlasStoreProcedureContext>();
            container.Register<IProjectService>(() => new ProjectService(container.GetInstance<IAtlasStoredProcedures>()));
            container.Register<IDashBoardService>(() => new DashBoardService(container.GetInstance<IAtlasStoredProcedures>()));
            container.Verify();
            _container = container;
        }

        public static TService Service<TService>()
            where TService : class
        {
            return _container.GetInstance<TService>();
        }
    }
}