using System;
using System.Collections.Generic;
using System.Linq;
using AIS.Data;
using AIS.Domain.Base;
using System.Data.Entity;

namespace AIS.Domain.Employee
{
    public class EmployeeService : Service<ATC_Employees> , IEmployeeService
    {
        public EmployeeService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }
        public override ATC_Employees FindById(int id)
        {
            return FindAll().AsQueryable().Include(t => t.PersonalInfo).Where(t => t.StaffID == id).First();
        }
        public ATC_Employees FindLogiUser(int id) {
            return FindAll().AsQueryable().Include(t => t.ATC_JobTitle).Where(t => t.StaffID == id).First();
        }
        public decimal GetWorkingTime(DateTime date, int staffId)
        {
            var listSalaryStatus = FindAll().Where(t => t.StaffID == staffId).AsQueryable().Include(t => t.ATC_SalaryStatus).First().ATC_SalaryStatus;
            return listSalaryStatus.Where(status => status.SalaryDate.Date <= date.Date).OrderByDescending(targets => targets.SalaryDate).First().ATC_WorkingHours.Hours;
        }

        public decimal CurrentWorkingHours(int StaffId)
        {
            var results = new List<ATC_SalaryStatus>();
            var listSalaryStatus = FindById(StaffId).ATC_SalaryStatus;
            if (listSalaryStatus.Count == 0)
            {
                return 0;
            }
            return listSalaryStatus.Last().ATC_WorkingHours.Hours;
        }

        

        public String GetFullName(int userId) {
            var user =  FindAll().Where(t => t.StaffID == userId).AsQueryable().Include(t => t.PersonalInfo).First()  ;
            return user != null ? user.PersonalInfo.FirstName + " " + user.PersonalInfo.LastName : String.Empty;
        }

        public String GetFullName(ATC_Employees user)
        {
            return user != null ? user.PersonalInfo.FirstName + " " + user.PersonalInfo.LastName : String.Empty;
        }

        public IEnumerable<ATC_SalaryStatus> GetSalaryStatusBetween(DateTime from, DateTime to, int StaffId)
        {
            var results = new List<ATC_SalaryStatus>();
            var listSalaryStatus = FindAll().Where(staff => staff.StaffID == StaffId).AsQueryable().Include(t => t.ATC_SalaryStatus).First().ATC_SalaryStatus;
            if (listSalaryStatus.Count == 0)
            {
                return results;
            }
            var checkIfManySalaryStatus = listSalaryStatus.Where(targets => targets.SalaryDate.Date >= from.Date && targets.SalaryDate.Date <= to.Date)
                .OrderByDescending(targets => targets.SalaryDate).ToList();
            if (checkIfManySalaryStatus.Count() == 0)
            {
                results.Add(listSalaryStatus.Where(before => before.SalaryDate.Date <= from.Date).OrderByDescending(targets => targets.SalaryDate).First());
            }
            else
            {
                var lastSalaryStatus = checkIfManySalaryStatus.Last();
                if (lastSalaryStatus.SalaryDate.Date > from.Date)
                {
                    var salaryStatusBeforeDateFrom = listSalaryStatus.Where(target => target.SalaryDate < lastSalaryStatus.SalaryDate)
                                .OrderByDescending(target => target.SalaryDate).First();
                    results.Add(salaryStatusBeforeDateFrom);
                }
                results.AddRange(checkIfManySalaryStatus.OrderBy(target => target.SalaryDate));
            }
            return results;
        }
        public IEnumerable<ATC_Employees> StaffsForManager(int ManagerId)
        {
            var listOfStaffs = new List<ATC_Employees>();
            int[] managerArray = { ManagerId };
            var listOfAllEmployee = FindAll().AsQueryable().Include(t => t.PersonalInfo).Include(t => t.LeaderOfStaffs).ToList();
            InitializeStaffsForManager(managerArray , listOfAllEmployee, ref listOfStaffs);
            return listOfStaffs.OrderBy(t => t.PersonalInfo.FirstName);
        }
        private void InitializeStaffsForManager(int[] managersId ,List<ATC_Employees> listOfAllEmployee , ref List<ATC_Employees> staffsList)
        {
            var tempListOfManagers = new List<int>();
            foreach (var item in managersId)
            {
                var currentStaffList = listOfAllEmployee.Where(staff => staff.DirectLeaderID == item && staff.PersonalInfo.fgDelete == false && !staff.PersonalInfo.FirstName.Contains("Manager"));
                if (currentStaffList.Count() > 0)
                {
                    foreach (var staff in currentStaffList)
                    {
                        staffsList.Add(staff);
                        if (staff.LeaderOfStaffs.Count() > 0)
                        {
                            tempListOfManagers.Add(staff.StaffID);
                        }
                    }
                }
            }
            if (tempListOfManagers.Count() == 0)
            {
                return;
            }
            else
            {
                InitializeStaffsForManager(tempListOfManagers.ToArray(), listOfAllEmployee, ref staffsList);
            }
        }
    }
}
