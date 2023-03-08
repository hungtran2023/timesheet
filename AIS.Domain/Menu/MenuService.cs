using AIS.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data.Entity;
using AIS.Domain.Employee;
using AIS.Domain.Function;

namespace AIS.Domain.Menu
{
    public class MenuService : IMenuService
    {
        IEmployeeService _employeeService;
        IFunctionService _functionService;
        List<ATC_Functions> listFunctions;

        public MenuService(IEmployeeService employeeService, IFunctionService functionService)
        {
            _employeeService = employeeService;
            _functionService = functionService;
            listFunctions = functionService.FindAll().ToList();
        }

        public List<Menu> GetMenu(int staffId)
        {
            var staff = _employeeService.FindAll().AsQueryable<ATC_Employees>()
                .Include(employee => employee.PersonalInfo.ATC_Users.ATC_Group.Select(group => group.ATC_Permissions.Select(permission => permission.ATC_Functions)))
                .Where(employee => employee.StaffID == staffId).First();
            var result = new List<Menu>();
            var listPermission = new List<ATC_Permissions>();
            var groupList = staff.PersonalInfo.ATC_Users.ATC_Group.AsQueryable<ATC_Group>()
                .Include(t => t.ATC_Permissions.Select(p => p.ATC_Functions)).ToList();
            foreach (var group in groupList)
            {
                if (group.ATC_Permissions.Count > 0)
                {
                    listPermission.AddRange(
                        group.ATC_Permissions.Where(function =>
                         function.ATC_Functions.fgAttribute == false
                         && (function.ATC_Functions.Form.Contains("/")
                         || function.ATC_Functions.Form.EndsWith(".asp"))
                    ).ToList());
                }
            }
            var listDistinctPermission = listPermission.GroupBy(per => per.FunctionID).Distinct().ToList().Select(x => x.First());
            foreach (var item in listDistinctPermission)
            {
                var tempFunction = item.ATC_Functions;
                var link = "";
                if (isLink(tempFunction.Form))
                {
                    link = tempFunction.Form;
                }
                var tempMenu = new Menu(tempFunction.Description, link, tempFunction.GroupID, (int)tempFunction.LevelOrder);
                result.Add(tempMenu);
            }
            return FindParents(result);
        }

        private List<Menu> FindParents(List<Menu> children)
        {
            var result = new List<Menu>();
            bool isNoMoreParent = true;
            while (children.Count > 0)
            {
                foreach (var currentChild in children.ToArray())
                {
                    if (currentChild.ParentId == null)
                    {
                        var menuInResult = result.Where(r => r.Name == currentChild.Name);
                        if (menuInResult.Count() > 0)
                        {
                            var resultItem = menuInResult.First();
                            var resultIndex = result.IndexOf(resultItem);
                            var listOfChildren = resultItem.Children;
                            listOfChildren.AddRange(currentChild.Children);
                            listOfChildren = listOfChildren.OrderBy(t => t.OrderLevel).ToList();
                            currentChild.Children.Clear();
                            currentChild.Children = listOfChildren;
                            result[resultIndex] = currentChild;
                        }
                        else
                        {
                            result.Add(currentChild);
                        }
                        children.Remove(currentChild);
                        break;
                    }
                    var tempParent = listFunctions.Where(function => function.FunctionID == currentChild.ParentId).First();
                    var parentInResult = result.Where(menu => menu.Name == tempParent.Description);
                    var parentAlreadyAdded = false;
                    var parentIndex = 0;
                    Menu tempMenu;
                    if (parentInResult.Count() > 0)
                    {
                        tempMenu = parentInResult.First();
                        parentIndex = result.IndexOf(tempMenu);
                        parentAlreadyAdded = true;
                    }
                    else
                    {
                        var link = "";
                        if (isLink(tempParent.Form))
                            link = tempParent.Form;
                        tempMenu = new Menu(tempParent.Description, link, tempParent.GroupID, (int)tempParent.LevelOrder);
                    }
                    var listOfMenuChildren = children.Where(menu => menu.ParentId == currentChild.ParentId).OrderBy(t => t.OrderLevel);
                    tempMenu.Children.AddRange(listOfMenuChildren);
                    children.RemoveAll(menu => listOfMenuChildren.Contains(menu));
                    if (parentAlreadyAdded)
                    {
                        result[parentIndex] = tempMenu;
                    }
                    else
                    {
                        result.Add(tempMenu);
                    }
                    isNoMoreParent = false;
                    break;
                }
            }
            if (isNoMoreParent)
            {
                var atlasStaff = new Menu("Atlas Staff", "aisnet/Employee/AtlasStaff");
                var completeTimesheetMenu = new Menu("Complete Timesheet", "tms/timesheet.asp");
                var assignMenu = new Menu("Assigned Project", "assignedproject.asp");
                var toolMenu = new Menu("Tools", "");
                toolMenu.Children.Add(new Menu("Preferences", "tools/preferences.asp"));
                toolMenu.Children.Add(new Menu("Change password", "tools/changepassword.asp"));
                toolMenu.Children.Add(new Menu("Forms", "tools/listforms.asp"));
                var logoutMenu = new Menu("Log out", "logout.asp");

                result.Insert(0, atlasStaff);
                result.Insert(1, completeTimesheetMenu);   
                result.Insert(2, assignMenu);
                result.Add(toolMenu);
                result.Add(logoutMenu);
                return result;
            }
            else
            {
                return FindParents(result);
            }
        }

        private bool isLink(String value) {
            return value.EndsWith(".asp") || Regex.IsMatch(value, "/");
        }
    }
}
