using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class TimsheetReportProjectSector
    {

        public string FullName { get; set; }
        public string BIRTHDAY { get; set; }
        public string JOBTITLE { get; set; }
        public string STARDATE { get; set; }
        public string LASTDATE { get; set; }
        public string STATTIDDISPLAY { get; set; }
        public int PERSONID { get; set; }
        public string DEPARTERMENT { get; set; }
        public string REPORTTO { get; set; }
        public decimal TOTALAVAVIBLES { get; set; }
        public decimal TOTALHOURSECTOR { get; set; }
        public decimal PECENTAGESECTOR { get; set; }
        public string SECTORTYPE { get; set; }
        public string SECTORNAME { get; set; }
    }

    public class TimsheetReportProjectServiceType
    {
        public string FullName { get; set; }
        public string BIRTHDAY { get; set; }
        public string JOBTITLE { get; set; }
        public string STARDATE { get; set; }
        public string LASTDATE { get; set; }
        public string STATTIDDISPLAY { get; set; }
        public int PERSONID { get; set; }
        public string DEPARTERMENT { get; set; }
        public string REPORTTO { get; set; }
        public decimal TOTALAVAVIBLES { get; set; }
        public decimal TOTALHOURSERVICECODE { get; set; }
        public decimal PECENTAGESERVICECODE { get; set; }
        public string SERVICECODE { get; set; }
        public string SERVICENAME { get; set; }
    }

    public class InformationPerson
    {
        public string FullName { get; set; }
        public string BIRTHDAY { get; set; }
        public string JOBTITLE { get; set; }
        public string STARDATE { get; set; }
        public string STATTIDDISPLAY { get; set; }
        public int PERSONID { get; set; }
        public string LASTDATE { get; set; }
        public string DEPARTERMENT { get; set; }
        public string REPORTTO { get; set; }
    }
    public class ReportProjectHours
    {
        public InformationPerson information { get; set; }
        public List<TimsheetReportProjectServiceType> totalHourServices { get; set; }
        public List<TimsheetReportProjectSector> totalHourSectors { get; set; }

        public ReportProjectHours()
        {

        }

    }
    public class ReportSumaryProjects
    {
        public List<InformationPerson> informations { get; set; }
        public List<TimsheetReportProjectServiceType> totalHourServices { get; set; }
        public List<TimsheetReportProjectSector> totalHourSectors { get; set; }
        public ReportSumaryProjects()
        {

        }
    }
    public class ProjectService
    {


        public decimal TOTALHOURSERVICECODE { get; set; }
        public decimal PECENTAGESERVICECODE { get; set; }
        public string SERVICECODE { get; set; }
        public string SERVICENAME { get; set; }

        public ProjectService()
        {

        }

    }
    public class ProjectSector
    {

        public decimal TOTALHOURSECTOR { get; set; }
        public decimal PECENTAGESECTOR { get; set; }
        public string SECTORTYPE { get; set; }
        public string SECTORNAME { get; set; }

        public ProjectSector()
        {

        }
    }

    public class ReportSumaryProjectEmployee
    {       
        public  InformationPerson  Information {get;set;}
        public List<TimsheetReportProjectSector> ListSector { get; set; }
        public List<TimsheetReportProjectServiceType> ListService { get; set; }

    }

    public class ReportSumaryEmployee
    {                    
        public List<ProjectSector> ListSector { get; set; }
        public List<ProjectService> ListService { get; set; }
       
    }
    public class ListSectors
    {
        public string SectorCode { get; set; }

        public string SectorName { get; set; }

    }

    public class ListServices
    {
        public string ServiceCode { get; set; }

        public string SeriveName { get; set; }

    }



}
