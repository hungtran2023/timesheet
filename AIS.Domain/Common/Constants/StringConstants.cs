using AIS.Domain.Common.Helper;

namespace AIS.Domain.Common.Constants
{
    public class StringConstants
    {
        public static readonly string UserEmailDefault = ConfigurationHelper.GetValueConfig("UsernameEmail");
        public static readonly string PassEmailDefault = ConfigurationHelper.GetValueConfig("PasswordEmail");
        public static readonly string BaseUrl = ConfigurationHelper.GetValueConfig("BaseUrl");
        public static readonly string LogOutURL = BaseUrl + ConfigurationHelper.GetValueConfig("LogOutURL");
        public static readonly string AuthoriserFromEmailURL = ConfigurationHelper.GetValueConfig("AuthoriserURL");
        public static readonly string TimeSheetListURL = BaseUrl + ConfigurationHelper.GetValueConfig("TimeSheetListURL");
        public static readonly string AuthoriserURL = BaseUrl + ConfigurationHelper.GetValueConfig("AuthoriserURL");
        public static readonly string AuthoriserRedirectInEmailURL = BaseUrl + ConfigurationHelper.GetValueConfig("AuthoriserURL") + "?fromEmail=true";
        public static readonly string TimesheetURL = BaseUrl + ConfigurationHelper.GetValueConfig("TimesheetUrl");
        public static readonly string WelcomeURL = BaseUrl + ConfigurationHelper.GetValueConfig("WelcomeURL");
        public static readonly string PreferenceURL = BaseUrl + ConfigurationHelper.GetValueConfig("PreferenceURL");
        public static readonly string MessageURL = BaseUrl + ConfigurationHelper.GetValueConfig("MessageURL");

        public static readonly string SessionDSN = "SessionDSN";
        public static readonly string SessionTimeOut = "SessionTimeOut";
        public static readonly string AISSiteUrl = "AISSiteURL";
        public static readonly string UserId = "USERID";
        public static readonly string SessionExpired = "SessionExpired";
        public static readonly string CurrentRate = "CURRENT_RATE";
        public static readonly string BalanceDay = "BALANCE_DAY";
        public static readonly string BalanceLastYear = "BALANCE_LAST_YEAR";
        public static readonly string LeaveUntilDay = "LEAVE_UNTIL_DAY";
        public static readonly string TotalHours = "TOTAL_HOURS";
        public static readonly string AnnualLeaveCurrentYear = "ANUAL_LEAVE_CURRENT_YEAR";
        public static readonly string AnnualLeaveReserved = "ANUAL_LEAVE_RESERVED";
        public static readonly string BalanceHours = "BALANCE_HOURS";

        public static readonly string EmailRequestApproval = "Email Request Approval";
        public static readonly string EmailInformApproveByAuthorizor1 = "Email Inform Approve By Authorizor1";
        public static readonly string EmailInformApproveByAnotherAuthorizor = "Email Inform Approve By Authorizor2";
        public static readonly string EmailProjectArchiving = "Email Project Archiving";
        public static readonly string EmailReject = "Email Reject";
        public static readonly string FlashCode = "%2f";
        public static readonly string DateOnlyFormat = "dd/MM/yyyy";
        public static readonly string Error = "Error";
        public static readonly string HourOnlyFormat = "HH:mm";
        public static readonly string DateTimeFormat = "dd/MM/yyyy HH:mm";
        public static readonly string Zero = "0";
        public static readonly string NewStatus = "New";
        public static readonly string AuthorisedStatus = "Authorised";
        public static readonly string TakenStatus = "Taken";
        public static readonly string ErrorMessage = "Error occurs.";
        public static readonly string TableTimesheet = "AIS.Data.ATC_Timesheet";
        public static readonly string RequestDeleteMessage = "The request is deleted successfully.";
        public static readonly string RequestAddMessage = "The request is added successfully.";
        public static readonly string RequestUpdateMessage = "The request is updated successfully.";
        public static readonly string RequestRejectMessage = "Request has been rejected.";
        public static readonly string RequestApproveMessage = "Request has been approved.";
        public static readonly string RequestErrorLessJoinDate = "Start date is before join date.";
        public static readonly string RequestErrorGreaterLeaveDate = "End date is after leave date.";
        public static readonly string RequestErrorStartTimeInvalid = "Start off time is invalid.";
        public static readonly string RequestErrorEndTimeInvalid = "End off time is invalid.";
        public static readonly string RequestOf = "The request of {0} ";
        public static readonly string RequestErrorNoWorkDays = "There are no workday between date from and date to.";
        public static readonly string RequestErrorAlreadyMade = "The request on {0} has already made. Please contact HR department for help.";
        public static readonly string TimesheetAddMessage = "Timesheet is added successfully!";
        public static readonly string TimesheetErrorFull = "the timesheet on {0} has already completed by another. Please contact HR department for help.\n";
        public static readonly string TimesheetNoTable = "No Timesheet table of next year";
        public static readonly string TimesheetAutomaticallyInserted = "Enter automatically by system ID: {0}";

        public static readonly int ProjectStatusClose = 3;
        public static readonly string yourProjects = "2";
        public static readonly string noinputdataProjects = "1";

        public static readonly string Close = "Close";
        public static readonly string Live = "Live";
        public static readonly string Archive = "Archive";

        public static readonly string AltasEmployee = "ATL";
        public static readonly string TPEmployee = "BPO";

        public static readonly string EmailITSupport = ConfigurationHelper.GetValueConfig("EmailITSupport");
        public static readonly string EmailCSOArchiving = ConfigurationHelper.GetValueConfig("EmailCSOArchiving");
        public static readonly string EmailDevTest = ConfigurationHelper.GetValueConfig("EmailDevTest");
        public static readonly string ProjectClosingEmail = ConfigurationHelper.GetValueConfig("ProjectClosingEmail");
        public static readonly string ProjectClosingEmailRedirect = BaseUrl + ConfigurationHelper.GetValueConfig("ProjectClosingEmail") + "?data=1";
        
        public static readonly string Pathphoto =  ConfigurationHelper.GetValueConfig("PathPhoto");

        public static readonly string AtlasStaffRedirect = BaseUrl + ConfigurationHelper.GetValueConfig("AtlasStaffRederect");




    }
}
