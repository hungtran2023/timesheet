namespace AIS.Domain.Common.Enum
{
    public enum AbsenceType
    {
        AnnualHoliday = 4,
        SickLeave = 6,
        OtherLeave = 7,
        UnpaidLeave = 8,
        SickLeaveWithCertificate = 9,

    }

    public enum AbsenceStatus
    {
        New = 0 ,
        InProgress = 1,
        Rejected = 2 ,
        Authorised = 3,
        Taken = 4 ,
        UnAuthorised = 5
    }


    public enum ChooseProject
    {
        YourProject = 0,
        InProgress = 1,
        Rejected = 2,
        Authorised = 3,
        Taken = 4,
        UnAuthorised = 5
    }
}