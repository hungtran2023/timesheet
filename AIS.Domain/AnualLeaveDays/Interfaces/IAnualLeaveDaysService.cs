namespace AIS.Domain.AnualLeaveDays
{
    public interface IAnualLeaveDaysService
    {
        double GetCurrentRate();
        double GetAnualLeaveBalance();
        double GetAnualLeaveBalanceLastYear();
        double GetLeaveUntilDay();
        double GetTotalHours();
        double GetAnualLeaveCurrentYear();
        double GetAnualLeaveReserved();
        double GetBalanceHours();
    }
}
