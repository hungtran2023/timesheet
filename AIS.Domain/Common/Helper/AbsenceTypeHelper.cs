using System;
using AIS.Domain.Common.Enum;

namespace AIS.Domain.Common.Helper
{
    public static class AbsenceTypeHelper
    {
        public static String ConvertAbsenceType(AbsenceType type) {
            var result = "";
            switch (type)
            {
                case AbsenceType.AnnualHoliday:
                    result = "Annual Holiday";
                    break;
                case AbsenceType.SickLeave:
                    result = "Sick Leave";
                    break;
                case AbsenceType.OtherLeave:
                    result = "Other Leave";
                    break;
                case AbsenceType.UnpaidLeave:
                    result = "Unpaid Leave";
                    break;
                case AbsenceType.SickLeaveWithCertificate:
                    result = "Sick Leave with certificate";
                    break;
                default:
                    break;
            }
            return result;
        }

        public static String ConvertAbsenceStatus(AbsenceStatus type )
        {
            var result = "";
            switch (type)
            {
                case AbsenceStatus.New:
                    result = "New";
                    break;
                case AbsenceStatus.InProgress:
                    result = "In-Progress";
                    break;
                case AbsenceStatus.Rejected:
                    result = "Rejected";
                    break;
                case AbsenceStatus.Authorised:
                    result = "Authorised";
                    break;
                case AbsenceStatus.Taken:
                    result = "Taken";
                    break;
                case AbsenceStatus.UnAuthorised:
                    result = "UnAuthorised";
                    break;
                default:
                    break;
            }
            return result;
        }

        public static String ConvertAbsenceStatus(AbsenceStatus type,DateTime date)
        {
            var result = "";
            switch (type)
            {
                case AbsenceStatus.New:
                    result = "New";
                    break;
                case AbsenceStatus.InProgress:
                    result = "In-Progress";
                    break;
                case AbsenceStatus.Rejected:
                    result = "Rejected";
                    break;
                case AbsenceStatus.Authorised:
                    result = "Authorised";
                    if (date < DateTime.Now)
                    {
                        result = "Taken";
                    }
                    break;
                default:
                    break;
            }
            return result;
        }

        public static int? GetStatusId(AbsenceStatus type, DateTime date)
        {
            int? result = null;
            switch (type)
            {
                case AbsenceStatus.New:
                    result = (int)AbsenceStatus.New;
                    break;
                case AbsenceStatus.InProgress:
                    result = (int)AbsenceStatus.InProgress;
                    break;
                case AbsenceStatus.Rejected:
                    result = (int)AbsenceStatus.Rejected;
                    break;
                case AbsenceStatus.Authorised:
                    result = (int) AbsenceStatus.Authorised;
                    if (date < DateTime.Now)
                    {
                        result = (int) AbsenceStatus.Taken;
                    }
                    break;
                default:
                    break;
            }
            return result;
        }
    }
}