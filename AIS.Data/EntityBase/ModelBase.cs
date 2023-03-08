using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data
{
    public partial class HR_CurrentJobtitle : Entity
    {
        public override int ObjId { get { return this.ObjId; } }
    }
    public partial class ATC_Department : Entity
    {
        public override int ObjId { get { return this.DepartmentID; } }
    }
    public partial class ATC_Employees : Entity
    {
        public override int ObjId { get { return this.StaffID; } }
    }

    public partial class HR_ReceiveReport : Entity
    {
        public override int ObjId { get { return this.UserID; } }
    }

    public partial class ATC_AbsenceRequests : Entity
    {
        public override int ObjId { get { return this.Id; } }
    }

    public partial class ATC_Holiday : Entity
    {
        public override int ObjId { get { return this.HolidayID; } }
    }

    public partial class ATC_Timesheet : Entity
    {
        public override int ObjId { get { return this.StaffID; } }
    }
    public partial class ATC_Functions : Entity
    {
        public override int ObjId { get { return this.FunctionID; } }
    }
    public partial class ATC_Events : Entity
    {
        public override int ObjId { get { return this.EventID; } }
    }
    public partial class ATC_EmailTemplate : Entity
    {
        public override int ObjId { get { return this.Id; } }
    }
    public partial class HR_Employee : Entity
    {
        public override int ObjId { get { return this.PersonID; } }
    }
    public partial class ATC_Preferences : Entity
    {
        public override int ObjId { get { return this.StaffID; } }
    }

    public partial class ATC_ProjectTracking : Entity
    {
        public override int ObjId { get { return this.ProjTrackerID; } }
    }

}
