using AIS.Data;
using AIS.Domain.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Domain.Preference
{
    public interface IPreferenceService : IService<ATC_Preferences>
    {
        int GetRowOfPage(int StaffId);
    }
}
