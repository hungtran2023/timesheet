using AIS.Data;
using AIS.Domain.Base;
using AIS.Domain.Common.Constants;
using System;

namespace AIS.Domain.Preference
{
    public class PreferenceService : Service<ATC_Preferences>, IPreferenceService
    {
        public PreferenceService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public int GetRowOfPage(int StaffId)
        {
            var result = this.FindByCriteria(t => t.Numofrows != null && t.StaffID == StaffId);
            if (result != null)
            {
                return (int)result.Numofrows;
            }
            return NumberConstants.NumOfRowsOnTableDefault;
        }
    }
}
