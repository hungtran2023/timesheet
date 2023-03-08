using System;
using AIS.Data;
using AIS.Domain.Base;

namespace AIS.Domain.Holiday
{
    public interface IHolidayService : IService<ATC_Holiday>
    {
        DateTime[] GetAll();
    }
}
