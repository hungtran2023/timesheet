using System;
using AIS.Data;
using AIS.Domain.Base;

namespace AIS.Domain.Function
{
    public class FunctionService : Service<ATC_Functions > , IFunctionService
    {
        public FunctionService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }
    }
}
