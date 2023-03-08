using AIS.Data;
using AIS.Domain.Base;
using System;

namespace AIS.Domain.Email.Interfaces
{
    public interface IEmailTemplateService :IService<ATC_EmailTemplate>
    {
        ATC_EmailTemplate GetEmailTemplateByType(string type);
    }
}
