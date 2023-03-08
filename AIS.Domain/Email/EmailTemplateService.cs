using AIS.Data;
using AIS.Domain.Base;
using AIS.Domain.Email.Interfaces;
using System;
using System.Linq;

namespace AIS.Domain.Email
{
    public class EmailTemplateService : Service<ATC_EmailTemplate>, IEmailTemplateService
    {
        public EmailTemplateService(IUnitOfWork unitofwork)
            : base(unitofwork)
        {
        }

        public ATC_EmailTemplate GetEmailTemplateByType(string type)
        {
            var result = this.FindAll()
                        .Where(x => x.Type == type)
                        .FirstOrDefault();
            return result;
        }
    }
}
