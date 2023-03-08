using AIS.Data;
using AIS.Domain.Common.Constants;
using AIS.Domain.Common.Helper;
using AIS.Domain.Email.Interfaces;
using AIS.Domain.HREmployee;
using AIS.Domain.SMTP;

namespace AIS.Domain.Email
{
    public class EmailService : IEmailService
    {
        private IHREmployeeService _hrEmployeeService;
        private IEmailTemplateService _emailTemplateService;

        public EmailService(IHREmployeeService hrEmployeeService , IEmailTemplateService emailTemplateService)
        {
            _hrEmployeeService = hrEmployeeService;
            _emailTemplateService = emailTemplateService;
        }

        public void SendMailForAbsenceRequest(EmailModel model)
        {
            var emailTemplate = _emailTemplateService.GetEmailTemplateByType(model.Template);
            var content = emailTemplate.Content;
            var manager = _hrEmployeeService.GetEmployeeInfoById(model.ManagerId);
            var requester = _hrEmployeeService.GetEmployeeInfoById(model.RequesterId);
            var emailHolder = new EmailReplaceHolderModel()
            {
                Manager = manager.Fullname,
                Requester = requester.Fullname,
                Note = model.Note,
                DateFrom = model.DateFrom.ToFormat(),
                DateTo = model.DateTo.ToFormat(),
            };
            var emailModel = new EmailSentModel()
            {
                Subject = emailTemplate.Subject,
                From = manager.EmailAddress_Ex,
                Cc = model.CC != null ? _hrEmployeeService.GetEmployeeInfoById((int)model.CC).EmailAddress_Ex : null,
                Bcc = model.BCC != null ? _hrEmployeeService.GetEmployeeInfoById((int)model.BCC).EmailAddress_Ex : null,
                EmailRecived = requester.EmailAddress_Ex
            };
            content = EmailTemplateHelper.ReplaceHolder(content, emailHolder);
            SendEmail(content, emailHolder, emailModel);
        }

        public void SendMailForAuthoriser(EmailModel model)
        {
            var emailTemplate = _emailTemplateService.GetEmailTemplateByType(model.Template);
            var content = emailTemplate.Content;
            var manager = _hrEmployeeService.GetEmployeeInfoById(model.ManagerId);
            var sender = _hrEmployeeService.GetEmployeeInfoById(model.RequesterId);
            var emailHolder = new EmailReplaceHolderModel()
            {
                Manager = manager.Fullname,
                Requester = sender.Fullname,
                Note = model.Note,
                LinkAprroveRequestForManager = StringConstants.AuthoriserRedirectInEmailURL
            };
            var emailModel = new EmailSentModel()
            {
                Subject = emailTemplate.Subject,
                From = sender.EmailAddress_Ex,
                EmailRecived = manager.EmailAddress_Ex
            };
            content = EmailTemplateHelper.ReplaceHolder(content, emailHolder);
            SendEmail(content, emailHolder, emailModel);
        }

        public bool SendEmail(string content, EmailReplaceHolderModel emailHolder, EmailSentModel emailModel)
        {
            var defaultUser = StringConstants.UserEmailDefault;
            var emailModel1 = new EmailSentModel
            {
                From = string.IsNullOrEmpty(emailModel.From) ? defaultUser : emailModel.From,
                Password = StringConstants.PassEmailDefault,
                EmailRecived = emailModel.EmailRecived,
                Cc = emailModel.Cc,
                Bcc = emailModel.Bcc,
                Subject = emailModel.Subject,
                Body = content,
                ServerMail = defaultUser
            };
            ISmtpService smtpService = new SmtpService();
            return smtpService.SendEmail(emailModel1);
        }

        public void SendMailForProjectArchiving(EmailModel model)
        {
            var defaultUser = StringConstants.UserEmailDefault;
            var emailTemplate = _emailTemplateService.GetEmailTemplateByType(model.Template);
            var content = emailTemplate.Content;
            var manager = _hrEmployeeService.GetEmployeeInfoById(model.ManagerId);
            
              var emailHolder = new EmailReplaceHolderModel()
            {
                Manager = manager.Fullname,
                ProjectID = model.APK7Character,
                EmailManager=manager.EmailAddress,
                TimeClose = model.TimeClose,
                Server= model.ServerPath,
           
            };
            var emailModel = new EmailSentModel()
            {
                Subject = emailTemplate.Subject,
                From = defaultUser,
               Cc = model.CC != null ? _hrEmployeeService.GetEmployeeInfoById((int)model.CC).EmailAddress_Ex : StringConstants.EmailITSupport,
                Bcc = model.BCC != null ? _hrEmployeeService.GetEmployeeInfoById((int)model.BCC).EmailAddress_Ex : StringConstants.EmailCSOArchiving,
                EmailRecived = manager.EmailAddress
            };
            emailModel.Subject = emailModel.Subject.Replace("#ProjectID#", model.APK7Character);
           
            content = EmailTemplateHelper.ReplaceHolder(content, emailHolder);
            SendEmail(content, emailHolder, emailModel);
        }
    }
}
