using AIS.Domain.Common.AISException;
using AIS.Domain.Common.Constants;
using AIS.Domain.SMTP;
using System;
using System.Collections.Generic;
using System.Web;

namespace AIS.Domain.Email
{
    public class EmailMessageModel
    {
        public List<string> FileAttachs;
        public string FileAttach { get; set; }
        public string FileNameAttach { get; set; }

        public string Sender { get; set; }
        public string ToValue { get; set; }
        public string CcValue { get; set; }
        public string BccValue { get; set; }
        public string ReplyTo { get; set; }
        public string Subject { get; set; }
        public string Content { get; set; }
        public bool IsBodyHtml { get; set; }

        public string Password { get; set; }
        public string User { get; set; }
        public string Email { get; set; }
        public string Host { get; set; }

        public int CompanyId { get; set; }
        public string TypeMail { get; set; }

        public int Id1 { get; set; }
        public int Id2 { get; set; }
        public int Id3 { get; set; }
        public HttpSessionStateBase Session;


        public string EmailRoom { get; set; }
        public string UserLogin { get; set; }
        public string PasswordLogin { get; set; }

        public DateTime FromDateEvent { get; set; }
        public DateTime ToDateEvent { get; set; }
        public string Location { get; set; }
        public string MailDisplayName { get; set; }

        public bool Send()
        {
            if (!string.IsNullOrEmpty(User)) return SetValueToSend();
            User = StringConstants.UserEmailDefault;
            Password = StringConstants.PassEmailDefault;
            return SetValueToSend();
        }

        private bool SetValueToSend()
        {
            if (string.IsNullOrEmpty(ToValue))
            {
                return false;
            }

            try
            {
                CreateAndSendEmail();
                return true;
            }
            catch (ArgumentNullException ar)
            {
                throw new EmailException(ar.Message);
            }
        }

        private void CreateAndSendEmail()
        {
            var emailModel = new EmailSentModel
            {
                From = string.IsNullOrEmpty(Sender) ? User : Sender,
                Password = Password,
                EmailRecived = ToValue,
                Cc = CcValue,
                Bcc = BccValue,
                Subject = Subject,
                Body = Content,
                ReplyTo = ReplyTo,
                FileAttach = FileAttach,
                FileAttachList = FileAttachs,
                Notifier = Email,
                ServerMail = User
            };
            if (!string.IsNullOrEmpty(MailDisplayName)) emailModel.MailDisplayName = MailDisplayName;
            ISmtpService smtpService = new SmtpService();
            smtpService.SendEmail(emailModel);
        }
    }
}
