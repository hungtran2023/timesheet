using System;
using System.Linq;
using System.Net.Mail;
using System.Configuration;
using System.Net;
using AIS.Domain.Email;
using System.Threading;

namespace AIS.Domain.SMTP
{
    public class SmtpService : ISmtpService
    {
        public static SmtpClient GetStmpClientService(string email, string password, string emailRecieve)
        {
            try
            {
                var fromAddress = new MailAddress(email);
                var stmpService = new SmtpClient()
                {
                    //This one for AIS livesite
                    //Host = ConfigurationManager.AppSettings["MailHost"],
                    //Port = 25,
                    //EnableSsl = false,
                    //DeliveryMethod = SmtpDeliveryMethod.Network,

                    Host = ConfigurationManager.AppSettings["MailHost"],
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, password)
                };

                return stmpService;
            }
            catch(Exception ex)
            {
                throw;
            }
        }

        public bool SendEmail(EmailSentModel emailModel)
        {
            var smtpClientToSendMail = GetStmpClientService(emailModel.ServerMail, emailModel.Password, emailModel.EmailRecived);
            if (smtpClientToSendMail == null) return false;
            var fromMail = GetRealFromEmail(emailModel);
            var email = new MailMessage {
                From = fromMail,
                Sender = fromMail
            };

            if (!string.IsNullOrEmpty(emailModel.EmailRecived))
            {
                email.To.Add(emailModel.EmailRecived.TrimEnd(';'));
            }

            if (!string.IsNullOrEmpty(emailModel.Cc))
            {
                email.CC.Add(emailModel.Cc);
            }

            if (!string.IsNullOrEmpty(emailModel.Bcc))
            {
                email.CC.Add(emailModel.Bcc);
            }

            if (!string.IsNullOrEmpty(emailModel.ReplyTo))
            {
                var replyToUser = new MailAddress(emailModel.ReplyTo);
                email.ReplyTo = replyToUser;
            }
            email.Subject = emailModel.Subject;
            email.Body = emailModel.Body;

            if (emailModel.FileAttach != null)
            {
                var attachment = new Attachment(emailModel.FileAttach);
                email.Attachments.Add(attachment);
            }
            else
            {
                if (emailModel.FileAttachList != null && emailModel.FileAttachList.Count > 0)
                {
                    foreach (var attach in emailModel.FileAttachList)
                    {
                        var attachment = new Attachment(attach);
                        email.Attachments.Add(attachment);
                    }
                }
            }
            email.IsBodyHtml = true;
            ThreadPool.QueueUserWorkItem(emails => smtpClientToSendMail.Send(email));
            return true;
        }

        private MailAddress GetRealFromEmail(EmailSentModel emailModel)
        {
            var displayName = string.IsNullOrEmpty(emailModel.MailDisplayName)
                ? "Atlas Information System"
                : emailModel.MailDisplayName;
            return new MailAddress(emailModel.From, displayName);
        }
    }
}
