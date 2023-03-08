using System.Collections.Generic;

namespace AIS.Domain.Email
{
    public class EmailSentModel
    {
        public EmailSentModel()
        {
        }
        public string From { get; set; }
        public string Password { get; set; }
        public string EmailRecived { get; set; }
        public string Bcc { get; set; }
        public string Cc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string ReplyTo { get; set; }
        public string FileAttach { get; set; }
        public List<string> FileAttachList { get; set; }
        public string Notifier { get; set; }
        public string ServerMail { get; set; }
        public string MailDisplayName { get; set; }
    }
}
