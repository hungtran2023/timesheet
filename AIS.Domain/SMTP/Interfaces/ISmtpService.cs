using AIS.Domain.Email;

namespace AIS.Domain.SMTP
{
    public interface ISmtpService
    {
        bool SendEmail(EmailSentModel emailModel);
    }
}
