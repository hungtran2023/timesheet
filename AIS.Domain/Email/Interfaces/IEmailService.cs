namespace AIS.Domain.Email.Interfaces
{
    public interface IEmailService
    {
        bool SendEmail(string content, EmailReplaceHolderModel emailHolder, EmailSentModel emailModel);
        void SendMailForAbsenceRequest(EmailModel model);
        void SendMailForAuthoriser(EmailModel model);

        void SendMailForProjectArchiving(EmailModel model);
    }
}
