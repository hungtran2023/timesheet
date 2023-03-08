using System;

namespace AIS.Domain.Common.AISException
{
    public class EmailException : BaseException
    {
        public new string Message { get; set; }

        public EmailException(Exception exception) : base("Email can not send.", exception)
        {
        }

        public EmailException(string message) : base(message)
        {
            Message = message;
        }

        public EmailException()
        {
            Message = "Email can not send!";
        }
    }
}
