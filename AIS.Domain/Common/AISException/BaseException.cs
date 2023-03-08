using System;

namespace AIS.Domain.Common.AISException
{
    public class BaseException : Exception
    {
        private readonly string _message;
        private readonly Exception _exception;

        public BaseException()
        {
        }

        public BaseException(string message)
        {
            _message = message;
        }

        public BaseException(string message, Exception exception)
        {
            _message = message;
            _exception = exception;
        }

        public string Mg
        {
            get { return _message; }
        }

        public Exception Exception
        {
            get { return _exception; }
        }
    }
}
