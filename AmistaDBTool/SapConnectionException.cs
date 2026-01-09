using System;

namespace AmistaDBTool
{
    public class SapConnectionException : Exception
    {
        public int ErrorCode { get; }

        public SapConnectionException() { }
        public SapConnectionException(string message) : base(message) { }
        public SapConnectionException(string message, Exception innerException) : base(message, innerException) { }
        public SapConnectionException(string message, int errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }
    }
}
