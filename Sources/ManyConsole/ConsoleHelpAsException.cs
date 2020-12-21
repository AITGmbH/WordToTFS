using System;
namespace ManyConsole
{
    public class ConsoleHelpAsException : Exception
    {
        public ConsoleHelpAsException(string message) : base(message)
        {
        }
    }
}
