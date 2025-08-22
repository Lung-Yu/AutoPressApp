using System;

namespace AutoPressApp.Services
{
    public class LogService
    {
        public event Action<string>? OnLog;
        public void Info(string message)
        {
            OnLog?.Invoke(message);
            System.Diagnostics.Debug.WriteLine(message);
        }
    }
}
