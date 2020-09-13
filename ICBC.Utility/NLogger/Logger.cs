using System.IO;
using Microsoft.Extensions.Logging;
using NLog;
namespace ICBC.Utility
{
    public class Logger
    {
        public static Microsoft.Extensions.Logging.ILogger GetLogger() => new Microsoft.Extensions.Logging.LoggerFactory() { 
                builder =>
                {
                   
                    .AddFilter("Microsoft", Microsoft.Extensions.Logging.LogLevel.Warning)
                    .AddFilter("System", Microsoft.Extensions.Logging.LogLevel.Warning)
                    .AddNLog("nlog.config")
                }
                }
                //}).CreateLogger("Robot Logging Service");public static Microsoft.Extensions.Logging.ILogger GetLogger() => Microsoft.Extensions.Logging.LoggerFactory.Create(
                //builder =>
                //{
                //    builder
                //    .AddFilter("Microsoft", Microsoft.Extensions.Logging.LogLevel.Warning)
                //    .AddFilter("System", Microsoft.Extensions.Logging.LogLevel.Warning)
                //    .AddNLog("nlog.config");
                //}).CreateLogger("Robot Logging Service");

        public static void LoadConfig()
        {
            LogManager.LoadConfiguration(string.Concat(Directory.GetCurrentDirectory(), "\\nlog.config"));
        }

    }
}
