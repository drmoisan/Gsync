using log4net.Repository.Hierarchy;
using log4net;
using Newtonsoft.Json.Serialization;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public class NLogTraceWriter : ITraceWriter
    {
        private static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(
                System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        internal virtual ILog Logger => _logger;

        private TraceLevel _levelFilter = TraceLevel.Verbose;
        public TraceLevel LevelFilter { get => _levelFilter; set => _levelFilter = value; }

        private List<string> _messageFilter = ["Deserialized JSON:", "Serialized JSON:"];
        public List<string> MessageFilter { get => _messageFilter; set => _messageFilter = value; }

        public void Trace(TraceLevel level, string message, Exception ex)
        {
            if (!MessageFilter.Select(message.Contains).Aggregate((a, b) => a | b))
            {
                var logFunction = GetLogFunction(level);
                logFunction?.Invoke(message, ex);
            }
        }

        private Action<string, Exception> GetLogFunction(TraceLevel level)
        {
            switch (level)
            {
                case TraceLevel.Error:
                    return Logger.Error;
                case TraceLevel.Warning:
                    return Logger.Warn;
                case TraceLevel.Info:
                    return Logger.Info;
                case TraceLevel.Off:
                    return null;
                default:
                    return Logger.Debug;
            }
        }

    }
}