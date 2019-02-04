using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCTABBYYActivities.model
{
    public class LogMessage
    {
        public LogMessage()
        {

        }

        public LogMessage(string message, LogType logType)
        {
            this.TypeOfLog = logType;
            this.Message = message;
        }


        public LogType TypeOfLog { get; set; }
        public string Message { get; set; }
    }
}
