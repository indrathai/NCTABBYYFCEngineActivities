using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCTABBYYActivities.config
{
    class FConfig
    {
        // full path to FCE dll
        public const string DllPath = "C:\\Program Files (x86)\\ABBYY SDK\\12\\FlexiCapture Engine\\Bin\\FCEngine.dll";
        //public const string DllPath = "P:\\FCEngine.dll";

        //Return full path to FCE dll
        public static string GetDllPath()
        {
            return DllPath;
        }
    }
}
