using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace RPADesktopActivityMonitor
{
    class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
        static void Main(string[] args)
        {
            SetProcessDPIAware();
            RunOptions opts = new RunOptions();
            if (args[1] == "--start-record")
            {
                opts.StartRecording = true;
            }
            else {
                opts.StopRecording = true;
            }
            RPADesktopActivityMonitor.GetInstance().Run(opts);            
        }        
    }
}
