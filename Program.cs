using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPADesktopActivityMonitor
{
    class Program
    {
        static void Main(string[] args)
        {            
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
