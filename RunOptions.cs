using CommandLine;
using System;

namespace RPADesktopActivityMonitor
{
    interface IRunOptions
    {
        [Option("start-record", HelpText = "start desktop recording", Default = false)]
        bool StartRecording { get; set; }
        [Option("stop-record", HelpText = "stop desktop recording", Default = false)]
        bool StopRecording { get; set; }
        
    }

    [Verb("run", HelpText = "Run process automation.")]

    public class RunOptions : IRunOptions
    {       
        public bool StartRecording { get; set; }
        public bool StopRecording { get; set; }
    }
}
