using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Threading.Tasks;

namespace RPADesktopActivityMonitor
{
    internal class RpaPopupManagement
    {
        static RpaPopupManagement _Instance = new RpaPopupManagement();
        public static RpaPopupManagement Instance { get => _Instance; }

        public delegate void PopupFeedbackReceivedEventHandler(Guid ID, int CloseMode);
        public delegate void RecordButtonFeedbackReceivedEventHandler();

        public event PopupFeedbackReceivedEventHandler PopupFeedbackReceived;
        public event RecordButtonFeedbackReceivedEventHandler RecordButtonFeedbackReceived;

        NamedPipeServerStream pipeFeedbackServer;
        NamedPipeClientStream pipeClient;
        StreamReader sr;
        StreamWriter sw;

        public void StartPopupServer()
        {
            var procs = Process.GetProcessesByName("PopupServer");
            if (procs.Length > 0)
            {
                return;
            }
            string PopupServerPath = Path.Combine(RPADesktopActivityMonitor.RuntimePath, "PopupServer.exe");
            ProcessStartInfo PopupProcInfo = new ProcessStartInfo()
            {
                FileName = PopupServerPath
            };
            Process PopupProcess = Process.Start(PopupProcInfo);
            PopupProcess.WaitForInputIdle();
        }

        public void InitPopupConnectivity()
        {
            pipeClient = new NamedPipeClientStream("popupMain");
            pipeFeedbackServer = new NamedPipeServerStream("popupFeedback", PipeDirection.InOut, 1);

            sr = new StreamReader(pipeFeedbackServer);
            sw = new StreamWriter(pipeClient);

            pipeClient.Connect();
            pipeFeedbackServer.WaitForConnection();
            Task.Run(ListenPopupFeedback);
        }

        private void ListenPopupFeedback()
        {
            string msg;
            while (pipeFeedbackServer.IsConnected)
            {
                try
                {
                    msg = sr.ReadLine();
                    if (msg == null) continue;
                    else if (msg == "StopRecord")
                    {
                        OnRecordButtonFeedbackReceived();
                    }
                    else
                    {
                        string[] param = msg.Split(new string[] { "$^^" }, StringSplitOptions.RemoveEmptyEntries);
                        if (param.Length == 2)
                        {
                            if (Guid.TryParse(param[0], out _) && Int32.TryParse(param[1], out _))
                            {
                                OnPopupFeedbackReceived(Guid.Parse(param[0]), Int32.Parse(param[1]));
                            }
                            else continue;
                        }
                        else continue;
                    }
                }
                catch
                {
                    if (pipeFeedbackServer.IsConnected) pipeFeedbackServer.Disconnect();
                    pipeFeedbackServer.Dispose();
                    return;
                }
            }
        }

        public void ShowPopup(int x, int y, string message, bool isNext)
        {
            string msg = $"{x}$^^{y}$^^{message}$^^{isNext}";
            sw.WriteLine(msg);
            sw.Flush();
        }

        public void RecordButtonCommand(bool isShow)
        {
            string msg;
            if (isShow) msg = "RecStarted";
            else msg = "RecStopped";
            sw.WriteLine(msg);
            sw.Flush();
        }
        //0 - internally
        //x - 1
        //next -2
        private void OnPopupFeedbackReceived(Guid ID, int CloseMode)
        {
            PopupFeedbackReceived?.Invoke(ID, CloseMode);
        }

        private void OnRecordButtonFeedbackReceived()
        {
            RecordButtonFeedbackReceived?.Invoke();
        }
    }
}
