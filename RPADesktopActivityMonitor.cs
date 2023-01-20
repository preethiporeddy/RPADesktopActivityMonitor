using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Threading;
using System.Windows.Forms;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net;
using System.Linq;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Dynamic;

namespace RPADesktopActivityMonitor
{
    public class RPADesktopActivityMonitor : IDisposable
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYUP = 0x0105;
        private uint[] keyMap = new uint[10];
        private int currentCount = 0, n = 0, m = 0, l = 0;
        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int x;
            public int y;
        }
        [StructLayout(LayoutKind.Sequential)]
        internal struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential)]
        internal struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
        private MSLLHOOKSTRUCT p, q;
        private KBDLLHOOKSTRUCT k;
        private IntPtr hook = IntPtr.Zero;
        private IntPtr mhook = IntPtr.Zero;
        private LowLevelKeyboardProc keyproc;
        private LowLevelMouseProc mouseproc;
        private System.Windows.Forms.Timer timecnt;
        private FileStream fp;
        private StreamWriter fw;
        private StreamReader fr;
        private bool fn = false, noprint = false, ishooked = false;
        private string typed = "", img = "", scr = "";
        private Dispatcher Dispatcher = null;
        public static readonly string TFRPA_RUNTIME_PATH = Environment.GetEnvironmentVariable("RUNTIME_PATH");
        public static string RuntimePath = !String.IsNullOrEmpty(TFRPA_RUNTIME_PATH) ? TFRPA_RUNTIME_PATH : AppDomain.CurrentDomain.BaseDirectory;
        /*public static string RuntimePath = "C:\\Users\\Techforce\\AppData\\Local\\Techforce.ai\\Super Assistant App\\resources\\static\\Techforce\\src";*/

        public NamedPipeClientStream clientStream = new NamedPipeClientStream(".", "superextension", PipeDirection.In);
        private static dynamic configObject = JsonConvert.DeserializeObject(File.ReadAllText(Path.Combine(RuntimePath, "..\\..\\Configs\\extnHost_chrome\\config.json")));
        public static string refreshtoken = configObject.token;
        public static bool IsAccessTokenStored = false;
        public static string StoredAccessToken = "";
        public static Task[] ToBeCompletedTasks;
        private static readonly RPADesktopActivityMonitor Instance = new RPADesktopActivityMonitor();   
       // private static RpaPopupManagement popupInstance = new RpaPopupManagement();

        private string appName = "";
        private string tempPath = Path.Combine(Path.GetTempPath(), $"dump.txt");        
        private int imgid = 0, scrid = 0;
        private Dictionary<int, string> scrDict = new Dictionary<int, string>();
        private Dictionary<int, string> imgDict = new Dictionary<int, string>();

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi)]
        static extern IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPStr)] string lpFileName);

        public RPADesktopActivityMonitor()
        {
            Dispatcher dispatcher = null;
            var manualResetEvent = new ManualResetEvent(false);
            var thread = new Thread(() =>
            {
                dispatcher = Dispatcher.CurrentDispatcher;
                var synchronizationContext = new DispatcherSynchronizationContext(dispatcher);
                SynchronizationContext.SetSynchronizationContext(synchronizationContext);

                manualResetEvent.Set();
                Dispatcher.Run();
            });
            thread.Start();
            manualResetEvent.WaitOne();
            manualResetEvent.Dispose();
            Dispatcher = dispatcher;

            //RpaPopupManagement.Instance.RecordButtonFeedbackReceived += DesktopInstance_RecordButtonFeedbackReceived;
        }
        private void DesktopInstance_RecordButtonFeedbackReceived()
        {
            RpaPopupManagement.Instance.RecordButtonCommand(false);
            byte[] dataArr = new byte[clientStream.InBufferSize];
            clientStream.Read(dataArr, 0, clientStream.InBufferSize);
            var Execoutput = Encoding.UTF8.GetString(dataArr);
            if (Execoutput.Length > 0)
            {
                dynamic executeoutputobject = JsonConvert.DeserializeObject(Execoutput);
                if (!(Execoutput.ToLower().Contains("userActivity")))
                {
                    if (executeoutputobject != null)
                    {
                        if (PropertyExists(executeoutputobject, "action"))
                        {
                            string actionVar = executeoutputobject.action.Value;
                            if (actionVar == "stopRecording")
                            {
                                Task.WaitAll(ToBeCompletedTasks);
                                stopLogging();
                                File.AppendAllText(tempPath, $"Finish will be executed{Environment.NewLine}");
                                finish(RuntimePath, executeoutputobject);
                                Console.WriteLine("Activity Monitor (Desktop) Stopped");
                                Dispose(); //CALL THIS IF YOU WANT THE PROGRAM TO EXIT
                                File.AppendAllText(tempPath, $"Finished Completely{Environment.NewLine}");
                                clientStream.Dispose();
                            }
                        }
                        else
                        {
                            clientStream.Read(dataArr, 0, clientStream.InBufferSize);
                            Execoutput = Encoding.UTF8.GetString(dataArr);
                            if (Execoutput.Length > 0)
                            {
                                executeoutputobject = JsonConvert.DeserializeObject(Execoutput);
                                if (executeoutputobject != null)
                                {
                                    if (PropertyExists(executeoutputobject, "action"))
                                    {
                                        string actionVar = executeoutputobject.action.Value;
                                        if (actionVar == "stopRecording")
                                        {
                                            Task.WaitAll(ToBeCompletedTasks);
                                            stopLogging();
                                            File.AppendAllText(tempPath, $"Finish will be executed{Environment.NewLine}");
                                            finish(RuntimePath, executeoutputobject);
                                            Console.WriteLine("Activity Monitor (Desktop) Stopped");
                                            Dispose(); //CALL THIS IF YOU WANT THE PROGRAM TO EXIT
                                            File.AppendAllText(tempPath, $"Finished Completely{Environment.NewLine}");
                                            clientStream.Dispose();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
        internal static RPADesktopActivityMonitor GetInstance() => Instance;

        private string Filter(string inp, bool cased = false)
        {
            string output = String.Empty;
            switch (inp)
            {
                case "D1":
                    if (cased) output = "!";
                    else output = "1";
                    break;
                case "NumPad1":
                    output = "1";
                    break;
                case "D2":
                    if (cased) output = "@";
                    else output = "2";
                    break;
                case "NumPad2":
                    output = "2";
                    break;
                case "D3":
                    if (cased) output = "#";
                    else output = "3";
                    break;
                case "NumPad3":
                    output = "3";
                    break;
                case "D4":
                    if (cased) output = "$";
                    else output = "4";
                    break;
                case "NumPad4":
                    output = "4";
                    break;
                case "D5":
                    if (cased) output = "%";
                    else output = "5";
                    break;
                case "NumPad5":
                    output = "5";
                    break;
                case "D6":
                    if (cased) output = "^";
                    else output = "6";
                    break;
                case "NumPad6":
                    output = "6";
                    break;
                case "D7":
                    if (cased) output = "&";
                    else output = "7";
                    break;
                case "NumPad7":
                    output = "7";
                    break;
                case "D8":
                    if (cased) output = "*";
                    else output = "8";
                    break;
                case "NumPad8":
                    output = "8";
                    break;
                case "D9":
                    if (cased) output = "(";
                    else output = "9";
                    break;
                case "NumPad9":
                    output = "9";
                    break;
                case "D0":
                    if (cased) output = ")";
                    else output = "0";
                    break;
                case "NumPad0":
                    output = "0";
                    break;
                case "LShiftKey":
                case "RShiftKey":
                    output = "shift";
                    break;
                case "LControlKey":
                case "RControlKey":
                    output = "ctrl";
                    break;
                case "LMenu":
                case "RMenu":
                    output = "alt";
                    break;
                case "LWin":
                case "RWin":
                    output = "win";
                    break;
                case "OemMinus":
                    if (cased) output = "_";
                    else output = "-";
                    break;
                case "Oemplus":
                    if (cased) output = "+";
                    else output = "=";
                    break;
                case "Oemtilde":
                    if (cased) output = "~";
                    else output = "`";
                    break;
                case "OemOpenBrackets":
                    if (cased) output = "{";
                    else output = "[";
                    break;
                case "Oem6":
                    if (cased) output = "}";
                    else output = "]";
                    break;
                case "Oem5":
                    if (cased) output = "|";
                    else output = "\\";
                    break;
                case "Oem1":
                    if (cased) output = ":";
                    else output = ";";
                    break;
                case "Oem7":
                    if (cased) output = "\"";
                    else output = "'";
                    break;
                case "Oemcomma":
                    if (cased) output = "<";
                    else output = ",";
                    break;
                case "OemPeriod":
                    if (cased) output = ">";
                    else output = ".";
                    break;
                case "OemQuestion":
                    if (cased) output = "?";
                    else output = "/";
                    break;
                default:
                    if (cased || Control.IsKeyLocked(Keys.Capital)) output = inp.ToUpper();
                    else output = inp.ToLower();
                    break;
            }
            return output;
        }

        private int Modifier(Keys key)
        {
            if (key == Keys.LShiftKey || key == Keys.RShiftKey) return 1;
            if (key == Keys.LControlKey || key == Keys.RControlKey) return 2;
            if (key == Keys.LWin || key == Keys.RWin) return 3;
            if (key == Keys.LMenu || key == Keys.RMenu) return 4;
            else return 5;
        }

        public void Init()
        {
            AppDomain.CurrentDomain.UnhandledException += crashManagement;
            AppDomain.CurrentDomain.ProcessExit += onExit;
            timecnt = new System.Windows.Forms.Timer();
            timecnt.Interval = 501;
            timecnt.Tick += new EventHandler(timecnt_Tick);
            keyproc = new LowLevelKeyboardProc(hookCallback);
            mouseproc = new LowLevelMouseProc(mhookCallback);
        }

        private void onExit(object sender, EventArgs e)
        {
            try
            {
                cancel();
            }
            catch (Exception f)
            {
                Environment.Exit(1);
            }
        }

        private void crashManagement(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                cancel();
            }
            catch (Exception f)
            {
                Environment.Exit(1);
            }
        }

        public void startLogging()
        {
            imgid = 0;
            scrid = 0;
            imgDict.Clear();
            scrDict.Clear();
            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"))) Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"), true);
            fp = File.Create(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"));
            fw = new StreamWriter(fp);
            fr = new StreamReader(fp);
            fw.AutoFlush = true;
            fw.WriteLine("[");
            if (File.Exists(tempPath)){
                File.Delete(tempPath);
            }
            File.AppendAllText(tempPath, $"Logging started{Environment.NewLine}");
            Dispatcher.Invoke(() =>
            {
                hook = SetHook(keyproc);
                mhook = SetmHook(mouseproc);
            });
            ishooked = true;
        }

        public void stopLogging()
        {
            Dispatcher.Invoke(() =>
            {
                UnhookWindowsHookEx(hook);
                UnhookWindowsHookEx(mhook);
            });
            ishooked = false;
            if (!String.IsNullOrEmpty(typed))
            {
                fw.WriteLine("    {");
                fw.WriteLine("        \"type\": \"record\",");
                fw.WriteLine("        \"actionType\": \"" + "kbd" + "\",");
                fw.WriteLine("        \"value\": \"" + typed + "\",");
                fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                fw.WriteLine("    },");
                typed = "";
            }
            //File.AppendAllText(tempPath, $"in stop logging {File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"))}{Environment.NewLine}");
        }
        public bool PropertyExists(dynamic obj, string name)
        {
            if (obj == null) return false;
            if (obj is ExpandoObject)
                return ((IDictionary<string, object>)obj).ContainsKey(name);
            if (obj is IDictionary<string, object> dict1)
                return dict1.ContainsKey(name);
            if (obj is IDictionary<string, JToken> dict2)
                return dict2.ContainsKey(name);
            return obj.GetType().GetProperty(name) != null;
        }
        public void Run(RunOptions opts)
        {
            if (opts.StartRecording)
            {
                /*string extninFile = Path.Combine(RPA.RuntimePath, "../../Configs/extnHost_chrome/extn.in");
                string ExecutionInstruction = JsonConvert.SerializeObject(outputobject);
                File.WriteAllText(extninFile, ExecutionInstruction);*/
                var procs = Process.GetProcessesByName("RPADesktopActivityMonitor");
                if (procs.Length > 0)
                {
                    int currentProcessID = Process.GetCurrentProcess().Id;
                    for (int i = 0; i < procs.Length; i++)
                    {
                        if (procs[i].Id == currentProcessID)
                        { }
                        else
                        {
                            procs[i].Kill();
                        }
                    }
                }
                Init();
                startLogging();
                Console.WriteLine("Activity Monitor (Desktop) Started");

                Task.Run(() =>
                {
                    clientStream.Connect();
                    byte[] dataArr = new byte[clientStream.InBufferSize];
                    clientStream.Read(dataArr, 0, clientStream.InBufferSize);
                    var Execoutput = Encoding.UTF8.GetString(dataArr);
                    if (Execoutput.Length > 0)
                    {
                        dynamic executeoutputobject = JsonConvert.DeserializeObject(Execoutput);
                        if (!(Execoutput.ToLower().Contains("userActivity")))
                        {
                            if (executeoutputobject != null)
                            {
                                if (PropertyExists(executeoutputobject, "action"))
                                {
                                    string actionVar = executeoutputobject.action.Value;
                                    if (actionVar == "stopRecording")
                                    {
                                        stopLogging();
                                        File.AppendAllText(tempPath, $"Finish will be executed{Environment.NewLine}");
                                        finish(RuntimePath, executeoutputobject);
                                        Console.WriteLine("Activity Monitor (Desktop) Stopped");
                                        Dispose(); //CALL THIS IF YOU WANT THE PROGRAM TO EXIT
                                        File.AppendAllText(tempPath, $"Finished Completely{Environment.NewLine}");
                                        clientStream.Dispose();
                                    }
                                }
                                else
                                {
                                    clientStream.Read(dataArr, 0, clientStream.InBufferSize);
                                    Execoutput = Encoding.UTF8.GetString(dataArr);
                                    if (Execoutput.Length > 0)
                                    {
                                        executeoutputobject = JsonConvert.DeserializeObject(Execoutput);
                                        if (executeoutputobject != null)
                                        {
                                            if (PropertyExists(executeoutputobject, "action"))
                                            {
                                                string actionVar = executeoutputobject.action.Value;
                                                if (actionVar == "stopRecording")
                                                {
                                                    stopLogging();
                                                    File.AppendAllText(tempPath, $"Finish will be executed{Environment.NewLine}");
                                                    finish(RuntimePath, executeoutputobject);
                                                    Console.WriteLine("Activity Monitor (Desktop) Stopped");
                                                    Dispose(); //CALL THIS IF YOU WANT THE PROGRAM TO EXIT
                                                    File.AppendAllText(tempPath, $"Finished Completely{Environment.NewLine}");
                                                    clientStream.Dispose();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }


                });
            }
        }

        private void UploadRecordedEventsToSuper(dynamic payload)
        {
            HttpClient httpClient = new HttpClient();
            /*tring boundary = "--------------------------" + DateTime.Now.Ticks.ToString() + DateTime.Now.Ticks.ToString().Substring(0, 6);
            MultipartFormDataContent form = new MultipartFormDataContent(boundary);
            var file_bytes = File.ReadAllBytes(filename);
            form.Add(new ByteArrayContent(file_bytes, 0, file_bytes.Length), "file", filename);*/
            var postUrl = new Uri("https://superapiqa.development.techforce.ai/botapi/draftSkills/extension/record");
            /*dynamic configObject = JsonConvert.DeserializeObject(File.ReadAllText(Path.Combine(RuntimePath, "../../Configs/extnHost_chrome/config.json")));
            string refreshtoken = configObject.token;*/
            string accesstoken = GetStoredAccessToken(refreshtoken);//"eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICJ0TXVIRlRmNjZYRlFGWEtPX3JlVF96ajlVTkRyLW1wX0JYQnpZSHZsYVd3In0.eyJleHAiOjE2NTc2MDE5OTAsImlhdCI6MTY0OTgyNTk5MCwiYXV0aF90aW1lIjoxNjQ5NzU1Njk2LCJqdGkiOiI4MjFiZmZkYS1jYjJmLTRkZGYtYTY0My1lOWU0ZjA3ZjgxYTUiLCJpc3MiOiJodHRwczovL2VtY2F1dGhhd3MudGVjaGZvcmNlLmFpL2F1dGgvcmVhbG1zL3RlY2hmb3JjZSIsImF1ZCI6ImFjY291bnQiLCJzdWIiOiI1NzViNDJjYi01YmViLTQ1NDgtOTU0OC02NGU5NWJmMDk5NTYiLCJ0eXAiOiJCZWFyZXIiLCJhenAiOiJzdXBlci1leHRlbnNpb24iLCJub25jZSI6ImIzNmViODRlLTk5Y2UtNDUwYS04YmZiLTEzN2NjMGUxYWMzNCIsInNlc3Npb25fc3RhdGUiOiIyN2E5M2EzZi1jY2NmLTRjYzYtOTAxNy02N2YwOGU2ZmE2NTEiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImNocm9tZS1leHRlbnNpb246Ly9sbmdocGlpZXBmYmdnY2trb2dtam1wYmhubmFkbG1lcCIsImNocm9tZS1leHRlbnNpb246Ly9ka2FrZmpsbG9lZGNtam5na2lvbWdtZWdmY2ZnbWdvZiIsImNocm9tZS1leHRlbnNpb246Ly9scGNkaWlpbmJnbm1wbmliYWFram9kYWxnaHBwaGtrYy8qIiwiY2hyb21lLWV4dGVuc2lvbjovL2duYWlnZmFoZmxkbGpoZGhnbGFqb2JoaGhvcG9tbG9nIiwiY2hyb21lLWV4dGVuc2lvbjovL2lrYmNnZm1qa21maWptZmtvbGtpYWptb2Znam9qbG5oIiwiY2hyb21lLWV4dGVuc2lvbjovL2RkaGhsb25pZGFma2xoY2dvb2drZWVjbmZsZ2dqZGVlIiwiY2hyb21lLWV4dGVuc2lvbjovL25qamhiZ2lwaHBvZGNnY2RraWtrb2RwcG5oYmRpYXBtIiwiY2hyb21lLWV4dGVuc2lvbjovL25oY2JlZWxhbWtncG5oYWxiYWNtaWthYmNtZGRpZWNoIiwiY2hyb21lLWV4dGVuc2lvbjovL2xicGhsYWpncGpkbGRlaWNvaG5lZ2FrYm1mZmZoY2ZlIiwiY2hyb21lLWV4dGVuc2lvbjovL21jZmdna2RmZGRuZGtsY2NoYWhsZ2pjZ2Vrb2xnYW9jIiwiY2hyb21lLWV4dGVuc2lvbjovL21vYWhwZW5na2tmb25qYW5nZGlvYWtiZ2xtbmVnbWdtIiwiY2hyb21lLWV4dGVuc2lvbjovL2hvcG9ma2Nrb2pnaG9qZG5wcGNnamRiZGtoaWtsbmNqIiwiY2hyb21lLWV4dGVuc2lvbjovL2ljZWVmZWFrb2tnZGlwam1oZWVlZmhqZWhubGRpb29qIiwiY2hyb21lLWV4dGVuc2lvbjovL3BubmNibG1wZGlrbmxvbXBra2Nha2pvYWhwcGtmZGVqIiwiY2hyb21lLWV4dGVuc2lvbjovL2NpZGRwcGhwaWtsZ2pnZmRkbGduaGZpYmJjb2xubGxrIiwiY2hyb21lLWV4dGVuc2lvbjovL2JmbmVqaG1vb2Nsa2JmYW5sYWluZGpna2RhbGxpa21qIiwiY2hyb21lLWV4dGVuc2lvbjovL21pamFibmlmbmRjb2VnbWttYWhmYWNjZWJjaWdtcGFuIiwiY2hyb21lLWV4dGVuc2lvbjovL2pqaGxkbWpoYm5mampwYmxqam9kZWlqbHBsb2ZqYm9lIiwiY2hyb21lLWV4dGVuc2lvbjovL2dwZG5tZm9hZG5saGlqZGljZWNpbWtsb21wamljaWdtIl0sInJlYWxtX2FjY2VzcyI6eyJyb2xlcyI6WyJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYWNjb3VudCI6eyJyb2xlcyI6WyJtYW5hZ2UtYWNjb3VudCIsIm1hbmFnZS1hY2NvdW50LWxpbmtzIiwidmlldy1wcm9maWxlIl19fSwic2NvcGUiOiJvcGVuaWQgcHJvZmlsZSBlbWFpbCIsImVtYWlsX3ZlcmlmaWVkIjp0cnVlLCJuYW1lIjoiUHJlZXRoaSBQb3JlZGR5IiwicHJlZmVycmVkX3VzZXJuYW1lIjoic3BvcmVkZHlAdGVjaGZvcmNlLmFpIiwiZ2l2ZW5fbmFtZSI6IlByZWV0aGkiLCJmYW1pbHlfbmFtZSI6IlBvcmVkZHkiLCJlbWFpbCI6InNwb3JlZGR5QHRlY2hmb3JjZS5haSJ9.DrbZZyGvE_tV4C_GZHc-QxfbSYS87bD27-X37v8rg4le-cOsaenYyk6xvKOXlu87o5mTSd1pjDwJ21sxtBxdSMJg7PyoT3FgYOfmqccQK-JUwKW7EGXayYsMj5HptdNl5E6DLB95Xh7kFstUO65QZ5IDyMp862fAFLvkK7ATe8D4r3QYh3b6jvkmMXSA-ZmfIwMbG90OFVLeoSDdPKLFe95yvVZXFvawI_CXB0vQU13EW4XJ6wfBuZh0RMVW0DZ0yUlclQ1nt9_G5blMowtJ4Dgv5DLL9_Yhz5BmqS0yQuN_YbnY3iLcGXWtW9NQSdFbVQIQ-JGZY7g6ddN_13QpAg";
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", string.Format("Bearer {0}", accesstoken));
            /* var boundaryParameter = form.Headers.ContentType
                                         .Parameters.Single(p => p.Name == "boundary");
             boundaryParameter.Value = boundaryParameter.Value.Replace("\"", "");*/
            HttpContent content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HttpResponseMessage response = httpClient.PostAsync(postUrl, content).Result;
            response.EnsureSuccessStatusCode();
            httpClient.Dispose();
            File.AppendAllText(tempPath, $"{response.EnsureSuccessStatusCode()}--{Environment.NewLine}");
            File.AppendAllText(tempPath, $"uploaded recorded events to super{Environment.NewLine}");
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            IntPtr processHandle = LoadLibrary("user32.dll");
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc, processHandle, 0);
        }

        private IntPtr SetmHook(LowLevelMouseProc mproc)
        {
            IntPtr processHandle = LoadLibrary("user32.dll");
            return SetWindowsHookEx(WH_MOUSE_LL, mproc, processHandle, 0);
        }

        private IntPtr hookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if ((nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN) || (nCode >= 0 && wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                k = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                uint currentKey = k.vkCode;
                int i;
                for (i = 0; i < currentCount; i++)
                {
                    if (keyMap[i] == currentKey)
                    {
                        n++;
                        break;
                    }
                }
                if (i == currentCount)
                    keyMap[currentCount++] = currentKey;
                noprint = false;
            }
            if ((nCode >= 0 && wParam == (IntPtr)WM_KEYUP) || (nCode >= 0 && wParam == (IntPtr)WM_SYSKEYUP))
            {
                k = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                if (!noprint && !isWindowExcluded()) printKeys();
                uint currentKey = k.vkCode;
                int i;
                for (i = 0; i < currentCount; ++i)
                {
                    if (keyMap[i] == currentKey)
                        break;
                }
                for (; i < currentCount - 1; ++i)
                {
                    keyMap[i] = keyMap[i + 1];
                }
                keyMap[currentCount--] = 0;
                if (currentCount == 0)
                    n = 0;
            }
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        private IntPtr mhookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (!String.IsNullOrEmpty(typed))
                {
                    fw.WriteLine("    {");
                    fw.WriteLine("        \"type\": \"record\",");
                    fw.WriteLine("        \"origin\": \"desktop\",");
                    fw.WriteLine("        \"description\": \"Keystroke\",");
                    fw.WriteLine("        \"actionType\":\"" + "kbd" + "\",");
                    fw.WriteLine("        \"value\": \"" + typed + "\",");
                    fw.WriteLine($"        \"app_name\": \"{appName}\",");
                    fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                    fw.WriteLine("    },");
                    typed = "";
                }
                if (isWindowExcluded()) goto ret;
                if (wParam == (IntPtr)WM_LBUTTONDOWN)
                {

                }
                else if (wParam == (IntPtr)WM_RBUTTONDOWN)
                {

                }
                else if (wParam == (IntPtr)WM_LBUTTONUP)
                {
                    p = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    if (fn == true)
                    {
                        timecnt.Stop();
                        fw.WriteLine("    {");
                        fw.WriteLine("        \"type\": \"record\",");
                        fw.WriteLine("        \"origin\": \"desktop\",");
                        fw.WriteLine("        \"description\": \"Double Click\",");
                        fw.WriteLine("        \"selectorType\": \"imgSelector\",");
                        fw.WriteLine("        \"actionType\":\"dbclick\",");
                        fw.WriteLine($"        \"image\": \"{img}\",");
                        fw.WriteLine($"        \"pageScreenshot\": \"{scr}\",");
                        fw.WriteLine($"        \"app_name\": \"{appName}\",");
                        fw.WriteLine("        \"coord\": {");
                        fw.WriteLine("            \"x\": " + p.pt.x + ",");
                        fw.WriteLine("            \"y\": " + p.pt.y);
                        fw.WriteLine("        },");
                        fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                        fw.WriteLine("    },");
                        fn = false;
                        return IntPtr.Zero;
                    }
                    if (!timecnt.Enabled)
                    {
                        var a = imgid;
                        var b = scrid;
                        imgDict.Add(a, "");
                        scrDict.Add(b, "");
                        img = a.ToString();
                        scr = b.ToString();
                        Task.Run(() => { imgDict[a] = takeSnap(p.pt.x, p.pt.y); });
                        Task.Run(() => { scrDict[b] = takeScreen(); });
                        imgid++;
                        scrid++;
                        timecnt.Start();
                        fn = true;
                    }
                }
                else if (wParam == (IntPtr)WM_RBUTTONUP)
                {
                    q = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    var a = imgid;
                    var b = scrid;
                    imgDict.Add(a, "");
                    scrDict.Add(b, "");
                    img = a.ToString();
                    scr = b.ToString();
                    Task.Run(() => { imgDict[a] = takeSnap(q.pt.x, q.pt.y); });
                    Task.Run(() => { scrDict[b] = takeScreen(); });
                    imgid++;
                    scrid++;
                    fw.WriteLine("    {");
                    fw.WriteLine("        \"type\": \"record\",");
                    fw.WriteLine("        \"origin\": \"desktop\",");
                    fw.WriteLine("        \"description\": \"Right Click\",");
                    fw.WriteLine("        \"selectorType\": \"imgSelector\",");
                    fw.WriteLine("        \"actionType\":\"rclick\",");
                    fw.WriteLine($"        \"image\": \"{img}\",");
                    fw.WriteLine($"        \"pageScreenshot\": \"{scr}\",");
                    fw.WriteLine($"        \"app_name\": \"{appName}\",");
                    fw.WriteLine("        \"coord\": {");
                    fw.WriteLine("            \"x\": " + q.pt.x + ",");
                    fw.WriteLine("            \"y\": " + q.pt.y);
                    fw.WriteLine("        },");
                    fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                    fw.WriteLine("    },");
                }
            }
        ret:
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        private void keyPress(bool cased = false)
        {
            int i = 0;
            string toPrint = "";
            if (cased) i++; //Experimental
            for (; i < currentCount; i++)
            {
                if ((keyMap[i] >= 48 && keyMap[i] <= 57) || (keyMap[i] >= 65 && keyMap[i] <= 90) || (Control.IsKeyLocked(Keys.NumLock) && keyMap[i] >= 96 && keyMap[i] <= 105))
                {
                    toPrint += "kbd";
                }
                else
                    toPrint += "kp";
                toPrint += "," + ((Keys)keyMap[i]).ToString() + ",";
            }
            parse_Json(toPrint, cased);
        }

        private void keyHold()
        {
            string toPrint = "";
            toPrint += "kh";
            for (int i = 0; i < currentCount; i++)
            {
                toPrint += "," + ((Keys)keyMap[i]).ToString();
            }
            parse_Json(toPrint);
        }

        private void keyComb()
        {
            string toPrint = "";
            toPrint += "kc";
            for (int i = 0; i < currentCount; i++)
            {
                toPrint += "," + ((Keys)keyMap[i]).ToString();
            }
            var b = scrid;
            scrDict.Add(b, "");
            scr = b.ToString();
            Task.Run(() => { scrDict[b] = takeScreen(); });
            scrid++;
            parse_Json(toPrint);
        }

        private void printKeys()
        {
            noprint = true;
            int cased = 0;
            if (currentCount > 1)
            {
                for (int i = 0; i < currentCount; i++)
                {
                    if (Modifier((Keys)keyMap[i]) == 1)
                    {
                        cased++; //Experimental
                        continue;
                    }
                    if (Modifier((Keys)keyMap[i]) == 2 || Modifier((Keys)keyMap[i]) == 3 || Modifier((Keys)keyMap[i]) == 4)
                    {
                        cased += Modifier((Keys)keyMap[i]); //Experimental
                        continue;
                    }
                    if ((Keys)keyMap[i] == Keys.Capital || (Keys)keyMap[i] == Keys.Scroll || (Keys)keyMap[i] == Keys.NumLock)
                    {
                        keyPress();
                        break;
                    }
                    else
                    {
                        if (!(keyMap[i] >= 48 && keyMap[i] <= 57) && !(keyMap[i] >= 65 && keyMap[i] <= 90) && !(keyMap[i] >= 186 && keyMap[i] <= 192) && !(keyMap[i] >= 219 && keyMap[i] <= 222) && cased == 1)
                        {
                            cased += 5;
                        }
                        if (cased == 1) keyPress(true); //Experimental
                        else if (cased > 1) keyComb();
                        else keyPress();
                        break;
                    }
                }

            }
            else if (n > 1)
            {
                keyHold();
            }
            else
            {
                keyPress();
            }

        }

        private void parse_Json(string input, bool cased = false)
        {
            if (!input.StartsWith("kbd") && !String.IsNullOrEmpty(typed))
            {
                fw.WriteLine("    {");
                fw.WriteLine("        \"type\": \"record\",");
                fw.WriteLine("        \"origin\": \"desktop\",");
                fw.WriteLine("        \"description\": \"Keystroke\",");
                fw.WriteLine("        \"actionType\":\"" + "kbd" + "\",");
                fw.WriteLine("        \"value\": \"" + typed + "\",");
                fw.WriteLine($"        \"app_name\": \"{appName}\",");
                fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                fw.WriteLine("    },");
                typed = "";
            }
            String[] strs = input.Split(',');
        kpkh:
            if (strs[0] == "kp" || strs[0] == "kh")
            {
                for (int i = 1; i < strs.Length; i++)
                {
                    if (strs[i] == "kbd")
                    {
                        input = "";
                        for (; i < strs.Length; i++)
                        {
                            input += strs[i] + ",";
                        }
                        strs = input.Split(',');
                        goto kbd;
                    }
                    if (strs[i] == "kp" || strs[i] == "kh") continue;
                    if (strs[i] == "") continue;
                    fw.WriteLine("    {");
                    fw.WriteLine("        \"type\": \"record\",");
                    fw.WriteLine("        \"origin\": \"desktop\",");
                    fw.WriteLine("        \"description\": \"Keypress\",");
                    fw.WriteLine("        \"actionType\":\"" + strs[0] + "\",");
                    fw.WriteLine("        \"value\": \"" + Filter(strs[i], cased) + "\",");
                    fw.WriteLine($"        \"app_name\": \"{appName}\",");
                    fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                    fw.WriteLine("    },");
                }
            }
        kbd:
            if (strs[0] == "kbd")
            {
                for (int i = 1; i < strs.Length; i++)
                {
                    if (strs[i] == "kp" || strs[i] == "kh")
                    {
                        input = "";
                        for (; i < strs.Length; i++)
                        {
                            input += strs[i] + ",";
                        }
                        strs = input.Split(',');
                        goto kpkh;
                    }
                    if (strs[i] == "kbd") continue;
                    if (strs[i] == "") continue;
                    strs[i] = Filter(strs[i], cased);
                    typed += strs[i];
                }
            }
            if (strs[0] == "kc")
            {
                fw.WriteLine("    {");
                fw.WriteLine("        \"type\": \"record\",");
                fw.WriteLine("        \"origin\": \"desktop\",");
                fw.WriteLine("        \"description\": \"Keyboard Shortcut\",");
                fw.WriteLine("        \"actionType\":\"" + "kbd_shortcut" + "\",");
                fw.WriteLine($"        \"pageScreenshot\": \"{scr}\",");
                fw.Write("        \"value\": [\"" + Filter(strs[1], cased));
                for (int i = 2; i < strs.Length; i++)
                {
                    if (strs[i] == "") continue;
                    fw.Write("\", \"" + Filter(strs[i], cased));
                }
                fw.Write("\"]," + Environment.NewLine);
                fw.WriteLine($"        \"app_name\": \"{appName}\",");
                fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
                fw.WriteLine("    },");
            }
        }

        private double parse_Timestamp()
        {
            return Math.Round((TimeZoneInfo.ConvertTimeToUtc(DateTime.Now) - new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds);
        }

        private void timecnt_Tick(object sender, EventArgs e)
        {
            timecnt.Stop();
            fw.WriteLine("    {");
            fw.WriteLine("        \"type\": \"record\",");
            fw.WriteLine("        \"origin\": \"desktop\",");
            fw.WriteLine("        \"description\": \"Click\",");
            fw.WriteLine("        \"selectorType\": \"imgSelector\",");
            fw.WriteLine("        \"actionType\":\"click\",");
            fw.WriteLine($"        \"image\": \"{img}\",");
            fw.WriteLine($"        \"pageScreenshot\": \"{scr}\",");
            fw.WriteLine($"        \"app_name\": \"{appName}\",");
            fw.WriteLine("        \"coord\": {");
            fw.WriteLine("            \"x\": " + p.pt.x + ",");
            fw.WriteLine("            \"y\": " + p.pt.y);
            fw.WriteLine("        },");
            fw.WriteLine($"        \"timeStamp\": {parse_Timestamp()}");
            fw.WriteLine("    },");
            fn = false;
        }

        public void finish(string path, dynamic executeoutputobject)
        {
            if (path == "" || executeoutputobject == null)
            {
                cancel();
                return;
            }
            while (timecnt.Enabled)
            {
                Thread.Sleep(100);
            }

            fw.Write("]");
            File.AppendAllText(tempPath, $"logging stopped{Environment.NewLine}");
            
            fp.Seek(-4, SeekOrigin.Current);
            fw.Write(" ");
            fp.Close();
            if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat")))
            {
                File.AppendAllText(tempPath, $"mn926 is not there {Environment.NewLine}");
                cancel();
                return;
            }
            if (imgDict.Values.Where(a => a == "").Count() > 0 || scrDict.Values.Where(a => a == "").Count() > 0)
            {
                Thread.Sleep(10000);
            }
            var prejson = File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"));
            if (prejson.StartsWith("[") && prejson.EndsWith("]"))
            {
                File.AppendAllText(tempPath, $"mn926 file is read{Environment.NewLine}");
                var desktopRecordArray = JsonConvert.DeserializeObject<dynamic[]>(prejson);
                File.AppendAllText(tempPath, $"string to json of mn926{Environment.NewLine}");
                for (int i = 0; i < desktopRecordArray.Length; i++)
                {
                    JValue tp = desktopRecordArray[i].actionType;
                    if ((string)tp.Value == "kbd_shortcut")
                    {
                        JValue fi = desktopRecordArray[i].pageScreenshot;
                        desktopRecordArray[i].pageScreenshot = scrDict[Int32.Parse((string)fi.Value)];
                    }
                    else if ((string)tp.Value == "click" || (string)tp.Value == "rclick" || (string)tp.Value == "dbclick")
                    {
                        JValue im = desktopRecordArray[i].image;
                        JValue fi = desktopRecordArray[i].pageScreenshot;
                        desktopRecordArray[i].image = imgDict[Int32.Parse((string)im.Value)];
                        desktopRecordArray[i].pageScreenshot = scrDict[Int32.Parse((string)fi.Value)];
                    }
                    else if ((string)tp.Value == "kp")
                    {
                        JValue vl = desktopRecordArray[i].value;
                        if ((string)vl.Value == "return")
                        {
                            desktopRecordArray[i].value = "enter";
                        }
                    }
                }
                File.AppendAllText(tempPath, $"image to url{Environment.NewLine}");
                File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"));
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"), JsonConvert.SerializeObject(desktopRecordArray, Formatting.Indented));
            }
            else
            {
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"), "");
            }
            File.AppendAllText(tempPath, $"write total data to mn926{Environment.NewLine}");
            string outputDesktopPath = Path.Combine(path, "output.json");
            string outputBrowserPath = Path.Combine(path, "outputChrome.json");

            if (File.Exists(outputDesktopPath)) File.Delete(outputDesktopPath);
            File.Copy(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"), outputDesktopPath);
            File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"));
            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images")))
            {
                #region To Save and Store Images
                //if (Directory.Exists(Path.Combine(Path.GetDirectoryName(dpath), @"Images"))) Directory.Delete(Path.Combine(Path.GetDirectoryName(dpath), @"Images"), true);
                //Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(dpath), @"Images"));
                //foreach (string file in Directory.GetFiles(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images")))
                //{
                //    string name = Path.GetFileName(file);
                //    string dest = Path.Combine(Path.Combine(Path.GetDirectoryName(dpath), @"Images"), name);
                //    File.Copy(file, dest);
                //}
                #endregion
                Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"), true);
            }
            File.AppendAllText(tempPath, $"clear temp data{Environment.NewLine}");


            string extnOUTcontents = JsonConvert.SerializeObject(executeoutputobject.data); //============================= ERROR HERE ===================================
                                                                                            //============================ Line no. 827 ==================================


            File.AppendAllText(tempPath, $"{Environment.NewLine}");
            using (StreamWriter writer = new StreamWriter(outputBrowserPath))
            {
                writer.Write(extnOUTcontents);
            }

            ConversionResult resultset = ConversionScript.PrepareRpaPayload(outputDesktopPath, outputBrowserPath);
            File.AppendAllText(tempPath, $"conversion to rpa json{Environment.NewLine}");
            Dictionary<string, object> payload = new Dictionary<string, object> {
                                            { "id", executeoutputobject.skillId.Value},
                                            { "type", "record"},
                                            { "actions", JsonConvert.DeserializeObject(resultset.Data)}
                                        };
            File.AppendAllText(tempPath, $"payload{Environment.NewLine}");
            Dictionary<string, object> dictObj = new Dictionary<string, object> {
                                            { "markerIndex",executeoutputobject.markerIndex.Value},
                                            { "skillId", executeoutputobject.skillId.Value },
                                            { "nodeId", executeoutputobject.nodeId.Value},
                                            { "payload",payload }
                                        };
            File.AppendAllText(tempPath, $"ovrerall payload{Environment.NewLine}");
            UploadRecordedEventsToSuper(dictObj);
        }

        private void cancel()
        {
            if (timecnt.Enabled) timecnt.Stop();
            if (fp != null) fp.Close();
            if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat")) && Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images")))
            {
                File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @".\Temp\mn926.dat"));
                Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"), true);
            }
        }

        private string takeSnap(int x, int y)
        {
            string filename = @"image_" + m.ToString().PadLeft(2, '0') + @".jpg";
            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Images\" + filename;
            if (!Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"))) Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"));
            Bitmap bmp = new Bitmap(60, 70, PixelFormat.Format32bppArgb);
            Graphics grp = Graphics.FromImage(bmp);
            grp.CopyFromScreen(x - 30, y - 35, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);
            bmp.Save(filepath, ImageFormat.Jpeg);
            m++;
            return UploadImageToSuper(filepath);
        }

        private string takeScreen()
        {
            string filename = @"full_image_" + l.ToString().PadLeft(2, '0') + @".jpg";
            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Images\" + filename;
            if (!Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"))) Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Images"));
            Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
            /*string tempPath = Path.GetTempPath();
            string screenDimensionfile = $"screenDimensionfile.txt";
            screenDimensionfile = Path.Combine(tempPath, screenDimensionfile);
            File.AppendAllText(screenDimensionfile,  $"{Screen.PrimaryScreen.Bounds.Width} , {Screen.PrimaryScreen.Bounds.Height}{ Environment.NewLine}");*/
            
            Graphics grp = Graphics.FromImage(bmp);
            grp.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy);
            bmp.Save(filepath, ImageFormat.Jpeg);
            l++;
            return UploadImageToSuper(filepath);
        }

        private string UploadImageToSuper(string filename)
        {
            HttpClient httpClient = new HttpClient();
            string boundary = "--------------------------" + DateTime.Now.Ticks.ToString() + DateTime.Now.Ticks.ToString().Substring(0, 6);
            MultipartFormDataContent form = new MultipartFormDataContent(boundary);
            var file_bytes = File.ReadAllBytes(filename);
            form.Add(new ByteArrayContent(file_bytes, 0, file_bytes.Length), "file", filename);
            var postUrl = new Uri("https://superapiqa.development.techforce.ai/botapi/upload");
            dynamic configObject = JsonConvert.DeserializeObject(File.ReadAllText(Path.Combine(RuntimePath, "../../Configs/extnHost_chrome/config.json")));
            string refreshtoken = configObject.token;
            string accesstoken = GetStoredAccessToken(refreshtoken);//"eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICJ0TXVIRlRmNjZYRlFGWEtPX3JlVF96ajlVTkRyLW1wX0JYQnpZSHZsYVd3In0.eyJleHAiOjE2NTc2MDE5OTAsImlhdCI6MTY0OTgyNTk5MCwiYXV0aF90aW1lIjoxNjQ5NzU1Njk2LCJqdGkiOiI4MjFiZmZkYS1jYjJmLTRkZGYtYTY0My1lOWU0ZjA3ZjgxYTUiLCJpc3MiOiJodHRwczovL2VtY2F1dGhhd3MudGVjaGZvcmNlLmFpL2F1dGgvcmVhbG1zL3RlY2hmb3JjZSIsImF1ZCI6ImFjY291bnQiLCJzdWIiOiI1NzViNDJjYi01YmViLTQ1NDgtOTU0OC02NGU5NWJmMDk5NTYiLCJ0eXAiOiJCZWFyZXIiLCJhenAiOiJzdXBlci1leHRlbnNpb24iLCJub25jZSI6ImIzNmViODRlLTk5Y2UtNDUwYS04YmZiLTEzN2NjMGUxYWMzNCIsInNlc3Npb25fc3RhdGUiOiIyN2E5M2EzZi1jY2NmLTRjYzYtOTAxNy02N2YwOGU2ZmE2NTEiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImNocm9tZS1leHRlbnNpb246Ly9sbmdocGlpZXBmYmdnY2trb2dtam1wYmhubmFkbG1lcCIsImNocm9tZS1leHRlbnNpb246Ly9ka2FrZmpsbG9lZGNtam5na2lvbWdtZWdmY2ZnbWdvZiIsImNocm9tZS1leHRlbnNpb246Ly9scGNkaWlpbmJnbm1wbmliYWFram9kYWxnaHBwaGtrYy8qIiwiY2hyb21lLWV4dGVuc2lvbjovL2duYWlnZmFoZmxkbGpoZGhnbGFqb2JoaGhvcG9tbG9nIiwiY2hyb21lLWV4dGVuc2lvbjovL2lrYmNnZm1qa21maWptZmtvbGtpYWptb2Znam9qbG5oIiwiY2hyb21lLWV4dGVuc2lvbjovL2RkaGhsb25pZGFma2xoY2dvb2drZWVjbmZsZ2dqZGVlIiwiY2hyb21lLWV4dGVuc2lvbjovL25qamhiZ2lwaHBvZGNnY2RraWtrb2RwcG5oYmRpYXBtIiwiY2hyb21lLWV4dGVuc2lvbjovL25oY2JlZWxhbWtncG5oYWxiYWNtaWthYmNtZGRpZWNoIiwiY2hyb21lLWV4dGVuc2lvbjovL2xicGhsYWpncGpkbGRlaWNvaG5lZ2FrYm1mZmZoY2ZlIiwiY2hyb21lLWV4dGVuc2lvbjovL21jZmdna2RmZGRuZGtsY2NoYWhsZ2pjZ2Vrb2xnYW9jIiwiY2hyb21lLWV4dGVuc2lvbjovL21vYWhwZW5na2tmb25qYW5nZGlvYWtiZ2xtbmVnbWdtIiwiY2hyb21lLWV4dGVuc2lvbjovL2hvcG9ma2Nrb2pnaG9qZG5wcGNnamRiZGtoaWtsbmNqIiwiY2hyb21lLWV4dGVuc2lvbjovL2ljZWVmZWFrb2tnZGlwam1oZWVlZmhqZWhubGRpb29qIiwiY2hyb21lLWV4dGVuc2lvbjovL3BubmNibG1wZGlrbmxvbXBra2Nha2pvYWhwcGtmZGVqIiwiY2hyb21lLWV4dGVuc2lvbjovL2NpZGRwcGhwaWtsZ2pnZmRkbGduaGZpYmJjb2xubGxrIiwiY2hyb21lLWV4dGVuc2lvbjovL2JmbmVqaG1vb2Nsa2JmYW5sYWluZGpna2RhbGxpa21qIiwiY2hyb21lLWV4dGVuc2lvbjovL21pamFibmlmbmRjb2VnbWttYWhmYWNjZWJjaWdtcGFuIiwiY2hyb21lLWV4dGVuc2lvbjovL2pqaGxkbWpoYm5mampwYmxqam9kZWlqbHBsb2ZqYm9lIiwiY2hyb21lLWV4dGVuc2lvbjovL2dwZG5tZm9hZG5saGlqZGljZWNpbWtsb21wamljaWdtIl0sInJlYWxtX2FjY2VzcyI6eyJyb2xlcyI6WyJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYWNjb3VudCI6eyJyb2xlcyI6WyJtYW5hZ2UtYWNjb3VudCIsIm1hbmFnZS1hY2NvdW50LWxpbmtzIiwidmlldy1wcm9maWxlIl19fSwic2NvcGUiOiJvcGVuaWQgcHJvZmlsZSBlbWFpbCIsImVtYWlsX3ZlcmlmaWVkIjp0cnVlLCJuYW1lIjoiUHJlZXRoaSBQb3JlZGR5IiwicHJlZmVycmVkX3VzZXJuYW1lIjoic3BvcmVkZHlAdGVjaGZvcmNlLmFpIiwiZ2l2ZW5fbmFtZSI6IlByZWV0aGkiLCJmYW1pbHlfbmFtZSI6IlBvcmVkZHkiLCJlbWFpbCI6InNwb3JlZGR5QHRlY2hmb3JjZS5haSJ9.DrbZZyGvE_tV4C_GZHc-QxfbSYS87bD27-X37v8rg4le-cOsaenYyk6xvKOXlu87o5mTSd1pjDwJ21sxtBxdSMJg7PyoT3FgYOfmqccQK-JUwKW7EGXayYsMj5HptdNl5E6DLB95Xh7kFstUO65QZ5IDyMp862fAFLvkK7ATe8D4r3QYh3b6jvkmMXSA-ZmfIwMbG90OFVLeoSDdPKLFe95yvVZXFvawI_CXB0vQU13EW4XJ6wfBuZh0RMVW0DZ0yUlclQ1nt9_G5blMowtJ4Dgv5DLL9_Yhz5BmqS0yQuN_YbnY3iLcGXWtW9NQSdFbVQIQ-JGZY7g6ddN_13QpAg";
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", string.Format("Bearer {0}", accesstoken));
            var boundaryParameter = form.Headers.ContentType.Parameters.Single(p => p.Name == "boundary");
            boundaryParameter.Value = boundaryParameter.Value.Replace("\"", "");
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HttpResponseMessage response = httpClient.PostAsync(postUrl, form).Result;
            response.EnsureSuccessStatusCode();
            httpClient.Dispose();
           // File.AppendAllText(tempPath, $"{response.EnsureSuccessStatusCode()}--{Environment.NewLine}");
            File.AppendAllText(tempPath, $"uploaded image to super{Environment.NewLine}");
            string sd = response.Content.ReadAsStringAsync().Result;
            dynamic imageRe = JsonConvert.DeserializeObject(sd);
            File.AppendAllText(tempPath, $"{sd}--{Environment.NewLine}");
            string FinalUrl = "https://superapiqa.development.techforce.ai/botapi/image/" + imageRe.data.image_id.Value;
            File.AppendAllText(tempPath, $"{FinalUrl}--{Environment.NewLine}");
            return FinalUrl;            
        }

        public static string GetAccessTokenFromRefreshToken(string refreshtoken)
        {
            var postUrl = new Uri("https://superssoqa.development.techforce.ai/auth/realms/techforce/protocol/openid-connect/token");
            Dictionary<string, string> dictObj = new Dictionary<string, string> {
                                { "client_id","super-extension" },
                                {"refresh_token",refreshtoken},
                                {"grant_type","refresh_token" }
                            };
            HttpClient httpClient = new HttpClient();
            var content = new FormUrlEncodedContent(dictObj);
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/x-www-form-urlencoded");
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HttpResponseMessage response = httpClient.PostAsync(postUrl, content).Result;
            string res = response.Content.ReadAsStringAsync().Result;
            dynamic resp = JsonConvert.DeserializeObject(res);
            IsAccessTokenStored = true;
            StoredAccessToken = resp.access_token;
            return resp.access_token;
        }
        public static string GetStoredAccessToken(string refreshToken)
        {
            if (File.Exists(Path.Combine(Path.GetTempPath(), "StoredAccessToken.txt")))
            {
                StoredAccessToken = File.ReadAllText(Path.Combine(Path.GetTempPath(), "StoredAccessToken.txt"));
                var getUrl = new Uri("https://superssoqa.development.techforce.ai/auth/realms/techforce/protocol/openid-connect/userinfo");
                HttpClient httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", string.Format("Bearer {0}", StoredAccessToken));
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HttpResponseMessage response = httpClient.GetAsync(getUrl).Result;
                if (response.IsSuccessStatusCode)
                {
                    return StoredAccessToken;
                }
                else
                {
                    return GetAccessTokenFromRefreshToken(refreshToken);
                }
            }
            else
            {
                return GetAccessTokenFromRefreshToken(refreshToken);
            }

        }
        private bool isWindowExcluded()
        {
            string processName = GetActiveApplicationName();
            appName = processName;
            return (processName == "chrome" || processName == "firefox" || processName == "opera" || processName == "msedge" || processName == "iexplore");
        }

        private string GetActiveApplicationName()
        {
            uint pID;
            string processName;
            GetWindowThreadProcessId(GetForegroundWindow(), out pID);
            processName = Process.GetProcessById((int)pID).ProcessName;
            return processName;
        }

        public void Dispose()
        {
            if (!ishooked) this.Dispatcher.InvokeShutdown();
        }
    }
}
