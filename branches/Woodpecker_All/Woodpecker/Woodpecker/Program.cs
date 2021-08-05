using System;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using ModuleLayer;
using log4net.Config;
using log4net;
using System.Reflection;

namespace Woodpecker
{
    static class Program
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            XmlConfigurator.Configure(new System.IO.FileInfo("./log4net.config"));      //log4net configure file
            //高Dpi設定
            if (Environment.OSVersion.Version.Major >= 6) { SetProcessDPIAware(); }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //Thread to show splash window
            Thread thUI = new Thread(new ThreadStart(ShowSplashWindow))
            {
                Name = "Splash UI",
                Priority = ThreadPriority.Highest,
                IsBackground = true
            };
            thUI.Start();

            //Thread to load time-consuming resources.
            Thread th = new Thread(new ThreadStart(LoadResources))
            {
                Name = "Resource Loader",
                Priority = ThreadPriority.Normal
            };
            th.Start();
            th.Join();

            if (SplashForm != null)
            {
                SplashForm.Invoke(new MethodInvoker(delegate { SplashForm.Close(); }));
            }
            thUI.Join();
            if (args.Length == 0)
            {
                Application.Run(new Form1());
            }
            else
            {
                //印出程式的名稱
                Console.WriteLine(AppDomain.CurrentDomain.FriendlyName);
                //印出傳入的參數
                Console.WriteLine(args[0].ToString());

                Application.Run(new Form1(args[0].ToString()));
            }
        }

        public static frm_Splash SplashForm
        {
            get;
            set;
        }

        private static void LoadResources()
        {
            for (int i = 1; i <= 10; i++)
            {
                /*if (SplashForm != null)
                {SplashForm.Invoke(new MethodInvoker(delegate 
                        {SplashForm.labelMark.Text = "Spider";}));}*/
                Thread.Sleep(100);
            }
            
            Add_ons Add_ons = new Add_ons();
            Add_ons.CreateConfig();//如果根目錄沒有Config.ini則創建//
            Add_ons.USB_Read();//讀取USB設備的Pid, Vid//

            //Add_ons.CreateExcelFile();
        }

        private static void ShowSplashWindow()
        {
            SplashForm = new frm_Splash();
            Application.Run(SplashForm);
        }

    }

    public static class GlobalData
    {   //global variables and classes//
        public static string MainSettingPath = Application.StartupPath + "\\Config.ini";
        public static string MailSettingPath = Application.StartupPath + "\\Mail.ini";
        public static string RcSettingPath = Application.StartupPath + "\\RC.ini";
        public static string StartupPath = Application.StartupPath;
        public static string MeasurePath = Application.StartupPath;
        public static ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);        //log4net

        public static int Scheduler_Row = 0;
        public static List<string> VidList = new List<string>();     //List is an Object inherited from System.Collections.Generic
        public static List<string> PidList = new List<string>();
        public static List<string> AutoBoxComPort_List = new List<string>();
        public static List<string> RcList = new List<string>();
        public static List<string> sourceList;
        public static int Schedule_Number = 0;
        public static int Schedule_1_Exist = 0;
        public static int Schedule_2_Exist = 0;
        public static int Schedule_3_Exist = 0;
        public static int Schedule_4_Exist = 0;
        public static int Schedule_5_Exist = 0;
        public static long Schedule_1_TestTime = 0;
        public static long Schedule_2_TestTime = 0;
        public static long Schedule_3_TestTime = 0;
        public static long Schedule_4_TestTime = 0;
        public static long Schedule_5_TestTime = 0;
        public static long Total_Test_Time = 0;
        public static int Loop_Number = 0;
        public static int Total_Loop = 0;
        public static int Schedule_Loop = 999999;
        public static int Schedule_Step;
        public static int caption_Num = 0;
        public static int caption_Sum = 0;
        public static int excel_Num = 0;
        public static int[] caption_NG_Num = new int[Schedule_Loop];
        public static int[] caption_Total_Num = new int[Schedule_Loop];
        public static float[] SumValue = new float[Schedule_Loop];
        public static int[] NGValue = new int[GlobalData.Schedule_Loop];
        public static float[] NGRateValue = new float[GlobalData.Schedule_Loop];
        //public static float[] ReferenceResult = new float[Schedule_Loop];
        public static bool FormSetting = true;
        public static bool FormSchedule = true;
        public static bool FormMail = true;
        public static bool FormLog = true;
        public static string RCDB = "";
        public static string IO_INPUT = "2,2,2,2,2,2";
        public static string Arduino_IO_INPUT = "Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,";
        public static uint IO_INPUT_value = 0x40;
        public static uint Arduino_IO_INPUT_value = 0x10000;
        public static string Arduino_Comport = "";
        public static string Arduino_Baudrate = "";
        public static string Arduino_Read_String = "";
        public static bool Arduino_outputFlag = false;
        public static bool Arduino_recFlag = false;
        public static int IO_PA10_0_COUNT = 0;
        public static int IO_PA10_1_COUNT = 0;
        public static int IO_PA11_0_COUNT = 0;
        public static int IO_PA11_1_COUNT = 0;
        public static int IO_PA14_0_COUNT = 0;
        public static int IO_PA14_1_COUNT = 0;
        public static int IO_PA15_0_COUNT = 0;
        public static int IO_PA15_1_COUNT = 0;
        public static int IO_PB1_0_COUNT = 0;
        public static int IO_PB1_1_COUNT = 0;
        public static int IO_PB7_0_COUNT = 0;
        public static int IO_PB7_1_COUNT = 0;
        public static int IO_Arduino2_0_COUNT = 0;
        public static int IO_Arduino2_1_COUNT = 0;
        public static int IO_Arduino3_0_COUNT = 0;
        public static int IO_Arduino3_1_COUNT = 0;
        public static int IO_Arduino4_0_COUNT = 0;
        public static int IO_Arduino4_1_COUNT = 0;
        public static int IO_Arduino5_0_COUNT = 0;
        public static int IO_Arduino5_1_COUNT = 0;
        public static int IO_Arduino6_0_COUNT = 0;
        public static int IO_Arduino6_1_COUNT = 0;
        public static int IO_Arduino7_0_COUNT = 0;
        public static int IO_Arduino7_1_COUNT = 0;
        public static int IO_Arduino8_0_COUNT = 0;
        public static int IO_Arduino8_1_COUNT = 0;
        public static int IO_Arduino9_0_COUNT = 0;
        public static int IO_Arduino9_1_COUNT = 0;
        public static string keyword_1 = "false";
        public static string keyword_2 = "false";
        public static string keyword_3 = "false";
        public static string keyword_4 = "false";
        public static string keyword_5 = "false";
        public static string keyword_6 = "false";
        public static string keyword_7 = "false";
        public static string keyword_8 = "false";
        public static string keyword_9 = "false";
        public static string keyword_10 = "false";
        public static int Rc_Number = 0;
        public static string Pass_Or_Fail = "";//測試結果//
        public static int Break_Out_Schedule = 0;//定時器中斷變數//
        public static int Break_Out_MyRunCamd;//是否跳出倒數迴圈，1為跳出//
        public static bool FormRC = false;
        public static bool FormGPIO = false;
        public static int TEXTBOX_FOCUS = 0;
        public static string label_Command = "";
        public static string label_Remark = "";
        public static string label_LoopNumber = "";
        public static bool VideoRecording = false;
        public static string srtstring = "";
        public static bool StartButtonPressed = false;//true = 按下START//false = 按下STOP//
        //public static PortConfigGroup portConfigGroup_A, portConfigGroup_B, portConfigGroup_C, portConfigGroup_D, portConfigGroup_E, portConfigGroup_Kline;
        public static PortConfigGroup portConfigGroup_A = new PortConfigGroup();
        public static PortConfigGroup portConfigGroup_B = new PortConfigGroup();
        public static PortConfigGroup portConfigGroup_C = new PortConfigGroup();
        public static PortConfigGroup portConfigGroup_D = new PortConfigGroup();
        public static PortConfigGroup portConfigGroup_E = new PortConfigGroup();
        public static PortConfigGroup portConfigGroup_Kline = new PortConfigGroup();
        public static List<PortConfigGroup> _portConfigList = new List<PortConfigGroup>() { portConfigGroup_A, portConfigGroup_B, portConfigGroup_C, portConfigGroup_D, portConfigGroup_E, portConfigGroup_Kline };

        public static Mod_RS232 m_SerialPort_A = new Mod_RS232();
        public static Mod_RS232 m_SerialPort_B = new Mod_RS232();
        public static Mod_RS232 m_SerialPort_C = new Mod_RS232();
        public static Mod_RS232 m_SerialPort_D = new Mod_RS232();
        public static Mod_RS232 m_SerialPort_E = new Mod_RS232();
        public static Mod_RS232 m_Arduino_Port = new Mod_RS232();

        public static string logAllText;
        public static string Measure_Backlight = "None";
        public static string Measure_Thermal = "None";
        //MessageBox.Show("RC Key is empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Question);//MessageBox範例

        /*public static GlobalData()
        {
            VidList = new List<string>();
        }*/
    }

    public class PortConfigGroup
    {
        public bool checkedValue;
        public string portLabel;
        public string portConfig;
        public string portName;
        public string portBR;
        public byte portLF;
    }
}
