using BlueRatLibrary;
using DirectX.Capture;
using jini;
using Microsoft.Win32.SafeHandles;      //support SafeFileHandle
using RedRat.IR;
using RedRat.RedRat3;
using RedRat.RedRat3.USB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;                            //support BindingFlags
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using System.Xml.Linq;
using USBClassLibrary;
using System.Net.Sockets;
using System.Net;
using BlockMessageLibrary;
using DTC_ABS;
using DTC_OBD;
using MySerialLibrary;
using KWP_2000;
using USB_CAN2C;
//using MaterialSkin.Controls;
//using MaterialSkin;
using System.ComponentModel;
using Microsoft.VisualBasic.FileIO;
using USB_VN1630A;
using ModuleLayer;
using log4net;
using Universal_Toolkit;
using Universal_Toolkit.Types;
//using NationalInstruments.DAQmx;

namespace Woodpecker
{
    public partial class Form1 : Form
    {
        private string _args;
        //private BackgroundWorker BackgroundWorker = new BackgroundWorker();
        //private Form_DGV_Autobox Form_DGV_Autobox = new Form_DGV_Autobox();
        //private TextBoxBuffer textBoxBuffer = new TextBoxBuffer(4096);

        private string MainSettingPath = GlobalData.MainSettingPath;    //Application.StartupPath + "\\Config.ini";
        private string MailPath = GlobalData.MailSettingPath;                  //Application.StartupPath + "\\Mail.ini";
        private string RcPath = GlobalData.RcSettingPath;                         //Application.StartupPath + "\\RC.ini";
        private static ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);        //log4net

        private Mod_RS232 serialPortA = new Mod_RS232();
        private Mod_RS232 serialPortB = new Mod_RS232();
        private Mod_RS232 serialPortC = new Mod_RS232();
        private Mod_RS232 serialPortD = new Mod_RS232();
        private Mod_RS232 serialPortE = new Mod_RS232();
        //private static DrvRS232 serialPortK = new DrvRS232();
        private MySerial MySerialPort = new MySerial();      //from Kline_Serial.cs

        private LogDumpping logDumpping = new LogDumpping();
        private RK2797 rk2797 = new RK2797();
        //Setting FSetting = new Setting();
        private Konica_Minolta CA210 = new Konica_Minolta();

        //宣告於keyword使用
        //public Queue<SerialReceivedData> data_queue;
        //public static Queue<byte> LogQueue_A = new Queue<byte>();
        private Queue<byte> SearchLogQueue_A = new Queue<byte>();
        private Queue<byte> TemperatureQueue_A = new Queue<byte>();
        private Queue<byte> SearchLogQueue_B = new Queue<byte>();
        private Queue<byte> SearchLogQueue_C = new Queue<byte>();
        private Queue<byte> SearchLogQueue_D = new Queue<byte>();
        private Queue<byte> SearchLogQueue_E = new Queue<byte>();
        private char Keyword_SerialPort_A_temp_char;
        private byte Keyword_SerialPort_A_temp_byte;
        private char Keyword_SerialPort_B_temp_char;
        private byte Keyword_SerialPort_B_temp_byte;
        private char Keyword_SerialPort_C_temp_char;
        private byte Keyword_SerialPort_C_temp_byte;
        private char Keyword_SerialPort_D_temp_char;
        private byte Keyword_SerialPort_D_temp_byte;
        private char Keyword_SerialPort_E_temp_char;
        private byte Keyword_SerialPort_E_temp_byte;

        private IRedRat3 redRat3 = null;
        private Add_ons Add_ons = new Add_ons();
        private RedRatDBParser RedRatData = new RedRatDBParser();
        private BlueRat MyBlueRat = new BlueRat();
        private static bool BlueRat_UART_Exception_status = false;

        private static void BlueRat_UARTException(Object sender, EventArgs e)
        {
            BlueRat_UART_Exception_status = true;
        }

        private bool FormIsClosing = false;
        private Capture capture = null;
        private Filters filters = null;
        private bool _captureInProgress;
        private bool StartButtonPressed = false;    //true = 按下START//false = 按下STOP//
        //private bool excelstat = false;
        private bool TimerPanel = false;
        //private bool VirtualRcPanel = false;
        private bool AcUsbPanel = false;
        private long timeCount = 0;
        private long TestTime = 0;
        private string videostring = "";
        private string srtstring = "";
        private bool TakePicture = false;
        private int TakePictureError = 0;


        //Schedule暫停用的參數
        private bool Pause = false;
        private ManualResetEvent SchedulePause = new ManualResetEvent(true);
        private ManualResetEvent ScheduleWait = new ManualResetEvent(true);

        private SafeDataGridView portos_online;
        private int Breakpoint;
        private int Nowpoint;
        private bool Breakfunction = false;
        //private const int CS_DROPSHADOW = 0x20000;      //宣告陰影參數
		
        private List<BlockMessage> MyBlockMessageList = new List<BlockMessage>();
        private ProcessBlockMessage MyProcessBlockMessage = new ProcessBlockMessage();

        //拖動無窗體的控件>>>>>>>>>>>>>>>>>>>>
        [DllImport("user32.dll")]
        new public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        new public static extern bool SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MOVE = 0xF010;
        public const int HTCAPTION = 0x0002;
        //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        //CanReader
        private CAN_USB2C Can_Usb2C = new CAN_USB2C();
        private USB_VECTOR_Lib Can_1630A = new USB_VECTOR_Lib();
        private int can_send = 0;
        private List<USB_CAN2C.CAN_Data> can_data_list = new List<USB_CAN2C.CAN_Data>();
        private bool set_timer_rate = false;
        private uint can_id;
        private Dictionary<uint, uint> can_rate = new Dictionary<uint, uint>();
        private Dictionary<uint, byte[]> can_data = new Dictionary<uint, byte[]>();

        //Klite error code
        public int kline_send = 0;
        public List<DTC_Data> ABS_error_list = new List<DTC_Data>();
        public List<DTC_Data> OBD_error_list = new List<DTC_Data>();

        //Serial Port Parameters
        public delegate void AddDataDelegate(String myString);
        public AddDataDelegate myDelegate1;
        private string logA_text = "", logB_text = "", logC_text = "", logD_text = "", logE_text = "", minolta_text = "", arduino_text = "", ftdi_text = "", ca310_text = "", canbus_text = "", kline_text = "", logAll_text = "", debug_text = "",
                       minolta_csv_report = "Sx, Sy, Lv, T, duv, Display mode, X, Y, Z, Date, Time, Scenario, Now measure count, Target measure count, Backlight sensor, Thanmal sensor, \r\n";		   
        public string portLabel_A = "Port A", portLabel_B = "Port B", portLabel_C = "Port C", portLabel_D = "Port D", portLabel_E = "Port E", portLabel_K = "Kline", portLabel_Arduino = "Arduino";
        public string serialPortConfig_A = "PortA", serialPortConfig_B = "PortB", serialPortConfig_C = "PortC", serialPortConfig_D = "PortD", serialPortConfig_E = "PortE", serialPortConfig_Arduino = "Arduino";
        public string serialPortName_A, serialPortName_B, serialPortName_C, serialPortName_D, serialPortName_E, serialPortName_Arduino;
        public string serialPortBR_A, serialPortBR_B, serialPortBR_C, serialPortBR_D, serialPortBR_E, serialPortBR_Arduino;
        private int log_max_length = 10000000, debug_max_length = 10000000;
        private FTDI_Lib Ftdi_lib = new FTDI_Lib();
        private PortInfo portinfo = new PortInfo();
        //Search temperature parameter
        List<Temperature_Data> temperatureList = new List<Temperature_Data> { };
        Queue<double> temperatureDouble = new Queue<double> { };

        byte[] dataset;
        string strValues1 = string.Empty;
        double currentTemperature = 0;
        System.Timers.Timer duringTimer = new System.Timers.Timer();

        bool ifStatementFlag = false;
        bool ChamberIsFound = false;
        bool TemperatureIsFound = false;
        bool PowerSupplyIsFound = false;
        string MaxTemperature = "", MinTemperature = "";
        string expectedVoltage = string.Empty;
        string PowerSupplyCommandLog = string.Empty;

        public Form1()
        {
            InitializeComponent();
            //setStyle();

            //Datagridview design
            DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].DefaultCellStyle.BackColor = Color.FromArgb(56, 56, 56);
            DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].DefaultCellStyle.ForeColor = Color.FromArgb(255, 255, 255);
            DataGridView_Schedule.Columns[0].DefaultCellStyle.BackColor = Color.FromArgb(56, 56, 56);
            DataGridView_Schedule.Columns[0].DefaultCellStyle.ForeColor = Color.FromArgb(255, 255, 255);

            InitComboboxSaveLog();

            //USB Connection//
            USBPort = new USBClass();
            USBDeviceProperties = new USBClass.DeviceProperties();
            USBPort.USBDeviceAttached += new USBClass.USBDeviceEventHandler(USBPort_USBDeviceAttached);
            USBPort.USBDeviceRemoved += new USBClass.USBDeviceEventHandler(USBPort_USBDeviceRemoved);
            USBPort.RegisterForDeviceChange(true, this);
            //USBTryBoxConnection();
            USBTryRedratConnection();
            USBTryCameraConnection();
            //MyUSBBoxDeviceConnected = false;
            MyUSBRedratDeviceConnected = false;
            MyUSBCameraDeviceConnected = false;
        }

        public Form1(string value)
        {
            InitializeComponent();
            //setStyle();

            if (!string.IsNullOrEmpty(value))
            {
                _args = value;
            }
            USBPort = new USBClass();
            USBDeviceProperties = new USBClass.DeviceProperties();
            USBPort.USBDeviceAttached += new USBClass.USBDeviceEventHandler(USBPort_USBDeviceAttached);
            USBPort.USBDeviceRemoved += new USBClass.USBDeviceEventHandler(USBPort_USBDeviceRemoved);
            USBPort.RegisterForDeviceChange(true, this);
            //USBTryBoxConnection();
            USBTryRedratConnection();
            USBTryCameraConnection();
            //MyUSBBoxDeviceConnected = false;
            MyUSBRedratDeviceConnected = false;
            MyUSBCameraDeviceConnected = false;
        }

        private void InitComboboxSaveLog()
        {
            List<string> portList = new List<string> { portLabel_A, portLabel_B, portLabel_C, portLabel_D, portLabel_E, portLabel_K, portLabel_Arduino, "Canbus" };
            foreach (string port in portList)
            {
                if (ini12.INIRead(MainSettingPath, port, "Checked", "") == "1")
                {
                    comboBox_savelog.Items.Add(port);
                }
                else if (ini12.INIRead(MainSettingPath, port, "Checked", "") == "0" || ini12.INIRead(MainSettingPath, port, "Checked", "") == "")
                {
                    comboBox_savelog.Items.Remove(port);
                }
            }

            if (ini12.INIRead(MainSettingPath, "Device", "ArduinoExist", "") == "1")
                comboBox_savelog.Items.Add("Arduino");

            if (ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "1")
                comboBox_savelog.Items.Add("Minolta");

            if (ini12.INIRead(MainSettingPath, "Canbus", "Log", "") == "1")
                comboBox_savelog.Items.Add("Canbus");

            if (comboBox_savelog.Items.Count > 1)
                comboBox_savelog.Items.Add("Port All");

            if (comboBox_savelog.Items.Count == 0)
            {
                button_savelog.Enabled = false;
                comboBox_savelog.Enabled = false;
            }
            else
            {
                button_savelog.Enabled = true;
                comboBox_savelog.Enabled = true;
                comboBox_savelog.SelectedIndex = 0;
            }
        }

        private void InitPortConfigParam()
        {
            //Initialize Port Config Parameters
            string[] labelArray = { portLabel_A, portLabel_B, portLabel_C, portLabel_D, portLabel_E, portLabel_K};
            string[] configArray = { serialPortConfig_A, serialPortConfig_B, serialPortConfig_C, serialPortConfig_D, serialPortConfig_E, portLabel_K};
            
            bool tst = GlobalData._portConfigList[0].Equals(GlobalData._portConfigList[2]);   //this is used to check the instance of portConfig_A independent or not 
            if (GlobalData._portConfigList.Count == labelArray.Length && GlobalData._portConfigList.Count == configArray.Length)
            {
                int i = 0;
                foreach (var portConfig in GlobalData._portConfigList)
                {
                    portConfig.portLabel = labelArray[i];
                    portConfig.portConfig = configArray[i];
                    if (ini12.INIRead(MainSettingPath, portConfig.portLabel, "Checked", "") == "" || ini12.INIRead(MainSettingPath, portConfig.portLabel, "Checked", "") == "0")
                    {
                        portConfig.checkedValue = false;
                        portConfig.portName = "";
                        portConfig.portBR = "";
                        portConfig.portLF = 0xFF;
                    }
                    else if (ini12.INIRead(MainSettingPath, portConfig.portLabel, "Checked", "") == "1")
                    {
                        portConfig.checkedValue = true;
                        portConfig.portName = ini12.INIRead(MainSettingPath, portConfig.portLabel, "PortName", "");
                        portConfig.portBR = ini12.INIRead(MainSettingPath, portConfig.portLabel, "BaudRate", "");
                        byte byteValue;
                        bool success = byte.TryParse(ini12.INIRead(MainSettingPath, portConfig.portLabel, "LineFeed", ""), out byteValue);
                        if (success)
                            portConfig.portLF = byteValue;
                        else
                            portConfig.portLF = 0x0D;
                    }
                    i++;
                }
            }
            else
            {
                //This is used to check if count of portConfigList is the same as local array size
                Console.WriteLine("[portConfigList] Out of index!!!");
                Application.Exit();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Normal;

            //根據dpi調整視窗尺寸
            Graphics graphics = CreateGraphics();
            float dpiX = graphics.DpiX;
            float dpiY = graphics.DpiY;
            /*if (dpiX == 96 && dpiY == 96)
            {
                this.Height = 600;
                this.Width = 1120;
            }*/
            int intPercent = (dpiX == 96) ? 100 : (dpiX == 120) ? 125 : 150;

            // 針對字體變更Form的大小
            this.Height = this.Height * intPercent / 100;

            // FwVersion
            label_FwVersion.Text = "Ver. " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
            {
                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "1")
                {
                    ConnectAutoBox1();
                }

                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "2" && ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "0" && ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "0")
                {
                    ConnectAutoBox2();
                }

                pictureBox_BlueRat.Image = Properties.Resources.ON;
				if(ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1" && ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "0" && ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "0")
				{
	                GP0_GP1_AC_ON();
	                GP2_GP3_USB_PC();
				}
            }
            else
            {
                pictureBox_BlueRat.Image = Properties.Resources.OFF;
                pictureBox_AcPower.Image = Properties.Resources.OFF;
                button_AcUsb.Enabled = false;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "ArduinoExist", "") == "1")
            {
                if (ini12.INIRead(MainSettingPath, "Device", "ArduinoPort", "") != "")
                {
                    GlobalData.m_Arduino_Port.OpenSerialPort(GlobalData.Arduino_Comport, GlobalData.Arduino_Baudrate, true);
                    pictureBox_ext_board.Image = Properties.Resources.ON;
                }
                else
                {
                    pictureBox_ext_board.Image = Properties.Resources.OFF;
                }
            }

            if (ini12.INIRead(MainSettingPath, "Device", "FTDIExist", "") == "1")
            {
                ConnectFtdi();
                if (portinfo.ftStatus == FtResult.Ok)
                    pictureBox_ftdi.Image = Properties.Resources.ON;
                else
                    pictureBox_ftdi.Image = Properties.Resources.OFF;
            }
            else
            {
                pictureBox_ftdi.Image = Properties.Resources.OFF;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
            {
                OpenRedRat3();
            }
            else
            {
                pictureBox_RedRat.Image = Properties.Resources.OFF;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
            {
                try
                {
                    pictureBox_Camera.Image = Properties.Resources.ON;
                    filters = new Filters();
                    Filter f;

                    comboBox_CameraDevice.Enabled = true;
                    ini12.INIWrite(MainSettingPath, "Camera", "VideoNumber", filters.VideoInputDevices.Count.ToString());

                    for (int c = 0; c < filters.VideoInputDevices.Count; c++)
                    {
                        f = filters.VideoInputDevices[c];
                        comboBox_CameraDevice.Items.Add(f.Name);
                        if (f.Name == ini12.INIRead(MainSettingPath, "Camera", "VideoName", ""))
                        {
                            comboBox_CameraDevice.Text = ini12.INIRead(MainSettingPath, "Camera", "VideoName", "");
                        }
                    }

                    if (comboBox_CameraDevice.Text == "" && filters.VideoInputDevices.Count > 0)
                    {
                        comboBox_CameraDevice.SelectedIndex = filters.VideoInputDevices.Count - 1;
                        ini12.INIWrite(MainSettingPath, "Camera", "VideoIndex", comboBox_CameraDevice.SelectedIndex.ToString());
                        ini12.INIWrite(MainSettingPath, "Camera", "VideoName", comboBox_CameraDevice.Text);
                    }
                    comboBox_CameraDevice.Enabled = false;
                }
                catch (Exception Ex)
                {
                    Console.WriteLine(Ex);
                    MessageBox.Show(Ex.Message.ToString(), "Camera audio open error!");
                }
            }
            else
            {
                pictureBox_Camera.Image = Properties.Resources.OFF;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "1")
            {
                if (CA210.Status() == false)
                {
                    CA210.Connect();
                    pictureBox_Minolta.Image = Properties.Resources.ON;
                }
                else
                    pictureBox_Minolta.Image = Properties.Resources.OFF;
            }
            else
            {
                pictureBox_Minolta.Image = Properties.Resources.OFF;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1")
            {
                if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1")
                {
                    String can_name;
                    List<String> dev_list = Can_Usb2C.FindUsbDevice();
                    can_name = string.Join(",", dev_list);
                    ini12.INIWrite(MainSettingPath, "Canbus", "DevName", can_name);
                    if (ini12.INIRead(MainSettingPath, "Canbus", "DevIndex", "") == "")
                        ini12.INIWrite(MainSettingPath, "Canbus", "DevIndex", "0");
                    if (ini12.INIRead(MainSettingPath, "Canbus", "BaudRate", "") == "")
                        ini12.INIWrite(MainSettingPath, "Canbus", "BaudRate", "500 Kbps");
                    ConnectUsbCAN();
                    pictureBox_can.Image = Properties.Resources.ON;
                }

                if (ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1")
                {
                    ConnectVectorCAN();
                    pictureBox_can.Image = Properties.Resources.ON;
                }
            }
            else
            {
                pictureBox_can.Image = Properties.Resources.OFF;
            }

            if (ini12.INIWrite(MainSettingPath, "Record", "ImportDB", "") == "1")
                button_Analysis.Visible = true;
            else
                button_Analysis.Visible = false;
            
            LoadRCDB();
            
            List<string> SchExist = new List<string> { };
            for (int i = 2; i < 6; i++)
            {
                SchExist.Add(ini12.INIRead(MainSettingPath, "Schedule" + i, "Exist", ""));
            }

            if (SchExist[0] != "")
            {
                if (SchExist[0] == "0")
                    button_Schedule2.Visible = false;
                else
                    button_Schedule2.Visible = true;
            }
            else
            {
                SchExist[0] = "0";
                button_Schedule2.Visible = false;
            }

            if (SchExist[1] != "")
            {
                if (SchExist[1] == "0")
                    button_Schedule3.Visible = false;
                else
                    button_Schedule3.Visible = true;
            }
            else
            {
                SchExist[1] = "0";
                button_Schedule3.Visible = false;
            }

            if (SchExist[2] != "")
            {
                if (SchExist[2] == "0")
                    button_Schedule4.Visible = false;
                else
                    button_Schedule4.Visible = true;
            }
            else
            {
                SchExist[2] = "0";
                button_Schedule4.Visible = false;
            }

            if (SchExist[3] != "")
            {
                if (SchExist[3] == "0")
                    button_Schedule5.Visible = false;
                else
                    button_Schedule5.Visible = true;
            }
            else
            {
                SchExist[3] = "0";
                button_Schedule5.Visible = false;
            }

            GlobalData.Schedule_2_Exist = int.Parse(SchExist[0]);
            GlobalData.Schedule_3_Exist = int.Parse(SchExist[1]);
            GlobalData.Schedule_4_Exist = int.Parse(SchExist[2]);
            GlobalData.Schedule_5_Exist = int.Parse(SchExist[3]);

            button_Pause.Enabled = false;
            button_Schedule.PerformClick();
            button_Schedule1.PerformClick();
            CheckForIllegalCrossThreadCalls = false;
            TopMost = true;
            TopMost = false;

            //setStyle();

            if (ini12.INIRead(MainSettingPath, "Device", "Software", "") == "All")
            {
                this.Text = "Woodpecker";
                label_RedRat.Visible = true;
                pictureBox_RedRat.Visible = true;
                label_Minolta.Visible = true;
                pictureBox_Minolta.Visible = true;
                label_ftdi.Visible = true;
                pictureBox_ftdi.Visible = true;
                button_VirtualRC.Visible = true;
            }

            InitPortConfigParam();
            this.button_Setting.Enabled = true;
        }

        #region -- USB Detect --
        //暫時移除有關盒子的插拔偵測，因為有其他無相關裝置運用到相同的VID和PID
        private bool USBTryBoxConnection()
        {
            if (GlobalData.AutoBoxComPort_List.Count != 0)
            {
                for (int i = 0; i < GlobalData.AutoBoxComPort_List.Count; i++)
                {
                    if (USBClass.GetUSBDevice(
                        uint.Parse("067B", System.Globalization.NumberStyles.AllowHexSpecifier),
                        uint.Parse("2303", System.Globalization.NumberStyles.AllowHexSpecifier),
                        ref USBDeviceProperties,
                        true))
                    {
                        if (GlobalData.AutoBoxComPort_List[i] == "COM15")
                        {
                            BoxConnect();
                        }
                    }
                }
                return true;
            }
            else
            {
                BoxDisconnect();
                return false;
            }
        }

        private bool USBTryRedratConnection()
        {
            if (USBClass.GetUSBDevice(uint.Parse("112A", System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse("0005", System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
            {
                //My Device is attached
                RedratConnect();
                return true;
            }
            else
            {
                RedratDisconnect();
                return false;
            }
        }

        private bool USBTryCameraConnection()
        {
            int DeviceNumber = GlobalData.VidList.Count;
            int VidCount = GlobalData.VidList.Count - 1;
            int PidCount = GlobalData.PidList.Count - 1;

            if (DeviceNumber != 0)
            {
                for (int i = 0; i < DeviceNumber; i++)
                {
                    if (USBClass.GetUSBDevice(uint.Parse(GlobalData.VidList[i], style: System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse(GlobalData.PidList[i], System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
                    {
                        CameraConnect();
                    }
                }
                return true;
            }
            else
            {
                CameraDisconnect();
                return false;
            }
        }

        private void USBPort_USBDeviceAttached(object sender, USBClass.USBDeviceEventArgs e)
        {
            /*
            if (!MyUSBBoxDeviceConnected)
            {
                Console.WriteLine("USBPort_USBDeviceAttached = " + MyUSBBoxDeviceConnected);
                if (USBTryBoxConnection())
                {
                    MyUSBBoxDeviceConnected = true;
                }
            }
            */

            if (!MyUSBRedratDeviceConnected)
            {
                if (USBTryRedratConnection())
                {
                    MyUSBRedratDeviceConnected = true;
                }
            }

            if (!MyUSBCameraDeviceConnected)
            {
                if (USBTryCameraConnection())
                {
                    MyUSBCameraDeviceConnected = true;
                }
            }
        }

        private void USBPort_USBDeviceRemoved(object sender, USBClass.USBDeviceEventArgs e)
        {
            /*
            if (!USBClass.GetUSBDevice(uint.Parse("067B", System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse("2303", System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
            {
                Console.WriteLine("USBPort_USBDeviceRemoved = " + MyUSBBoxDeviceConnected);
                //My Device is removed
                MyUSBBoxDeviceConnected = false;
                USBTryBoxConnection();
            }
            */

            if (!USBClass.GetUSBDevice(uint.Parse("112A", System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse("0005", System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
            {
                //My Redrat is removed
                MyUSBRedratDeviceConnected = false;
                USBTryRedratConnection();
            }
            /*
            if (!USBClass.GetUSBDevice(uint.Parse("045E", System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse("0766", System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false) ||
                !USBClass.GetUSBDevice(uint.Parse("114D", System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse("8C00", System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
            {
                //My Camera is removed
                MyUSBCameraDeviceConnected = false;
                USBTryCameraConnection();
            }
            */
            int DeviceNumber = GlobalData.VidList.Count;

            if (DeviceNumber != 0)
            {
                for (int i = 0; i < DeviceNumber; i++)
                {
                    if (!USBClass.GetUSBDevice(uint.Parse(GlobalData.VidList[i], style: System.Globalization.NumberStyles.AllowHexSpecifier), uint.Parse(GlobalData.PidList[i], System.Globalization.NumberStyles.AllowHexSpecifier), ref USBDeviceProperties, false))
                    {
                        MyUSBCameraDeviceConnected = false;
                        USBTryCameraConnection();
                    }
                }
            }
        }

        private void BoxConnect()       //TO DO: Inset your connection code here
        {
            pictureBox_BlueRat.Image = Properties.Resources.ON;
        }

        private void BoxDisconnect()        //TO DO: Insert your disconnection code here
        {
            pictureBox_BlueRat.Image = Properties.Resources.OFF;
        }

        private void RedratConnect()        //TO DO: Inset your connection code here
        {
            ini12.INIWrite(MainSettingPath, "Device", "RedRatExist", "1");
            pictureBox_RedRat.Image = Properties.Resources.ON;
        }

        private void RedratDisconnect()     //TO DO: Insert your disconnection code here
        {
            ini12.INIWrite(MainSettingPath, "Device", "RedRatExist", "0");
            pictureBox_RedRat.Image = Properties.Resources.OFF;
        }

        private void CameraConnect()        //TO DO: Inset your connection code here
        {
            if (ini12.INIRead(MainSettingPath, "Device", "Name", "") != "")
            {
                ini12.INIWrite(MainSettingPath, "Device", "CameraExist", "1");
                pictureBox_Camera.Image = Properties.Resources.ON;
                if (StartButtonPressed == false)
                    button_Camera.Enabled = true;
            }
        }

        private void CameraDisconnect()     //TO DO: Insert your disconnection code here
        {
            ini12.INIWrite(MainSettingPath, "Device", "CameraExist", "0");
            pictureBox_Camera.Image = Properties.Resources.OFF;
            if (StartButtonPressed == false)
                button_Camera.Enabled = false;
        }

        protected override void WndProc(ref Message m)
        {
            USBPort.ProcessWindowsMessage(ref m);
            base.WndProc(ref m);
        }
        #endregion

        private void OnCaptureComplete(object sender, EventArgs e)
        {
            // Demonstrate the Capture.CaptureComplete event.
            Debug.WriteLine("Capture complete.");
        }

        //執行緒控制label.text
        private delegate void UpdateUICallBack(string value, Control ctl);
        private void UpdateUI(string value, Control ctl)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack uu = new UpdateUICallBack(UpdateUI);
                Invoke(uu, value, ctl);
            }
            else
            {
                ctl.Text = value;
            }
        }

        //執行緒控制 datagriveiew
        private delegate void UpdateUICallBack1(string value, DataGridView ctl);
        private void GridUI(string i, DataGridView gv)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack1 uu = new UpdateUICallBack1(GridUI);
                Invoke(uu, i, gv);
            }
            else
            {
                DataGridView_Schedule.ClearSelection();
                gv.Rows[int.Parse(i)].Selected = true;
            }
        }

        // 執行緒控制 datagriverew的scorllingbar
        private delegate void UpdateUICallBack3(string value, DataGridView ctl);
        private void Gridscroll(string i, DataGridView gv)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack3 uu = new UpdateUICallBack3(Gridscroll);
                Invoke(uu, i, gv);
            }
            else
            {
                //DataGridView1.ClearSelection();
                //gv.Rows[int.Parse(i)].Selected = true;
                gv.FirstDisplayedScrollingRowIndex = int.Parse(i);
            }
        }

        //執行緒控制 txtbox1
        private delegate void UpdateUICallBack2(string value, Control ctl);
        private void Txtbox1(string value, Control ctl)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack2 uu = new UpdateUICallBack2(Txtbox1);
                Invoke(uu, value, ctl);
            }
            else
            {
                ctl.Text = value;
            }
        }

        //執行緒控制 txtbox2
        private delegate void UpdateUICallBack4(string value, Control ctl);
        private void Txtbox2(string value, Control ctl)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack4 uu = new UpdateUICallBack4(Txtbox2);
                Invoke(uu, value, ctl);
            }
            else
            {
                ctl.Text = value;
            }
        }

        //執行緒控制 txtbox3
        private delegate void UpdateUICallBack5(string value, Control ctl);
        private void Txtbox3(string value, Control ctl)
        {
            if (InvokeRequired)
            {
                UpdateUICallBack5 uu = new UpdateUICallBack5(Txtbox3);
                Invoke(uu, value, ctl);
            }
            else
            {
                ctl.Text = value;
            }
        }

        protected void OpenRedRat3()
        {
            int dev = 0;
            string intdev = ini12.INIRead(MainSettingPath, "RedRat", "RedRatIndex", "");

            if (intdev != "-1")
                dev = int.Parse(intdev);

            var devices = RedRat3USBImpl.FindDevices();

            // 假若設定值大於目前device個數，直接更改為目前device個數
            if (dev >= devices.Count)
                dev = devices.Count - 1;

            if (devices.Count > 0)
            {
                //RedRat已連線
                redRat3 = (IRedRat3)devices[dev].GetRedRat();

                //pictureBox1綠燈
                pictureBox_RedRat.Image = Properties.Resources.ON;
            }
            else
                pictureBox_RedRat.Image = Properties.Resources.OFF;
        }

        private void ConnectAutoBox1()
        {   // RS232 Setting
            serialPortWood.StopBits = System.IO.Ports.StopBits.One;
            serialPortWood.PortName = ini12.INIRead(MainSettingPath, "Device", "AutoboxPort", "");
            //serialPort3.BaudRate = int.Parse(ini12.INIRead(sPath, "SerialPort", "Baudrate", ""));
            if (serialPortWood.IsOpen == false)
            {
                serialPortWood.Open();
                object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(serialPortWood);
                hCOM = (SafeFileHandle)stream.GetType().GetField("_handle", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(stream);
            }
            else
            {
                Console.WriteLine(DateTime.Now.ToString("h:mm:ss tt") + " - Cannot connect to AutoBox.\n");
            }
        }

        private void ConnectAutoBox2()
        {
            uint temp_version;
            string curItem = ini12.INIRead(MainSettingPath, "Device", "AutoboxPort", "");
            if (MyBlueRat.Connect(curItem) == true)
            {
                temp_version = MyBlueRat.FW_VER;
                float v = temp_version;
                label_BoxVersion.Text = "_" + (v / 100).ToString("0.00");

                // 在第一次/或長時間未使用之後,要開始使用BlueRat跑Schedule之前,建議執行這一行,確保BlueRat的起始狀態一致 -- 正常情況下不執行並不影響BlueRat運行,但為了找問題方便,還是請務必執行
                MyBlueRat.Force_Init_BlueRat();
                MyBlueRat.Reset_SX1509();

                byte SX1509_detect_status;
                SX1509_detect_status = MyBlueRat.TEST_Detect_SX1509();

                if (SX1509_detect_status == 3)
                {
                    pictureBox_ext_board.Image = Properties.Resources.ON;
                    // Error, need to check SX1509 connection
                }
                else
                {
                    pictureBox_ext_board.Image = Properties.Resources.OFF;
                }

                hCOM = MyBlueRat.ReturnSafeFileHandle();
                BlueRat_UART_Exception_status = false;
                UpdateRCFunctionButtonAfterConnection();
            }
            else
            {
                Console.WriteLine(DateTime.Now.ToString("h:mm:ss tt") + " - Cannot connect to BlueRat.\n");
            }
        }

        private void ConnectArduino()
        {
            OpenSerialPort("Arduino");
        }

        private void ConnectFtdi()
        {
            portinfo.I2C_Channel_Conf.ClockRate = Ft_I2C_ClockRate.I2C_CLOCK_STANDARD_MODE;
            portinfo.I2C_Channel_Conf.LatencyTimer = 200;
            portinfo.I2C_Channel_Conf.Options = 0x00000001;     //FtConfigOptions.I2C_DISABLE_3PHASE_CLOCKING;
            portinfo.PortNum = 0;
            portinfo.ftStatus = Ftdi_lib.I2C_Init(out portinfo.ftHandle, portinfo);
            if (portinfo.ftStatus == FtResult.Ok)
                pictureBox_ftdi.Image = Properties.Resources.ON;
            else
                pictureBox_ftdi.Image = Properties.Resources.OFF;
        }

        private void DisconnectAutoBox1()
        {
            serialPortWood.Close();
        }

        private void DisconnectAutoBox2()
        {
            if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
            {
                if (MyBlueRat.Disconnect() == true)
                {
                    if (BlueRat_UART_Exception_status)
                    {
                        //Serial_UpdatePortName(); 
                    }
                    BlueRat_UART_Exception_status = false;
                }
                else
                {
                    Console.WriteLine(DateTime.Now.ToString("h:mm:ss tt") + " - Cannot disconnect from RS232.\n");
                }
            }
        }

        private void DisconnectArduino()
        {
            if (serialPort_Arduino.IsOpen == true)
            {
                CloseSerialPort("Arduino");
            }
        }

        private void DisconnectFtdi()
        {
            portinfo.ftStatus = Ftdi_lib.I2C_DeInit(portinfo.ftHandle);
        }

        protected void ConnectUsbCAN()
        {
            uint status;

            status = Can_Usb2C.Connect();
            if (status == 1)
            {
                status = Can_Usb2C.StartCAN();
                if (status == 1)
                {
                    timer_canbus.Enabled = true;
                    pictureBox_can.Image = Properties.Resources.ON;
                }
                else
                {
                    pictureBox_can.Image = Properties.Resources.OFF;
                }
            }
            else
            {
                pictureBox_can.Image = Properties.Resources.OFF;
            }
        }

        protected void ConnectVectorCAN()
        {
            uint status;

            status = Can_1630A.Connect();
            if (status == 1)
            {
                status = Can_1630A.StartCAN();
                if (status == 1)
                {
                    timer_canbus.Enabled = true;
                    pictureBox_can.Image = Properties.Resources.ON;
                }
                else
                {
                    pictureBox_can.Image = Properties.Resources.OFF;
                }
            }
            else
            {
                pictureBox_can.Image = Properties.Resources.OFF;
            }
        }

        public void Autocommand_RedRat(string Caller, string SigData)
        {
            string redcon = "";

            //讀取設備//
            if (Caller == "Form1")
            {
                RedRatData.RedRatLoadSignalDB(ini12.INIRead(MainSettingPath, "RedRat", "DBFile", ""));
                redcon = ini12.INIRead(MainSettingPath, "RedRat", "Brands", "");
            }
            else if (Caller == "FormRc")
            {
                string SelectRcLastTimePath = ini12.INIRead(RcPath, "Setting", "SelectRcLastTimePath", "");
                RedRatData.RedRatLoadSignalDB(ini12.INIRead(SelectRcLastTimePath, "Info", "DBFile", ""));
                redcon = ini12.INIRead(SelectRcLastTimePath, "Info", "Brands", "");
            }

            try
            {
                if (RedRatData.SignalDB.GetIRPacket(redcon, SigData).ToString() == "RedRat.IR.DoubleSignal")
                {
                    DoubleSignal sig = (DoubleSignal)RedRatData.SignalDB.GetIRPacket(redcon, SigData);
                    if (redRat3 != null)
                        redRat3.OutputModulatedSignal(sig);
                }
                else
                {
                    ModulatedSignal sig2 = (ModulatedSignal)RedRatData.SignalDB.GetIRPacket(redcon, SigData);
                    if (redRat3 != null)
                        redRat3.OutputModulatedSignal(sig2);
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex);
                MessageBox.Show(Ex.Message.ToString(), "Transmit RC signal fail!");
            }
        }

        private Boolean D = false;
        public void Autocommand_BlueRat(string Caller, string SigData)
        {
            try
            {
                if (Caller == "Form1")
                {
                    RedRatData.RedRatLoadSignalDB(ini12.INIRead(MainSettingPath, "RedRat", "DBFile", ""));
                    RedRatData.RedRatSelectDevice(ini12.INIRead(MainSettingPath, "RedRat", "Brands", ""));
                }
                else if (Caller == "FormRc")
                {
                    string SelectRcLastTimePath = ini12.INIRead(RcPath, "Setting", "SelectRcLastTimePath", "");
                    RedRatData.RedRatLoadSignalDB(ini12.INIRead(SelectRcLastTimePath, "Info", "DBFile", ""));
                    RedRatData.RedRatSelectDevice(ini12.INIRead(SelectRcLastTimePath, "Info", "Brands", ""));
                }

                RedRatData.RedRatSelectRCSignal(SigData, D);

                if (RedRatData.Signal_Type_Supported != true)
                {
                    return;
                }

                // Use UART to transmit RC signal
                int rc_duration = MyBlueRat.SendOneRC(RedRatData) / 1000 + 1;
                RedRatDBViewer_Delay(rc_duration);
                /*
                int SysDelay = int.Parse(columns_wait);
                if (SysDelay <= rc_duration)
                {
                    RedRatDBViewer_Delay(rc_duration);
                }
                */
                if ((RedRatData.RedRatSelectedSignalType() == (typeof(DoubleSignal))) || (RedRatData.RC_ToggleData_Length_Value() > 0))
                {
                    RedRatData.RedRatSelectRCSignal(SigData, D);
                    D = !D;
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex);
                MessageBox.Show(Ex.Message.ToString(), "Transmit RC signal fail!");
            }
        }

        private void UpdateRCFunctionButtonAfterConnection()
        {
            if ((MyBlueRat.CheckConnection() == true))
            {
                if ((RedRatData != null) && (RedRatData.SignalDB != null) && (RedRatData.SelectedDevice != null) && (RedRatData.SelectedSignal != null))
                {
                    button_Start.Enabled = true;
                }
                else
                {
                    button_Start.Enabled = false;
                }
            }
        }
        /*
                static async System.Threading.Tasks.Task Delay(int iSecond)
                {
                    await System.Threading.Tasks.Task.Delay(iSecond);
                }

                async Task RedRatDBViewer_Delay(int delay_ms)
                {
                    try
                    {
                        await Delay(delay_ms);
                        //System.Threading.Thread.Sleep(delay_ms);
                    }
                    catch (TaskCanceledException ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
        */
        
        // Woodpecker debug function 
        private void debug_process(string log)
        {
            try
            {
                debug_text = string.Concat(debug_text, "[Debug] [" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + log + "\r\n");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                Serialportsave("Debug");
            }
        }

        // Log record function
        /*  moved to Log_Dump.cs
        private void log_process(string port, string log)
        */

        // 這個主程式專用的delay的內部資料與function
        static bool RedRatDBViewer_Delay_TimeOutIndicator = false;
        private void RedRatDBViewer_Delay_OnTimedEvent(object source, ElapsedEventArgs e)
        {
            //log.Debug("RedRatDBViewer_Delay_TimeOutIndicator_True");
            RedRatDBViewer_Delay_TimeOutIndicator = true;
        }

        private void RedRatDBViewer_Delay(int delay_ms)
        {
            log.Debug("RedRatDBViewer_Delay_S");
            if (delay_ms <= 0) return;
            System.Timers.Timer aTimer = new System.Timers.Timer(delay_ms);
            //aTimer.Interval = delay_ms;
            aTimer.Elapsed += new ElapsedEventHandler(RedRatDBViewer_Delay_OnTimedEvent);
            aTimer.SynchronizingObject = this.TimeLabel2;
            RedRatDBViewer_Delay_TimeOutIndicator = false;
            aTimer.Enabled = true;
            aTimer.Start();
            while ((FormIsClosing == false) && (RedRatDBViewer_Delay_TimeOutIndicator == false))
            {
                if (temperatureDouble.Count() > 0 || timer_matched)
                {
                    if (temperatureDouble.Count() > 0)
                    {
                        currentTemperature = temperatureDouble.Dequeue();
                        label_Command.Text = "Condition: " + currentTemperature + ", SHOT: " + currentTemperature;
                    }
                    else if (timer_matched)
                    {
                        label_Command.Text = "Timer: matched.";
                        timer_matched = false;
                    }

                    GlobalData.caption_Num++;
                    if (GlobalData.Loop_Number == 1)
                        GlobalData.caption_Sum = GlobalData.caption_Num;
                    Jes();
                }

                if (logA_text != null && logA_text.Length > log_max_length)
                //if (logA_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        //Serialportsave("A");
                        logDumpping.LogDumpToFile(serialPortConfig_A, GlobalData.portConfigGroup_A.portName, ref logA_text);
                    }
                    else
                    {
                        logA_text = string.Empty;
                    }
                }

                if (logB_text != null && logB_text.Length > log_max_length)
                //if (logB_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        //Serialportsave("B");
                        logDumpping.LogDumpToFile(serialPortConfig_B, GlobalData.portConfigGroup_B.portName, ref logB_text);
                    }
                    else
                    {
                        logB_text = string.Empty;
                    }
                }

                if (logC_text != null && logC_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        logDumpping.LogDumpToFile(serialPortConfig_C, GlobalData.portConfigGroup_C.portName, ref logC_text);
                    }
                    else
                    {
                        logC_text = string.Empty;
                    }
                }

                if (logD_text != null && logD_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        logDumpping.LogDumpToFile(serialPortConfig_D, GlobalData.portConfigGroup_D.portName, ref logD_text);
                    }
                    else
                    {
                        logD_text = string.Empty;
                    }
                }

                if (logE_text != null && logE_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        logDumpping.LogDumpToFile(serialPortConfig_E, GlobalData.portConfigGroup_E.portName, ref logE_text);
                    }
                    else
                    {
                        logE_text = string.Empty;
                    }
                }

                if (logAll_text != null && logAll_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        Serialportsave("All");
                    }
                    else
                    {
                        logAll_text = string.Empty;
                    }
                }

                if (canbus_text != null && canbus_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        Serialportsave("Canbus");
                    }
                    else
                    {
                        canbus_text = string.Empty;
                    }
                }

                if (kline_text != null && kline_text.Length > log_max_length)
                {
                    if (ini12.INIRead(MainSettingPath, "Autosavelog", "Checked", "") == "1")
                    {
                        Serialportsave("KlinePort");
                    }
                    else
                    {
                        kline_text = string.Empty;
                    }
                }

                //log.Debug("RedRatDBViewer_Delay_TimeOutIndicator_false");
                Application.DoEvents();
                System.Threading.Thread.Sleep(1);//釋放CPU//

                if (GlobalData.Break_Out_MyRunCamd == 1)//強制讓schedule直接停止//
                {
                    GlobalData.Break_Out_MyRunCamd = 0;
                    log.Debug("Break_Out_MyRunCamd_0");
                    break;
                }
            }

            aTimer.Stop();
            aTimer.Dispose();
            log.Debug("RedRatDBViewer_Delay_E");
        }

        // 這個usbcan專用的delay的內部資料與function
        static bool UsbCAN_Delay_TimeOutIndicator = false;
        static UInt64 UsbCAN_Count = 0;
        private void UsbCAN_Delay_UsbOnTimedEvent(object source, ElapsedEventArgs e)
        {
            uint columns_times = can_id;
            byte[] columns_serial = can_data[columns_times];
            int columns_interval = (int)can_rate[columns_times];
            Can_Usb2C.TransmitData(columns_times, columns_serial);
            Console.WriteLine("USB_Can_Send (Repeat): " + UsbCAN_Count + " times.");

            string Outputstring = "ID: 0x";
            //Outputstring += columns_times + " Data: " + columns_serial;
            DateTime dt = DateTime.Now;
            string canbus_log_text = "[Send_UsbCAN] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
            logDumpping.LogCat(ref canbus_text, canbus_log_text);        //Replaced by another Debugging function//log_process("Canbus", canbus_log_text);
            UsbCAN_Count++;
            UsbCAN_Delay_TimeOutIndicator = true;
        }

        private void UsbCAN_Delay(int delay_ms)
        {
            //Console.WriteLine("UsbCAN_Delay: Start.");
            if (delay_ms <= 0) return;
            System.Timers.Timer UsbCAN_Timer = new System.Timers.Timer(delay_ms);
            UsbCAN_Timer.Interval = delay_ms;
            UsbCAN_Timer.Elapsed += new ElapsedEventHandler(UsbCAN_Delay_UsbOnTimedEvent);
            UsbCAN_Timer.Enabled = true;
            UsbCAN_Timer.Start();
            UsbCAN_Timer.AutoReset = true;

            while ((FormIsClosing == false) && (UsbCAN_Delay_TimeOutIndicator == false))
            {
                //Console.WriteLine("UsbCAN_Delay_TimeOutIndicator: false.");
                Application.DoEvents();
                System.Threading.Thread.Sleep(1);//釋放CPU//

                if (GlobalData.Break_Out_MyRunCamd == 1)//強制讓schedule直接停止//
                {
                    GlobalData.Break_Out_MyRunCamd = 0;
                    //Console.WriteLine("Break_Out_MyRunCamd = 0");
                    break;
                }
            }
            UsbCAN_Timer.Stop();
            UsbCAN_Timer.Dispose();
            //Console.WriteLine("UsbCAN_Delay: Stop.");
        }

        // 這個vectorcan專用的delay的內部資料與function
        static bool VectorCAN_Delay_TimeOutIndicator = false;
        static UInt64 VectorCAN_Count = 0;
        private void VectorCAN_Delay_UsbOnTimedEvent(object source, ElapsedEventArgs e)
        {
            uint columns_times = can_id;
            byte[] columns_serial = can_data[columns_times];
            int columns_interval = (int)can_rate[columns_times];
            Can_1630A.LoopCANTransmit(columns_times, (uint)columns_interval, columns_serial);
            Console.WriteLine("VectorCAN_Can_Send (Repeat): " + VectorCAN_Count + " times.");

            string Outputstring = "ID: 0x";
            //Outputstring += columns_times + " Data: " + columns_serial;
            DateTime dt = DateTime.Now;
            string canbus_log_text = "[Send_VectorCAN] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
            logDumpping.LogCat(ref canbus_text, canbus_log_text);        //Replaced by another Debugging function//log_process("Canbus", canbus_log_text);
            VectorCAN_Count++;
            VectorCAN_Delay_TimeOutIndicator = true;
        }

        private void VectorCAN_Delay(int delay_ms)
        {
            //Console.WriteLine("VectorCAN_Delay: Start.");
            if (delay_ms <= 0) return;
            System.Timers.Timer VectorCAN_Timer = new System.Timers.Timer(delay_ms);
            VectorCAN_Timer.Interval = delay_ms;
            VectorCAN_Timer.Elapsed += new ElapsedEventHandler(VectorCAN_Delay_UsbOnTimedEvent);
            VectorCAN_Timer.Enabled = true;
            VectorCAN_Timer.Start();
            VectorCAN_Timer.AutoReset = true;

            while ((FormIsClosing == false) && (VectorCAN_Delay_TimeOutIndicator == false))
            {
                //Console.WriteLine("VectorCAN_Delay_TimeOutIndicator: false.");
                Application.DoEvents();
                System.Threading.Thread.Sleep(1);//釋放CPU//

                if (GlobalData.Break_Out_MyRunCamd == 1)//強制讓schedule直接停止//
                {
                    GlobalData.Break_Out_MyRunCamd = 0;
                    //Console.WriteLine("Break_Out_MyRunCamd = 0");
                    break;
                }
            }
            VectorCAN_Timer.Stop();
            VectorCAN_Timer.Dispose();
            //Console.WriteLine("VectorCAN_Delay: Stop.");
        }

        private void Log(string msg)
        {
            textBox_serial.Invoke(new EventHandler(delegate
            {
                textBox_serial.Text = msg.Trim();
                PortA.WriteLine(msg.Trim());
            }));
        }

        public static string ByteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }

        #region -- SerialPort Setup --
        protected void OpenSerialPort(string Port)
        {
            switch (Port)
            {
                case "A":
                    try
                    {
                        if (PortA.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port A", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    PortA.StopBits = System.IO.Ports.StopBits.One;
                                    break;
                                case "Two":
                                    PortA.StopBits = System.IO.Ports.StopBits.Two;
                                    break;
                            }
                            PortA.PortName = ini12.INIRead(MainSettingPath, "Port A", "PortName", "");
                            PortA.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port A", "BaudRate", ""));
                            PortA.DataBits = int.Parse(ini12.INIRead(MainSettingPath, "Port A", "DataBit", ""));
                            PortA.ReadTimeout = 2000;
                            // serialPort2.Encoding = System.Text.Encoding.GetEncoding(1252);

//                          PortA.DataReceived += new SerialDataReceivedEventHandler(SerialPort1_DataReceived);       // DataReceived呼叫函式
                            PortA.Open();
                            object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortA);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "PortA Error");
                    }
                    break;
                case "B":
                    try
                    {
                        if (PortB.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port B", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    PortB.StopBits = System.IO.Ports.StopBits.One;
                                    break;
                                case "Two":
                                    PortB.StopBits = System.IO.Ports.StopBits.Two;
                                    break;
                            }
                            PortB.PortName = ini12.INIRead(MainSettingPath, "Port B", "PortName", "");
                            PortB.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port B", "BaudRate", ""));
                            PortB.DataBits = int.Parse(ini12.INIRead(MainSettingPath, "Port B", "DataBit", ""));
                            PortB.ReadTimeout = 2000;
                            // serialPort2.Encoding = System.Text.Encoding.GetEncoding(1252);

                            // PortB.DataReceived += new SerialDataReceivedEventHandler(SerialPort2_DataReceived);       // DataReceived呼叫函式
                            PortB.Open();
                            object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortB);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "PortB Error");
                    }
                    break;
                case "C":
                    try
                    {
                        if (PortC.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port C", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    PortC.StopBits = System.IO.Ports.StopBits.One;
                                    break;
                                case "Two":
                                    PortC.StopBits = System.IO.Ports.StopBits.Two;
                                    break;
                            }
                            PortC.PortName = ini12.INIRead(MainSettingPath, "Port C", "PortName", "");
                            PortC.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port C", "BaudRate", ""));
                            PortC.DataBits = int.Parse(ini12.INIRead(MainSettingPath, "Port C", "DataBit", ""));
                            PortC.ReadTimeout = 2000;
                            // serialPort3.Encoding = System.Text.Encoding.GetEncoding(1252);

                            // PortC.DataReceived += new SerialDataReceivedEventHandler(SerialPort3_DataReceived);       // DataReceived呼叫函式
                            PortC.Open();
                            object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortC);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "PortC Error");
                    }
                    break;
                case "D":
                    try
                    {
                        if (PortD.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port D", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    PortD.StopBits = System.IO.Ports.StopBits.One;
                                    break;
                                case "Two":
                                    PortD.StopBits = System.IO.Ports.StopBits.Two;
                                    break;
                            }
                            PortD.PortName = ini12.INIRead(MainSettingPath, "Port D", "PortName", "");
                            PortD.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port D", "BaudRate", ""));
                            PortD.DataBits = int.Parse(ini12.INIRead(MainSettingPath, "Port D", "DataBit", ""));
                            PortD.ReadTimeout = 2000;
                            // serialPort3.Encoding = System.Text.Encoding.GetEncoding(1252);

                            // PortD.DataReceived += new SerialDataReceivedEventHandler(SerialPort4_DataReceived);       // DataReceived呼叫函式
                            PortD.Open();
                            object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortD);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "PortD Error");
                    }
                    break;
                case "E":
                    try
                    {
                        if (PortE.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port E", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    PortE.StopBits = System.IO.Ports.StopBits.One;
                                    break;
                                case "Two":
                                    PortE.StopBits = System.IO.Ports.StopBits.Two;
                                    break;
                            }
                            PortE.PortName = ini12.INIRead(MainSettingPath, "Port E", "PortName", "");
                            PortE.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port E", "BaudRate", ""));
                            PortE.DataBits = int.Parse(ini12.INIRead(MainSettingPath, "Port E", "DataBit", ""));
                            PortE.ReadTimeout = 2000;
                            // serialPort3.Encoding = System.Text.Encoding.GetEncoding(1252);

                            // PortE.DataReceived += new SerialDataReceivedEventHandler(SerialPort5_DataReceived);       // DataReceived呼叫函式
                            PortE.Open();
                            object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortE);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "PortE Error");
                    }
                    break;
                case "Arduino":
                    try
                    {
                        serialPort_Arduino.StopBits = System.IO.Ports.StopBits.One;
                        serialPort_Arduino.PortName = ini12.INIRead(MainSettingPath, "Device", "ArduinoPort", "");
                        serialPort_Arduino.BaudRate = 9600;
                        serialPort_Arduino.DataBits = 8;
                        serialPort_Arduino.ReadTimeout = 2000;
                        serialPort_Arduino.Encoding = System.Text.Encoding.GetEncoding(1252);

                        serialPort_Arduino.DataReceived += new SerialDataReceivedEventHandler(serialport_arduino_datareceived);       // DataReceived呼叫函式
                        serialPort_Arduino.Open();
                        object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(serialPort_Arduino);
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "ArduinoPort Error");
                    }
                    break;
                case "kline":
                    try
                    {
                        string Kline_Exist = ini12.INIRead(MainSettingPath, "Kline", "Checked", "");

                        if (Kline_Exist == "1" && MySerialPort.IsPortOpened() == false)
                        {
                            string curItem = ini12.INIRead(MainSettingPath, "Kline", "PortName", "");
                            if (MySerialPort.OpenPort(curItem) == true)
                            {
                                //BlueRat_UART_Exception_status = false;
                                timer_kline.Enabled = true;
                            }
                            else
                            {
                                timer_kline.Enabled = false;
                            }
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "KlinePort Error");
                    }
                    break;
                default:
                    break;
            }
        }

        protected void CloseSerialPort(string Port)
        {
            switch (Port)
            {
                case "A":
                    PortA.Dispose();
                    PortA.Close();
                    break;
                case "B":
                    PortB.Dispose();
                    PortB.Close();
                    break;
                case "C":
                    PortC.Dispose();
                    PortC.Close();
                    break;
                case "D":
                    PortD.Dispose();
                    PortD.Close();
                    break;
                case "E":
                    PortE.Dispose();
                    PortE.Close();
                    break;
                case "Arduino":
                    serialPort_Arduino.Dispose();
                    serialPort_Arduino.Close();
                    break;
                case "kline":
                    MySerialPort.Dispose();
                    MySerialPort.ClosePort();
                    break;
                default:
                    break;
            }
        }
        #endregion

        #region -- Old SerialPort Setup --
        protected void OpenSerialPort1()
        {
            try
            {
                if (PortA.IsOpen == false)
                {
                    string stopbit = ini12.INIRead(MainSettingPath, "Port A", "StopBits", "");
                    switch (stopbit)
                    {
                        case "One":
                            PortA.StopBits = System.IO.Ports.StopBits.One;
                            break;
                        case "Two":
                            PortA.StopBits = System.IO.Ports.StopBits.Two;
                            break;
                    }
                    PortA.PortName = ini12.INIRead(MainSettingPath, "Port A", "PortName", "");
                    PortA.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port A", "BaudRate", ""));
                    PortA.ReadTimeout = 2000;
                    // serialPort2.Encoding = System.Text.Encoding.GetEncoding(1252);

//                    PortA.DataReceived += new SerialDataReceivedEventHandler(SerialPort1_DataReceived);       // DataReceived呼叫函式
                    PortA.Open();
                    object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(PortA);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message.ToString(), "SerialPort1 Error");
            }
        }
        /*
                protected PortDataContainer OpenSerialPort1(SerialPort sp)
                {
                    PortDataContainer sp_data = new PortDataContainer();
                    sp_data.serial_port = sp;
                    PortDataContainer.PortDictionary.Add(sp.PortName, sp_data);
                    try
                    {
                        if (serialPort1.IsOpen == false)
                        {
                            string stopbit = ini12.INIRead(MainSettingPath, "Port A", "StopBits", "");
                            switch (stopbit)
                            {
                                case "One":
                                    serialPort1.StopBits = StopBits.One;
                                    break;
                                case "Two":
                                    serialPort1.StopBits = StopBits.Two;
                                    break;
                            }
                            serialPort1.PortName = ini12.INIRead(MainSettingPath, "Port A", "PortName", "");
                            serialPort1.BaudRate = int.Parse(ini12.INIRead(MainSettingPath, "Port A", "BaudRate", ""));
                            serialPort1.DataBits = 8;
                            serialPort1.Parity = (Parity)0;
                            serialPort1.ReceivedBytesThreshold = 1;
                            serialPort1.ReadTimeout = 2000;
                            // serialPort1.Encoding = System.Text.Encoding.GetEncoding(1252);

                            serialPort1.DataReceived += new SerialDataReceivedEventHandler(SerialPort1_DataReceived);       // DataReceived呼叫函式
                            serialPort1.Open();
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message.ToString(), "SerialPort1 Error");
                    }
                    return sp_data;
                }
        */
        protected void CloseSerialPort1()
        {
            PortA.Dispose();
            PortA.Close();
        }
        #endregion

        #region -- 接受SerialPort資料 --

        private void logA_analysis()
        {
            logDumpping.LogDataReceiving(GlobalData.m_SerialPort_A, GlobalData.portConfigGroup_A.portConfig, GlobalData.portConfigGroup_A.portLF, ref logA_text);
        }
        private void logB_analysis()
        {
            logDumpping.LogDataReceiving(GlobalData.m_SerialPort_B, GlobalData.portConfigGroup_B.portConfig, GlobalData.portConfigGroup_B.portLF, ref logB_text);
        }
        private void logC_analysis()
        {
            logDumpping.LogDataReceiving(GlobalData.m_SerialPort_C, GlobalData.portConfigGroup_C.portConfig, GlobalData.portConfigGroup_C.portLF, ref logC_text);
        }
        private void logD_analysis()
        {
            logDumpping.LogDataReceiving(GlobalData.m_SerialPort_D, GlobalData.portConfigGroup_D.portConfig, GlobalData.portConfigGroup_D.portLF, ref logD_text);
        }
        private void logE_analysis()
        {
            logDumpping.LogDataReceiving(GlobalData.m_SerialPort_E, GlobalData.portConfigGroup_E.portConfig, GlobalData.portConfigGroup_E.portLF, ref logE_text);
        }

        private void logA_RK2797()
        {
            while (GlobalData.m_SerialPort_A.IsOpen() == true)
            {
                rk2797.Package_add_queue(GlobalData.m_SerialPort_A);
                rk2797.Package_queue_to_list(GlobalData.m_SerialPort_A);
                rk2797.Package_queue_to_catch(GlobalData.m_SerialPort_A);
            }
        }
        private void logB_RK2797()
        {
            while (GlobalData.m_SerialPort_B.IsOpen() == true)
            {
                rk2797.Package_add_queue(GlobalData.m_SerialPort_B);
                rk2797.Package_queue_to_list(GlobalData.m_SerialPort_B);
                rk2797.Package_queue_to_catch(GlobalData.m_SerialPort_B);
            }
        }
        private void logC_RK2797()
        {
            while (GlobalData.m_SerialPort_C.IsOpen() == true)
            {
                rk2797.Package_add_queue(GlobalData.m_SerialPort_C);
                rk2797.Package_queue_to_list(GlobalData.m_SerialPort_C);
                rk2797.Package_queue_to_catch(GlobalData.m_SerialPort_C);
            }
        }
        private void logD_RK2797()
        {
            while (GlobalData.m_SerialPort_D.IsOpen() == true)
            {
                rk2797.Package_add_queue(GlobalData.m_SerialPort_D);
                rk2797.Package_queue_to_list(GlobalData.m_SerialPort_D);
                rk2797.Package_queue_to_catch(GlobalData.m_SerialPort_D);
            }
        }
        private void logE_RK2797()
        {
            while (GlobalData.m_SerialPort_E.IsOpen() == true)
            {
                rk2797.Package_add_queue(GlobalData.m_SerialPort_E);
                rk2797.Package_queue_to_list(GlobalData.m_SerialPort_E);
                rk2797.Package_queue_to_catch(GlobalData.m_SerialPort_E);
            }
        }
        #endregion

        #region -- 接受SerialPort1資料 --
        /*
                public class SerialReceivedData
                {
                    private List<Byte> data;
                    private DateTime time_stamp;
                    public void SetData(List<Byte> d) { data = d; }
                    public void SetTimeStamp(DateTime t) { time_stamp = t; }
                    public List<Byte> GetData() { return data; }
                    public DateTime GetTimeStamp() { return time_stamp; }
                }

                public class PortDataContainer
                {
                    static public Dictionary<string, Object> PortDictionary;
                    static public bool data_available;
                    public SerialPort serial_port;
                    public Queue<SerialReceivedData> data_queue;
                    //public List<SerialReceivedData> received_data = new List<SerialReceivedData>(); // just-received and to be processed
                    public Queue<Byte> log_data; // processed and stored for log_save
                    public PortDataContainer()
                    {
                        PortDictionary = new Dictionary<string, Object>();
                        data_queue = new Queue<SerialReceivedData>();
                        log_data = new Queue<Byte>();
                        data_available = false;
                    }
                }
                byte[] dataset1 = new byte[0];
                byte[] dataset2 = new byte[0];
                byte[] dataset3 = new byte[0];
                public void AddDataMethod1(String myString)
                {
                    DateTime dt;
                    if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
                    {
                        // hex to string
                        string hexValues = BitConverter.ToString(dataset1).Replace("-", "");
                        dt = DateTime.Now;

                        // Joseph
                        hexValues = hexValues.Replace(Environment.NewLine, "\r\n" + "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  "); //OK
                                                                                                                                       // hexValues = String.Concat("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n");
                        textBox1.AppendText(hexValues);
                        // End

                        // Jeremy
                        // textBox1.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
                        // textBox1.AppendText(hexValues + "\r\n");
                        // End
                    }
                    else
                    {
                        // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
                        string text = Encoding.ASCII.GetString(dataset1);

                        dt = DateTime.Now;
                        text = text.Replace(Environment.NewLine, "\r\n" + "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  "); //OK
                        textBox1.AppendText(text);
                    }
                    Thread.Sleep(1);
                }

                public PortDataContainer PortA = new PortDataContainer();
        */

        //private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        //{
        //    try
        //    {
        //        int data_to_read = PortA.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            dataset = new byte[data_to_read];

        //            PortA.Read(dataset, 0, data_to_read);
        //            int index = 0;
        //            while (data_to_read > 0)
        //            {
        //                byte data_byte = dataset[index];
        //                LogQueue_A.Enqueue(data_byte);
        //                SearchLogQueue_A.Enqueue(data_byte);
        //                if (TemperatureIsFound == true)
        //                    TemperatureQueue_A.Enqueue(data_byte);
        //                index++;
        //                data_to_read--;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //}

        //private void logA_analysis()
        //{
        //    while (PortA.IsOpen == true)
        //    {
        //        int data_to_read = PortA.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            byte[] dataset = new byte[data_to_read];
        //            PortA.Read(dataset, 0, data_to_read);

        //            for (int index = 0; index < data_to_read; index++)
        //            {
        //                byte input_ch = dataset[index];
        //                logA_recorder(input_ch);
        //                if (TemperatureIsFound == true)
        //                {
        //                    log_temperature(input_ch);
        //                }
        //            }
        //            //else
        //            //{
        //            //    logA_recorder(0x00,true); // tell log_recorder no more data for now.
        //            //}
        //        }
        //        //else
        //        //{
        //        //    logA_recorder(0x00,true); // tell log_recorder no more data for now.
        //        //}
        //    }
        //}

        const int byteMessage_max_Hex = 16;
        const int byteMessage_max_Ascii = 256;
        byte[] byteMessage_A = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length_A = 0;

        private void logA_recorder(byte ch, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage_A[byteMessage_length_A] = ch;
                    byteMessage_length_A++;
                }
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_A >= byteMessage_max_Hex) /*|| (SaveToLog == true)*/)
                {
                    string dataValue = BitConverter.ToString(byteMessage_A).Replace("-", "").Substring(0, byteMessage_length_A * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logA_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_A = 0;
                }
            }
            else
            {
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_A >= byteMessage_max_Ascii))
                {
                    string dataValue = Encoding.ASCII.GetString(byteMessage_A).Substring(0, byteMessage_length_A);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logA_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_A = 0;
                }
                else
                {
                    byteMessage_A[byteMessage_length_A] = ch;
                    byteMessage_length_A++;
                }
            }
        }

        //private void logA_recorder()
        //{
        //    DateTime dt;
        //    byte myByteList;

        //    if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
        //    {
        //        while (LogQueue_A.Count > 0)
        //        {
        //            myByteList = LogQueue_A.Dequeue();
        //            byteMessage_A[byteMessage_length_A] = myByteList;
        //            byteMessage_length_A++;
        //            if ((myByteList == 0x0A) || (myByteList == 0x0D) || (byteMessage_length_A >= byteMessage_max_A))
        //            {
        //                string dataValue = BitConverter.ToString(byteMessage_A).Replace("-", "").Substring(0, byteMessage_length_A * 2);
        //                dt = DateTime.Now;
        //                dataValue = "[Receive_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
        //                log_process("A", dataValue);
        //                log_process("All", dataValue);
        //                byteMessage_length_A = 0;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        while (LogQueue_A.Count > 0)
        //        {
        //            myByteList = LogQueue_A.Dequeue();
        //            if ((myByteList == 0x0A) || (myByteList == 0x0D) || (byteMessage_length_A >= byteMessage_max_A))
        //            {
        //                string dataValue = "";
        //                dataValue = Encoding.ASCII.GetString(byteMessage_A);
        //                dataValue = dataValue.Substring(0, byteMessage_length_A);
        //                dt = DateTime.Now;
        //                dataValue = "[Receive_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
        //                log_process("A", dataValue);
        //                log_process("All", dataValue);
        //                byteMessage_length_A = 0;
        //            }
        //            else
        //            {
        //                byteMessage_A[byteMessage_length_A] = myByteList;
        //                byteMessage_length_A++;
        //            }
        //        }
        //    }
        //}

        const int byteTemperature_max = 64;
        byte[] byteTemperature = new byte[byteTemperature_max];
        int byteTemperature_length = 0;

        private void log_temperature(byte ch)
        {
            const int packet_len = 16;
            const int header_offset_1 = -16;
            const int header_offset_2 = -15;
            const int temp_ch_offset = -14;
            const int temp_unit_02 = -13;
            const int temp_unit_01 = -12;
            const int temp_polarity_offset = -11;
            const int temp_dp_offset = -10;
            const int temp_data8_offset = -9;
            const int temp_data7_offset = -8;
            const int temp_data6_offset = -7;
            const int temp_data5_offset = -6;
            const int temp_data4_offset = -5;
            const int temp_data3_offset = -4;
            const int temp_data2_offset = -3;
            const int temp_data1_offset = -2;
            const double temp_abs_value = 0.05;

            // If data_buffer is too long, cut off data not needed
            if (byteTemperature_length >= byteTemperature_max)
            {
                int destinationIndex = 0;
                for (int i = (byteTemperature_max - packet_len); i < byteTemperature_max; i++)
                {
                    byteTemperature[destinationIndex++] = byteTemperature[i];
                }
                byteTemperature_length = destinationIndex;
            }

            byteTemperature[byteTemperature_length] = ch;
            byteTemperature_length++;

            if (ch == 0x0D)
            {
                if (((byteTemperature_length + header_offset_1) >= 0) &&
                     (byteTemperature[byteTemperature_length + header_offset_1] == 0x02) &&
                     (byteTemperature[byteTemperature_length + header_offset_2] == '4'))
                {
                    // Packet is valid here
                    if (byteTemperature[byteTemperature_length + temp_ch_offset] == Temperature_Data.temperatureChannel)
                    {
                        // Channel number is checked and ok here
                        if ((byteTemperature[byteTemperature_length + temp_unit_02] == '0'))
                        {
                            if ((byteTemperature[byteTemperature_length + temp_unit_01] == '1')
                                || (byteTemperature[byteTemperature_length + temp_unit_01] == '2'))
                            {
                                if ((byteTemperature[byteTemperature_length + temp_data1_offset] != 0x18))
                                {
                                    // data is valid
                                    int DP_convert = '0';
                                    int byteArray_position = 0;
                                    byte[] byteArray = new byte[8];
                                    for (int pos = byteTemperature_length + temp_data8_offset;
                                                pos <= (byteTemperature_length + temp_data1_offset);
                                                pos++)
                                    {
                                        byteArray[byteArray_position] = byteTemperature[pos];
                                        byteArray_position++;
                                    }

                                    string tempSubstring = System.Text.Encoding.Default.GetString(byteArray);
                                    double digit = Math.Pow(10, Convert.ToInt64(byteTemperature[byteTemperature_length + temp_dp_offset] - DP_convert));
                                    double currentTemperature = Convert.ToDouble(Convert.ToInt32(tempSubstring) / digit);

                                    // is value negative?
                                    if (byteTemperature[byteTemperature_length + temp_polarity_offset] == '1')
                                    {
                                        currentTemperature = -currentTemperature;
                                    }

                                    // is value Fahrenheit?
                                    if (byteTemperature[byteTemperature_length + temp_unit_01] == '2')
                                    { 
                                        currentTemperature = (currentTemperature - 32) / 1.8;
                                        currentTemperature = Math.Round((currentTemperature),2,MidpointRounding.AwayFromZero);
                                    }

                                    // check whether 2 temperatures are close enough
                                    if (Math.Abs(previousTemperature-currentTemperature) >= temp_abs_value)
                                    {
                                        previousTemperature = currentTemperature;
                                        foreach (Temperature_Data item in temperatureList)
                                        {
                                            if (item.temperatureList == currentTemperature &&
                                                item.temperatureShot == true)
                                            {
                                                Console.WriteLine("~~~ targetTemperature ~~~ " + previousTemperature + " ~~~ currentTemperature ~~~ " + currentTemperature);
                                                temperatureDouble.Enqueue(currentTemperature);
                                                Console.WriteLine("~~~ Enqueue temperature ~~~ " + currentTemperature);
                                            }

                                            if (item.temperatureList == currentTemperature &&
                                                item.temperaturePause == true)
                                            {
                                                label_Command.Text = "Condition: " + item.temperatureList + ", PAUSE: " + currentTemperature;
                                                button_Pause.PerformClick();
                                                Console.WriteLine("Temperature: " + currentTemperature + "~~~~~~~~~Temperature matched. Pause the schedule.~~~~~~~~~");
                                            }

                                            if (item.temperatureList == currentTemperature &&
                                                     item.temperaturePort != "" &&
                                                     item.temperatureLog != "" &&
                                                     item.temperatureNewline != "")
                                            {
                                                label_Command.Text = "Condition: " + item.temperatureList + ", Log: " + currentTemperature;
                                                if (item.temperatureLog.Contains('|'))
                                                {
                                                    string[] logArray = item.temperatureLog.Split('|');
                                                    switch (item.temperaturePort)
                                                    {
                                                        case "A":
                                                            for (int i = 0; i < logArray.Length; i++)
                                                                ReplaceNewLine(GlobalData.m_SerialPort_A, logArray[i], item.temperatureNewline);
                                                            break;
                                                        case "B":
                                                            for (int i = 0; i < logArray.Length; i++)
                                                                ReplaceNewLine(GlobalData.m_SerialPort_B, logArray[i], item.temperatureNewline);
                                                            break;
                                                        case "C":
                                                            for (int i = 0; i < logArray.Length; i++)
                                                                ReplaceNewLine(GlobalData.m_SerialPort_C, logArray[i], item.temperatureNewline);
                                                            break;
                                                        case "D":
                                                            for (int i = 0; i < logArray.Length; i++)
                                                                ReplaceNewLine(GlobalData.m_SerialPort_D, logArray[i], item.temperatureNewline);
                                                            break;
                                                        case "E":
                                                            for (int i = 0; i < logArray.Length; i++)
                                                                ReplaceNewLine(GlobalData.m_SerialPort_E, logArray[i], item.temperatureNewline);
                                                            break;
                                                    }
                                                }
                                                else
                                                {
                                                    switch (item.temperaturePort)
                                                    {
                                                        case "A":
                                                            ReplaceNewLine(GlobalData.m_SerialPort_A, item.temperatureLog, item.temperatureNewline);
                                                            break;
                                                        case "B":
                                                            ReplaceNewLine(GlobalData.m_SerialPort_B, item.temperatureLog, item.temperatureNewline);
                                                            break;
                                                        case "C":
                                                            ReplaceNewLine(GlobalData.m_SerialPort_C, item.temperatureLog, item.temperatureNewline);
                                                            break;
                                                        case "D":
                                                            ReplaceNewLine(GlobalData.m_SerialPort_D, item.temperatureLog, item.temperatureNewline);
                                                            break;
                                                        case "E":
                                                            ReplaceNewLine(GlobalData.m_SerialPort_E, item.temperatureLog, item.temperatureNewline);
                                                            break;
                                                    }
                                                }
                                                Console.WriteLine("Temperature: " + currentTemperature + "~~~~~~~~~Temperature matched. Send the log to device.~~~~~~~~~");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                byteTemperature_length = 0;
            }
        }

        //private void logA_temperature()
        //{
        //    byte myByteList;

        //    while (TemperatureQueue_A.Count > 0)
        //    {
        //        myByteList = TemperatureQueue_A.Dequeue();
        //        if (myByteList == 0x0D)
        //        {
        //            byteTemperature[byteTemperature_length] = myByteList;
        //            if (byteTemperature[0] == 0x02 && byteTemperature[1] == 0x34 && byteTemperature[2] == Temperature_Data.temperatureChannel && byteTemperature[14] != 0x18 && byteTemperature[15] == 0x0D)
        //            {
        //                int DP_convert = 48;
        //                int byteArray_position = 0;
        //                byte[] byteArray = new byte[4];
        //                for (int byteMessage_position = 11; byteMessage_position < 15; byteMessage_position++)
        //                {
        //                    byteArray[byteArray_position] = byteTemperature[byteMessage_position];
        //                    byteArray_position++;
        //                }
        //                string tempSubstring = System.Text.Encoding.Default.GetString(byteArray);
        //                double digit = Math.Pow(10, Convert.ToInt32(byteTemperature[6] - DP_convert));
        //                double currentTemperature = Math.Round(Convert.ToDouble(Convert.ToInt32(tempSubstring)) / digit, 0, MidpointRounding.AwayFromZero);
        //                if (targetTemperature != currentTemperature)
        //                {
        //                    Console.WriteLine("~~~ targetTemperature ~~~ " + targetTemperature + " ~~~ currentTemperature ~~~ " + currentTemperature);
        //                    temperatureDouble.Enqueue(currentTemperature);
        //                    Console.WriteLine("~~~ Enqueue temperature ~~~ " + currentTemperature);
        //                    targetTemperature = currentTemperature;
        //                }
        //                byteTemperature_length = 0;
        //            }
        //            else
        //            {
        //                byteTemperature_length = 0;
        //            }
        //        }
        //        else
        //        {
        //            byteTemperature[byteTemperature_length] = myByteList;
        //            byteTemperature_length++;
        //        }
        //    }
        //}

        const int byteChamber_max = 64;
        byte[] byteChamber = new byte[byteChamber_max];
        int byteChamber_length = 0;

        private void logA_chamber(byte ch)
        {
            const int header_data1_offset = -9;
            const int header_data2_offset = -8;
            const int length_data_offset = -7;
            const int data_actual2_offset = -6;
            const int data_actual1_offset = -5;
            const int data_target2_offset = -4;
            const int data_target1_offset = -3;
            const int crc16_highbit_offset = -2;
            const int crc16_lowbit_offset = -1;

            byteChamber[byteChamber_length] = ch;
            byteChamber_length++;

            if ((byteChamber[byteChamber_length + header_data1_offset] == 0x01) &&
                (byteChamber[byteChamber_length + header_data2_offset] == 0x03) &&
                (byteChamber[byteChamber_length + length_data_offset] == 0x04))
            {
                byte[] byteActual = new byte[2];
                byte[] byteTarget = new byte[2];
                byteActual[0] = byteChamber[byteChamber_length + data_actual2_offset];
                byteActual[1] = byteChamber[byteChamber_length + data_actual1_offset];
                byteTarget[0] = byteChamber[byteChamber_length + data_target2_offset];
                byteTarget[1] = byteChamber[byteChamber_length + data_target1_offset];
                string stringActual = System.Text.Encoding.Default.GetString(byteActual);
                string stringTarget = System.Text.Encoding.Default.GetString(byteTarget);
                int intActual = Convert.ToInt32(stringActual, 16);
                int intTarget = Convert.ToInt32(stringTarget, 16);
            }
            byteChamber_length = 0;
        }

        /*
        //Jeremy code
        private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int data = 500;
            Object serial_data_obj;
            SerialPort sp = (SerialPort)sender;
            PortDataContainer.PortDictionary.TryGetValue(sp.PortName, out serial_data_obj);
            PortDataContainer serial_port_data = (PortDataContainer)serial_data_obj;

            if (serial_port_data == null)
                return;

            try
            {
                int data_to_read = serialPort1.BytesToRead;
                if (data_to_read > 0)
                {
                    //if (data_to_read > data)
                    //    data_to_read = data;
                    Byte[] dataset = new byte[data_to_read];
                    serialPort1.Read(dataset, 0, data_to_read);
                    List<Byte> data_list = dataset.ToList();

                    SerialReceivedData enqueue_data = new SerialReceivedData();
                    enqueue_data.SetData(data_list);
                    enqueue_data.SetTimeStamp(DateTime.Now);
                    serial_port_data.data_queue.Enqueue(enqueue_data);
                    PortDataContainer.data_available = true;

                    Thread DataThread = new Thread(new ThreadStart(test));
                    DataThread.Start();
                    //SerialPortTxtbox1(Encoding.ASCII.GetString(dataset1), textBox1);
                    //textBoxBuffer.Put(Encoding.ASCII.GetString(dataset1));
                    //string s = "";
                    //textBox1.Invoke(this.myDelegate1, new Object[] { s });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        void Th1()
        {
                SerialPortTxtbox1(Encoding.ASCII.GetString(dataset1), textBox1);
                Thread.Sleep(1);
        }

        //委派 txtbox1
        private delegate void UpdateUISerialPort1(string value, Control ctrl);
        private void SerialPortTxtbox1(string value, Control ctrl)
        {
            if (ctrl == null || value == null) return;
            if (ctrl.InvokeRequired)
            {
                UpdateUISerialPort1 uu = new UpdateUISerialPort1(SerialPortTxtbox1);
                this.Invoke(uu, value, ctrl);
            }
            else
            {
                if (ctrl == textBox1)
                {
                    DateTime dt;
                    dt = DateTime.Now;
                    value = value.Replace(Environment.NewLine, "\r\n" + "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  "); //OK
                    textBox1.AppendText(value);
                }
            }
        }


        bool test_is_running = false;

        private void test()
        {
            while (test_is_running == true) { Thread.Sleep(1); }
            
            //while (true)
            {
                test_is_running = true;
                if (PortDataContainer.data_available == true)
                {
                    PortDataContainer.data_available = false;
                    foreach (var port in PortDataContainer.PortDictionary)
                    {
                        PortDataContainer serial_port_data = (PortDataContainer)port.Value;
                        while (serial_port_data.data_queue.Count > 0)
                        {
                            SerialReceivedData dequeue_data = serial_port_data.data_queue.Dequeue();
                            Byte[] dataset = dequeue_data.GetData().ToArray();
                            DateTime dt = dequeue_data.GetTimeStamp();

                            // The following code is almost the same as before

                            int index = 0;
                            int data_to_read = dequeue_data.GetData().Count;
                            while (data_to_read > 0)
                            {
                                serial_port_data.log_data.Enqueue(dataset[index]);
                                index++;
                                data_to_read--;
                            }
                            
                            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
                            {
                                // hex to string
                                string hexValues = BitConverter.ToString(dataset).Replace("-", "");
                                //DateTime.Now.ToShortTimeString();
                                //dt = DateTime.Now;

                                // Joseph
                                hexValues = hexValues.Replace(Environment.NewLine, "\r\n" + "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  "); //OK
                                // hexValues = String.Concat("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n");
                                log_text = string.Concat(log_text, hexValues);
                                // textBox1.AppendText(hexValues);
                                // End

                                // Jeremy
                                // textBox1.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
                                // textBox1.AppendText(hexValues + "\r\n");
                                // End
                            }
                            else
                            {
                                // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
                                string strValues = Encoding.ASCII.GetString(dataset);
                                dt = DateTime.Now;
                                strValues = strValues.Replace(Environment.NewLine, "\r\n" + "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  "); //OK
                                log_text = string.Concat(log_text, strValues);

                                //textBox1.AppendText(text);
                            }
                            
                        }
                    }
                }
            }
            test_is_running = false;
        }
        */
        #endregion

        #region -- 接受SerialPort2資料 --
        // private void SerialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        // {
        //     try
        //     {
        //         int data_to_read = PortB.BytesToRead;
        //         if (data_to_read > 0)
        //         {
        //             byte[] dataset = new byte[data_to_read];

        //             PortB.Read(dataset, 0, data_to_read);
        //             int index = 0;
        //             while (data_to_read > 0)
        //             {
        //                 SearchLogQueue_B.Enqueue(dataset[index]);
        //                 index++;
        //                 data_to_read--;
        //             }

        //             DateTime dt;
        //             if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
        //             {
        //                 // hex to string
        //                 string hexValues = BitConverter.ToString(dataset).Replace("-", "");
        //                 DateTime.Now.ToShortTimeString();
        //                 dt = DateTime.Now;

        //                 // Joseph
        //                 hexValues = "[Receive_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n"; //OK
        //                 logB_text = string.Concat(logB_text, hexValues);
        //                 logAll_text = string.Concat(logAll_text, hexValues);
        //                 // textBox2.AppendText(hexValues);
        //                 // End

        //                 // Jeremy
        //                 // textBox2.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
        //                 // textBox2.AppendText(hexValues + "\r\n");
        //                 // End
        //             }
        //             else
        //             {
        //                 // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
        //                 string strValues = Encoding.ASCII.GetString(dataset);
        //                 dt = DateTime.Now;

        //                 if (strValues.Contains("\r"))
        //                 {
        //                     string[] log = strValues.Split('\r');
        //                     foreach (string s in log)
        //                     {
        //                         Thread.Sleep(500);
        //                         strValues1 = "[Receive_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + s + "\r\n";
        //                         if (//s.Substring(2, 1) == Temperature_Data.temperatureChannel &&
        //                             !s.Contains('\u0018') &&
        //                             strValues1.Substring(strValues1.IndexOf('\u0002') + 1, strValues1.IndexOf('\r') - strValues1.IndexOf('\u0002') - 1).Length == 14 &&
        //                             TemperatureIsFound)
        //                         {
        //                             string tempSubstring = strValues1.Substring(strValues1.IndexOf('\u0002') + 11, 4);
        //                             double digit = Math.Pow(10, Convert.ToInt32(strValues1.Substring(strValues1.IndexOf('\u0002') + 6, 1)));
        //                             double currentTemperature = Math.Round(Convert.ToDouble(Convert.ToInt32(tempSubstring)) / digit, 0, MidpointRounding.AwayFromZero);
        //                             if (previousTemperature != currentTemperature)
        //                             {
        //                                 //temperatureDouble.Enqueue(strValues1);
        //                                 previousTemperature = currentTemperature;
        //                             }
        //                         }
        //                         logB_text = string.Concat(logB_text, strValues1);
        //	logAll_text = string.Concat(logAll_text, strValues);
        //                     }
        //                 }
        //                 else
        //                 {
        //                     strValues = "[Receive_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strValues + "\r\n";
        //                     logB_text = string.Concat(logB_text, strValues);
        //logAll_text = string.Concat(logAll_text, strValues);
        //                 }
        //                 //textBox2.AppendText(strValues);
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine(ex.Message);
        //     }
        // }


        //private void logB_analysis()
        //{
        //    while (PortB.IsOpen == true)
        //    {
        //        int data_to_read = PortB.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            byte[] dataset = new byte[data_to_read];
        //            PortB.Read(dataset, 0, data_to_read);

        //            for (int index = 0; index < data_to_read; index++)
        //            {
        //                byte input_ch = dataset[index];
        //                logB_recorder(input_ch);
        //                if (TemperatureIsFound == true)
        //                {
        //                    log_temperature(input_ch);
        //                }
        //            }
        //        }
        //        //else
        //        //{
        //        //    logB_recorder(0x00,true); // tell log_recorder no more data for now.
        //        //}
        //    }
        //}

        byte[] byteMessage_B = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length_B = 0;

        private void logB_recorder(byte ch, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage_B[byteMessage_length_B] = ch;
                    byteMessage_length_B++;
                }
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_B >= byteMessage_max_Hex) /*|| (SaveToLog == true)*/)
                {
                    string dataValue = BitConverter.ToString(byteMessage_B).Replace("-", "").Substring(0, byteMessage_length_B * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logB_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_B = 0;
                }
            }
            else
            {
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_B >= byteMessage_max_Ascii))
                {
                    string dataValue = Encoding.ASCII.GetString(byteMessage_B).Substring(0, byteMessage_length_B);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logB_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_B = 0;
                }
                else
                {
                    byteMessage_B[byteMessage_length_B] = ch;
                    byteMessage_length_B++;
                }
            }
        }
        #endregion

        #region -- 接受SerialPort3資料 --
        // private void SerialPort3_DataReceived(object sender, SerialDataReceivedEventArgs e)
        // {
        //     try
        //     {
        //         int data_to_read = PortC.BytesToRead;
        //         if (data_to_read > 0)
        //         {
        //             byte[] dataset = new byte[data_to_read];

        //             PortC.Read(dataset, 0, data_to_read);
        //             int index = 0;
        //             while (data_to_read > 0)
        //             {
        //                 SearchLogQueue_C.Enqueue(dataset[index]);
        //                 index++;
        //                 data_to_read--;
        //             }

        //             DateTime dt;
        //             if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
        //             {
        //                 // hex to string
        //                 string hexValues = BitConverter.ToString(dataset).Replace("-", "");
        //                 DateTime.Now.ToShortTimeString();
        //                 dt = DateTime.Now;

        //                 // Joseph
        //                 hexValues = "[Receive_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n"; //OK
        //                 logC_text = string.Concat(logC_text, hexValues);
        //                 logAll_text = string.Concat(logAll_text, hexValues);
        //                 // textBox3.AppendText(hexValues);
        //                 // End

        //                 // Jeremy
        //                 // textBox3.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
        //                 // textBox3.AppendText(hexValues + "\r\n");
        //                 // End
        //             }
        //             else
        //             {
        //                 // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
        //                 string strValues = Encoding.ASCII.GetString(dataset);
        //                 dt = DateTime.Now;

        //                 if (strValues.Contains("\r"))
        //                 {
        //                     string[] log = strValues.Split('\r');
        //                     foreach (string s in log)
        //                     {
        //                         Thread.Sleep(500);
        //                         strValues1 = "[Receive_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + s + "\r\n";
        //                         if (//s.Substring(2, 1) == Temperature_Data.temperatureChannel &&
        //                             !s.Contains('\u0018') &&
        //                             strValues1.Substring(strValues1.IndexOf('\u0002') + 1, strValues1.IndexOf('\r') - strValues1.IndexOf('\u0002') - 1).Length == 14 &&
        //                             TemperatureIsFound)
        //                         {
        //                             string tempSubstring = strValues1.Substring(strValues1.IndexOf('\u0002') + 11, 4);
        //                             double digit = Math.Pow(10, Convert.ToInt32(strValues1.Substring(strValues1.IndexOf('\u0002') + 6, 1)));
        //                             double currentTemperature = Math.Round(Convert.ToDouble(Convert.ToInt32(tempSubstring)) / digit, 0, MidpointRounding.AwayFromZero);
        //                             if (previousTemperature != currentTemperature)
        //                             {
        //                                 //temperatureDouble.Enqueue(strValues1);
        //                                 previousTemperature = currentTemperature;
        //                             }
        //                         }
        //                         logC_text = string.Concat(logC_text, strValues1);
        //	logAll_text = string.Concat(logAll_text, strValues1);
        //                     }
        //                 }
        //                 else
        //                 {
        //                     strValues = "[Receive_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strValues + "\r\n";
        //                     logC_text = string.Concat(logC_text, strValues);
        //logAll_text = string.Concat(logAll_text, strValues);
        //                 }
        //                 //textBox3.AppendText(strValues);
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine(ex.Message);
        //     }
        // }

        //private void logC_analysis()
        //{
        //    while (PortC.IsOpen == true)
        //    {
        //        int data_to_read = PortC.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            byte[] dataset = new byte[data_to_read];
        //            PortC.Read(dataset, 0, data_to_read);

        //            for (int index = 0; index < data_to_read; index++)
        //            {
        //                byte input_ch = dataset[index];
        //                logC_recorder(input_ch);
        //                if (TemperatureIsFound == true)
        //                {
        //                    log_temperature(input_ch);
        //                }
        //            }
        //        }
        //        //else
        //        //{
        //        //    logD_recorder(0x00,true); // tell log_recorder no more data for now.
        //        //}
        //    }
        //}

        byte[] byteMessage_C = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length_C = 0;

        private void logC_recorder(byte ch, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage_C[byteMessage_length_C] = ch;
                    byteMessage_length_C++;
                }
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_C >= byteMessage_max_Hex) /*|| (SaveToLog == true)*/)
                {
                    string dataValue = BitConverter.ToString(byteMessage_C).Replace("-", "").Substring(0, byteMessage_length_C * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logC_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_C = 0;
                }
            }
            else
            {
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_C >= byteMessage_max_Ascii))
                {
                    string dataValue = Encoding.ASCII.GetString(byteMessage_C).Substring(0, byteMessage_length_C);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logC_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_C = 0;
                }
                else
                {
                    byteMessage_C[byteMessage_length_C] = ch;
                    byteMessage_length_C++;
                }
            }
        }

        #endregion

        #region -- 接受SerialPort4資料 --
        // private void SerialPort4_DataReceived(object sender, SerialDataReceivedEventArgs e)
        // {
        //     try
        //     {
        //         int data_to_read = PortD.BytesToRead;
        //         if (data_to_read > 0)
        //         {
        //             byte[] dataset = new byte[data_to_read];

        //             PortD.Read(dataset, 0, data_to_read);
        //             int index = 0;
        //             while (data_to_read > 0)
        //             {
        //                 SearchLogQueue_D.Enqueue(dataset[index]);
        //                 index++;
        //                 data_to_read--;
        //             }

        //             DateTime dt;
        //             if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
        //             {
        //                 // hex to string
        //                 string hexValues = BitConverter.ToString(dataset).Replace("-", "");
        //                 DateTime.Now.ToShortTimeString();
        //                 dt = DateTime.Now;

        //                 // Joseph
        //                 hexValues = "[Receive_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n"; //OK
        //                 logD_text = string.Concat(logD_text, hexValues);
        //                 logAll_text = string.Concat(logAll_text, hexValues);
        //                 // textBox4.AppendText(hexValues);
        //                 // End

        //                 // Jeremy
        //                 // textBox4.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
        //                 // textBox4.AppendText(hexValues + "\r\n");
        //                 // End
        //             }
        //             else
        //             {
        //                 // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
        //                 string strValues = Encoding.ASCII.GetString(dataset);
        //                 dt = DateTime.Now;

        //                 if (strValues.Contains("\r"))
        //                 {
        //                     string[] log = strValues.Split('\r');
        //                     foreach (string s in log)
        //                     {
        //                         Thread.Sleep(500);
        //                         strValues1 = "[Receive_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + s + "\r\n";
        //                         if (//s.Substring(2, 1) == Temperature_Data.temperatureChannel &&
        //                             !s.Contains('\u0018') &&
        //                             strValues1.Substring(strValues1.IndexOf('\u0002') + 1, strValues1.IndexOf('\r') - strValues1.IndexOf('\u0002') - 1).Length == 14 &&
        //                             TemperatureIsFound)
        //                         {
        //                             string tempSubstring = strValues1.Substring(strValues1.IndexOf('\u0002') + 11, 4);
        //                             double digit = Math.Pow(10, Convert.ToInt32(strValues1.Substring(strValues1.IndexOf('\u0002') + 6, 1)));
        //                             double currentTemperature = Math.Round(Convert.ToDouble(Convert.ToInt32(tempSubstring)) / digit, 0, MidpointRounding.AwayFromZero);
        //                             if (previousTemperature != currentTemperature)
        //                             {
        //                                 //temperatureDouble.Enqueue(strValues1);
        //                                 previousTemperature = currentTemperature;
        //                             }
        //                         }
        //                         logD_text = string.Concat(logD_text, strValues1);
        //	logAll_text = string.Concat(logAll_text, strValues1);
        //                     }
        //                 }
        //                 else
        //                 {
        //                     strValues = "[Receive_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strValues + "\r\n";
        //                     logD_text = string.Concat(logD_text, strValues);
        //logAll_text = string.Concat(logAll_text, strValues);
        //                 }
        //                 //textBox4.AppendText(strValues);
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine(ex.Message);
        //     }
        // }

        //private void logD_analysis()
        //{
        //    while (PortD.IsOpen == true)
        //    {
        //        int data_to_read = PortD.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            byte[] dataset = new byte[data_to_read];
        //            PortD.Read(dataset, 0, data_to_read);

        //            for (int index = 0; index < data_to_read; index++)
        //            {
        //                byte input_ch = dataset[index];
        //                logD_recorder(input_ch);
        //                if (TemperatureIsFound == true)
        //                {
        //                    log_temperature(input_ch);
        //                }
        //            }
        //        }
        //        //else
        //        //{
        //        //    logD_recorder(0x00,true); // tell log_recorder no more data for now.
        //        //}
        //    }
        //}

        byte[] byteMessage_D = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length_D = 0;

        private void logD_recorder(byte ch, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage_D[byteMessage_length_D] = ch;
                    byteMessage_length_D++;
                }
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_D >= byteMessage_max_Hex) /*|| (SaveToLog == true)*/)
                {
                    string dataValue = BitConverter.ToString(byteMessage_D).Replace("-", "").Substring(0, byteMessage_length_D * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logD_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_D = 0;
                }
            }
            else
            {
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_D >= byteMessage_max_Ascii))
                {
                    string dataValue = Encoding.ASCII.GetString(byteMessage_D).Substring(0, byteMessage_length_D);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logD_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_D = 0;
                }
                else
                {
                    byteMessage_D[byteMessage_length_D] = ch;
                    byteMessage_length_D++;
                }
            }
        }

        #endregion

        #region -- 接受SerialPort5資料 --
        // private void SerialPort5_DataReceived(object sender, SerialDataReceivedEventArgs e)
        // {
        //     try
        //     {
        //         int data_to_read = PortE.BytesToRead;
        //         if (data_to_read > 0)
        //         {
        //             byte[] dataset = new byte[data_to_read];

        //             PortE.Read(dataset, 0, data_to_read);
        //             int index = 0;
        //             while (data_to_read > 0)
        //             {
        //                 SearchLogQueue_E.Enqueue(dataset[index]);
        //                 index++;
        //                 data_to_read--;
        //             }

        //             DateTime dt;
        //             if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
        //             {
        //                 // hex to string
        //                 string hexValues = BitConverter.ToString(dataset).Replace("-", "");
        //                 DateTime.Now.ToShortTimeString();
        //                 dt = DateTime.Now;

        //                 // Joseph
        //                 hexValues = "[Receive_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + hexValues + "\r\n"; //OK
        //                 logE_text = string.Concat(logE_text, hexValues);
        //                 logAll_text = string.Concat(logAll_text, hexValues);
        //                 // textBox5.AppendText(hexValues);
        //                 // End

        //                 // Jeremy
        //                 // textBox5.AppendText("[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  ");
        //                 // textBox5.AppendText(hexValues + "\r\n");
        //                 // End
        //             }
        //             else
        //             {
        //                 // string text = String.Concat(Encoding.ASCII.GetString(dataset).Where(c => c != 0x00));
        //                 string strValues = Encoding.ASCII.GetString(dataset);
        //                 dt = DateTime.Now;

        //                 if (strValues.Contains("\r"))
        //                 {
        //                     string[] log = strValues.Split('\r');
        //                     foreach (string s in log)
        //                     {
        //                         Thread.Sleep(500);
        //                         strValues1 = "[Receive_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + s + "\r\n";
        //                         if (//s.Substring(2, 1) == Temperature_Data.temperatureChannel &&
        //                             !s.Contains('\u0018') &&
        //                             strValues1.Substring(strValues1.IndexOf('\u0002') + 1, strValues1.IndexOf('\r') - strValues1.IndexOf('\u0002') - 1).Length == 14 &&
        //                             TemperatureIsFound)
        //                         {
        //                             string tempSubstring = strValues1.Substring(strValues1.IndexOf('\u0002') + 11, 4);
        //                             double digit = Math.Pow(10, Convert.ToInt32(strValues1.Substring(strValues1.IndexOf('\u0002') + 6, 1)));
        //                             double currentTemperature = Math.Round(Convert.ToDouble(Convert.ToInt32(tempSubstring)) / digit, 0, MidpointRounding.AwayFromZero);
        //                             if (previousTemperature != currentTemperature)
        //                             {
        //                                 //temperatureDouble.Enqueue(strValues1);
        //                                 previousTemperature = currentTemperature;
        //                             }
        //                         }
        //                         logE_text = string.Concat(logE_text, strValues1);
        //	logAll_text = string.Concat(logAll_text, strValues1);
        //                     }
        //                 }
        //                 else
        //                 {
        //                     strValues = "[Receive_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strValues + "\r\n";
        //                     logE_text = string.Concat(logE_text, strValues);
        //logAll_text = string.Concat(logAll_text, strValues);
        //                 }
        //                 //textBox5.AppendText(strValues);
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine(ex.Message);
        //     }
        // }

        //private void logE_analysis()
        //{
        //    while (PortE.IsOpen == true)
        //    {
        //        int data_to_read = PortE.BytesToRead;
        //        if (data_to_read > 0)
        //        {
        //            byte[] dataset = new byte[data_to_read];
        //            PortE.Read(dataset, 0, data_to_read);

        //            for (int index = 0; index < data_to_read; index++)
        //            {
        //                byte input_ch = dataset[index];
        //                logE_recorder(input_ch);
        //                if (TemperatureIsFound == true)
        //                {
        //                    log_temperature(input_ch);
        //                }
        //            }
        //        }
        //        //else
        //        //{
        //        //    logB_recorder(0x00,true); // tell log_recorder no more data for now.
        //        //}
        //    }
        //}

        byte[] byteMessage_E = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length_E = 0;

        private void logE_recorder(byte ch, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage_E[byteMessage_length_E] = ch;
                    byteMessage_length_E++;
                }
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_E >= byteMessage_max_Hex) /*|| (SaveToLog == true)*/)
                {
                    string dataValue = BitConverter.ToString(byteMessage_E).Replace("-", "").Substring(0, byteMessage_length_E * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logE_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_E = 0;
                }
            }
            else
            {
                if ((ch == 0x0A) || (ch == 0x0D) || (byteMessage_length_E >= byteMessage_max_Ascii))
                {
                    string dataValue = Encoding.ASCII.GetString(byteMessage_E).Substring(0, byteMessage_length_E);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    logDumpping.LogCat(ref logE_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                    byteMessage_length_E = 0;
                }
                else
                {
                    byteMessage_E[byteMessage_length_E] = ch;
                    byteMessage_length_E++;
                }
            }
        }

        #endregion

        #region -- 接受ArduinoPort資料 --
        private void serialport_arduino_datareceived(object sender, SerialDataReceivedEventArgs e)
        {
            string Read_Arduino_Data;
            try
            {
                if (serialPort_Arduino.IsOpen)     //此处可能没有必要判断是否打开串口，但为了严谨性，我还是加上了
                {
                    byte[] byteRead = new byte[serialPort_Arduino.BytesToRead];    //BytesToRead:sp1接收的字符个数
                    string dataValue = serialPort_Arduino.ReadLine(); //注意：回车换行必须这样写，单独使用"\r"和"\n"都不会有效果
                    if (dataValue.Contains("io i"))
                        Read_Arduino_Data = dataValue;
                    serialPort_Arduino.DiscardInBuffer();                      //清空SerialPort控件的Buffer 
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        dataValue = "[Receive_Port_Arduino] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                    }
                    log.Debug("Ardroud receive:" + dataValue);
                    GlobalData.Arduino_recFlag = true;
                    logDumpping.LogCat(ref arduino_text, dataValue);
                    logDumpping.LogCat(ref logAll_text, dataValue);
                }
                else
                {
                    Console.WriteLine("请打开某个串口", "错误提示");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message, "出错提示");
            }
        }

        #endregion

        #region -- 儲存SerialPort的log --
        private void Serialportsave(string Port)
        {
            string fName = "";

            // 讀取ini中的路徑
            fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
            switch (Port)
            {
                case "A":
                    string t = fName + "\\_PortA_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logA_text);
                    MYFILE.Close();
                    Txtbox1("", textBox_serial);
                    logA_text = String.Empty;
                    break;
                case "B":
                    t = fName + "\\_PortB_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logB_text);
                    MYFILE.Close();
                    logB_text = String.Empty;
                    break;
                case "C":
                    t = fName + "\\_PortC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logC_text);
                    MYFILE.Close();
                    logC_text = String.Empty;
                    break;
                case "D":
                    t = fName + "\\_PortD_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logD_text);
                    MYFILE.Close();
                    logD_text = String.Empty;
                    break;
                case "E":
                    t = fName + "\\_PortE_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logE_text);
                    MYFILE.Close();
                    logE_text = String.Empty;
                    break;
                case "Arduino":
                    t = fName + "\\_Arduino_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(arduino_text);
                    MYFILE.Close();
                    arduino_text = String.Empty;
                    break;
                case "CA310":
                    t = fName + "\\_CA310_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".csv";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(ca310_text);
                    MYFILE.Close();
                    ca310_text = String.Empty;
                    break;
                case "Canbus":
                    t = fName + "\\_Canbus_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(canbus_text);
                    MYFILE.Close();
                    canbus_text = String.Empty;
                    break;
                case "KlinePort":
                    t = fName + "\\_Kline_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(kline_text);
                    MYFILE.Close();
                    kline_text = String.Empty;
                    break;
                case "All":
                    t = fName + "\\_AllPort_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logAll_text);
                    MYFILE.Close();
                    logAll_text = String.Empty;
                    break;
                case "Debug":
                    t = fName + "\\_Debug_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(debug_text);
                    MYFILE.Close();
                    debug_text = String.Empty;
                    break;
            }
        }
        #endregion

        #region -- Save CA310/210 report --
        private void createCA210folder()
        {
            string csvFolder = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\" + "Measure_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            if (Directory.Exists(csvFolder))
            {

            }
            else
            {
                Directory.CreateDirectory(csvFolder);
                GlobalData.MeasurePath = csvFolder;
                minolta_csv_report = "Sx, Sy, Lv, T, duv, Display mode, X, Y, Z, Date, Time, Scenario, Now measure count, Target measure count, Backlight sensor, Thanmal sensor, \r\n";
            }
        }

        private void saveCA210csv(string filename)
        {
            string folder = GlobalData.MeasurePath;
            string file = filename;
            if (file == "")
                file = GlobalData.MeasurePath + "\\Minolta_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            else
                file = GlobalData.MeasurePath + "\\" + file + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            StreamWriter MYFILE = new StreamWriter(file, false, Encoding.ASCII);
            MYFILE.Write(minolta_csv_report);
            MYFILE.Close();
            minolta_csv_report = "Sx, Sy, Lv, T, duv, Display mode, X, Y, Z, Date, Time, Scenario, Now measure count, Target measure count, Backlight sensor, Thanmal sensor, \r\n";
        }
        #endregion

        string logReceived = string.Empty;
        string logAdd = string.Empty;
        bool ChamberCheck = false;
        bool PowerSupplyCheck = false;
        double previousTemperature = -300;
        bool timer_matched = false;

        private void timer_duringShot_Tick(object sender, EventArgs e)
        {
            timer_matched = true;
        }

        #region -- 關鍵字比對 - serialport_1 --
        private void MyLog1Camd()
        {
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortA_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));
            byte myByteList;
            const int byteMessage_max = 1024;
            byte[] byteMessage = new byte[byteMessage_max];
            int byteMessage_length = 0;

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_A.Count > 0)
                {
                    myByteList = SearchLogQueue_A.Dequeue();

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport1", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region 0x0A || 0x0D
                        if ((myByteList == 0x0A) || (myByteList == 0x0D) || (byteMessage_length >= byteMessage_max))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string compare_words = Encoding.ASCII.GetString(byteMessage);
                                if (Convert.ToInt32(compare_words.Length - 1) >= 1 && compare_words.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (compare_words.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog1_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(textBox_serial.Text);
                                        MYFILE.Close();
                                        Txtbox1("", textBox_serial);
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox1.AppendText(my_string + '\n');
                            byteMessage_length = 0;
                        }
                        #endregion
                        else
                        {
                            byteMessage[byteMessage_length] = myByteList;
                            byteMessage_length++;
                        }
                    }
                    else
                    {
                        if ((myByteList == 0x0A) || (myByteList == 0x0D))
                        {
                            byteMessage_length = 0;
                        }
                        else
                        {
                            byteMessage[byteMessage_length] = myByteList;
                            byteMessage_length++;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        /*
        private void MyLog1Camd()
        {
            string my_string = "";
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortA_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));
            

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_A.Count > 0)
                {
                    Keyword_SerialPort_A_temp_byte = SearchLogQueue_A.Dequeue();
                    Keyword_SerialPort_A_temp_char = (char)Keyword_SerialPort_A_temp_byte;

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport1", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region \n
                        if ((Keyword_SerialPort_A_temp_char == '\n'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog1_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(textBox_serial.Text);
                                        MYFILE.Close();
                                        Txtbox1("", textBox_serial);
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox1.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        #endregion

                        #region \r
                        else if ((Keyword_SerialPort_A_temp_char == '\r'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);

                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);

                                    //////////////////////////////////////////////////////////////////////Create the compare csv file////////////////////
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog1_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(textBox_serial.Text);
                                        MYFILE.Close();
                                        Txtbox1("", textBox_serial);
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox1.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        #endregion

                        else
                        {
                            my_string = my_string + Keyword_SerialPort_A_temp_char;
                        }
                    }
                    else
                    {
                        if ((Keyword_SerialPort_A_temp_char == '\n'))
                        {
                            //textBox1.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        else if ((Keyword_SerialPort_A_temp_char == '\r'))
                        {
                            //textBox1.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        else
                        {
                            my_string = my_string + Keyword_SerialPort_A_temp_char;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        */
        #endregion

        #region -- 關鍵字比對 - serialport_2 --
        private void MyLog2Camd()
        {
            string my_string = "";
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortB_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_B.Count > 0)
                {
                    Keyword_SerialPort_B_temp_byte = SearchLogQueue_B.Dequeue();
                    Keyword_SerialPort_B_temp_char = (char)Keyword_SerialPort_B_temp_byte;

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport2", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region \n
                        if ((Keyword_SerialPort_B_temp_char == '\n'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logB_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }

                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox2.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        #endregion

                        #region \r
                        else if ((Keyword_SerialPort_B_temp_char == '\r'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);

                                    //////////////////////////////////////////////////////////////////////Create the compare csv file////////////////////
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logB_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox2.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        #endregion

                        else
                        {
                            my_string = my_string + Keyword_SerialPort_B_temp_char;
                        }
                    }
                    else
                    {

                        if ((Keyword_SerialPort_B_temp_char == '\n'))
                        {
                            //textBox2.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        else if ((Keyword_SerialPort_B_temp_char == '\r'))
                        {
                            //textBox2.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        else
                        {
                            my_string = my_string + Keyword_SerialPort_B_temp_char;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        #endregion

        #region -- 關鍵字比對 - serialport_3 --
        private void MyLog3Camd()
        {
            string my_string = "";
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortC_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_C.Count > 0)
                {
                    Keyword_SerialPort_C_temp_byte = SearchLogQueue_C.Dequeue();
                    Keyword_SerialPort_C_temp_char = (char)Keyword_SerialPort_C_temp_byte;

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport3", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region \n
                        if ((Keyword_SerialPort_C_temp_char == '\n'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logB_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }

                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox2.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        #endregion

                        #region \r
                        else if ((Keyword_SerialPort_C_temp_char == '\r'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);

                                    //////////////////////////////////////////////////////////////////////Create the compare csv file////////////////////
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog3_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logC_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        #endregion

                        else
                        {
                            my_string = my_string + Keyword_SerialPort_C_temp_char;
                        }
                    }
                    else
                    {

                        if ((Keyword_SerialPort_C_temp_char == '\n'))
                        {
                            //textBox3.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        else if ((Keyword_SerialPort_C_temp_char == '\r'))
                        {
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        else
                        {
                            my_string = my_string + Keyword_SerialPort_C_temp_char;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        #endregion

        #region -- 關鍵字比對 - serialport_4 --
        private void MyLog4Camd()
        {
            string my_string = "";
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortD_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_D.Count > 0)
                {
                    Keyword_SerialPort_D_temp_byte = SearchLogQueue_D.Dequeue();
                    Keyword_SerialPort_D_temp_char = (char)Keyword_SerialPort_D_temp_byte;

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport4", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region \n
                        if ((Keyword_SerialPort_D_temp_char == '\n'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logB_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }

                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox2.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        #endregion

                        #region \r
                        else if ((Keyword_SerialPort_D_temp_char == '\r'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);

                                    //////////////////////////////////////////////////////////////////////Create the compare csv file////////////////////
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog3_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logC_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        #endregion

                        else
                        {
                            my_string = my_string + Keyword_SerialPort_D_temp_char;
                        }
                    }
                    else
                    {

                        if ((Keyword_SerialPort_D_temp_char == '\n'))
                        {
                            //textBox3.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        else if ((Keyword_SerialPort_D_temp_char == '\r'))
                        {
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        else
                        {
                            my_string = my_string + Keyword_SerialPort_D_temp_char;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        #endregion

        #region -- 關鍵字比對 - serialport_5 --
        private void MyLog5Camd()
        {
            string my_string = "";
            string csvFile = ini12.INIRead(MainSettingPath, "Record", "LogPath", "") + "\\PortE_keyword.csv";
            int[] compare_number = new int[10];
            bool[] send_status = new bool[10];
            int compare_paremeter = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", ""));

            while (StartButtonPressed == true)
            {
                while (SearchLogQueue_E.Count > 0)
                {
                    Keyword_SerialPort_E_temp_byte = SearchLogQueue_E.Dequeue();
                    Keyword_SerialPort_E_temp_char = (char)Keyword_SerialPort_E_temp_byte;

                    if (Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Comport5", "")) == 1 && Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "")) > 0)
                    {
                        #region \n
                        if ((Keyword_SerialPort_E_temp_char == '\n'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logB_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }

                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox2.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        #endregion

                        #region \r
                        else if ((Keyword_SerialPort_E_temp_char == '\r'))
                        {
                            for (int i = 0; i < compare_paremeter; i++)
                            {
                                string compare_string = ini12.INIRead(MainSettingPath, "LogSearch", "Text" + i, "");
                                int compare_num = Convert.ToInt32(ini12.INIRead(MainSettingPath, "LogSearch", "Times" + i, ""));
                                string[] ewords = my_string.Split(new string[] { compare_string }, StringSplitOptions.None);
                                if (Convert.ToInt32(ewords.Length - 1) >= 1 && my_string.Contains(compare_string) == true)
                                {
                                    compare_number[i] = compare_number[i] + (ewords.Length - 1);
                                    //Console.WriteLine(compare_string + ": " + compare_number[i]);

                                    //////////////////////////////////////////////////////////////////////Create the compare csv file////////////////////
                                    if (System.IO.File.Exists(csvFile) == false)
                                    {
                                        StreamWriter sw1 = new StreamWriter(csvFile, false, Encoding.UTF8);
                                        sw1.WriteLine("Key words, Setting times, Search times, Time");
                                        sw1.Dispose();
                                    }
                                    StreamWriter sw2 = new StreamWriter(csvFile, true);
                                    sw2.Write(compare_string + ",");
                                    sw2.Write(compare_num + ",");
                                    sw2.Write(compare_number[i] + ",");
                                    sw2.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                                    sw2.Close();

                                    ////////////////////////////////////////////////////////////////////////////////////////////////MAIL//////////////////
                                    if (compare_number[i] > compare_num && send_status[i] == false)
                                    {
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Nowvalue", i.ToString());
                                        ini12.INIWrite(MainSettingPath, "LogSearch", "Display" + i, compare_number[i].ToString());
                                        if (ini12.INIRead(MailPath, "Mail Info", "From", "") != ""
                                            && ini12.INIRead(MailPath, "Mail Info", "To", "") != ""
                                            && ini12.INIRead(MainSettingPath, "LogSearch", "Sendmail", "") == "1")
                                        {
                                            FormMail FormMail = new FormMail();
                                            FormMail.logsend();
                                            send_status[i] = true;
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF ON//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "ACcontrol", "") == "1")
                                    {
                                        byte[] val1;
                                        val1 = new byte[2];
                                        val1[0] = 0;

                                        bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("0");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = false;
                                                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                }
                                            }
                                        }

                                        System.Threading.Thread.Sleep(5000);

                                        jSuccess = PL2303_GP0_Enable(hCOM, 1);
                                        if (!jSuccess)
                                        {
                                            Log("GP0 output enable FAILED.");
                                        }
                                        else
                                        {
                                            uint val;
                                            val = (uint)int.Parse("1");
                                            bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                            if (bSuccess)
                                            {
                                                {
                                                    PowerState = true;
                                                    pictureBox_AcPower.Image = Properties.Resources.ON;
                                                }
                                            }
                                        }
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////AC OFF//////////////////
                                    if (compare_number[i] % compare_num == 0
                                        && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1"
                                        && ini12.INIRead(MainSettingPath, "LogSearch", "AC OFF", "") == "1")
                                    {
                                        byte[] val1 = new byte[2];
                                        val1[0] = 0;
                                        uint val = (uint)int.Parse("0");

                                        bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                                        bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                                        bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                                        bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);

                                        PowerState = false;

                                        pictureBox_AcPower.Image = Properties.Resources.OFF;
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SAVE LOG//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Savelog", "") == "1")
                                    {
                                        string fName = "";

                                        // 讀取ini中的路徑
                                        fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                        string t = fName + "\\_SaveLog3_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                        StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                        MYFILE.Write(logC_text);
                                        MYFILE.Close();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////STOP//////////////////
                                    if (compare_number[i] % compare_num == 0 && ini12.INIRead(MainSettingPath, "LogSearch", "Stop", "") == "1")
                                    {
                                        button_Start.PerformClick();
                                    }
                                    ////////////////////////////////////////////////////////////////////////////////////////////////SCHEDULE//////////////////
                                    if (compare_number[i] % compare_num == 0)
                                    {
                                        int keyword_numer = i + 1;
                                        switch (keyword_numer)
                                        {
                                            case 1:
                                                GlobalData.keyword_1 = "true";
                                                break;

                                            case 2:
                                                GlobalData.keyword_2 = "true";
                                                break;

                                            case 3:
                                                GlobalData.keyword_3 = "true";
                                                break;

                                            case 4:
                                                GlobalData.keyword_4 = "true";
                                                break;

                                            case 5:
                                                GlobalData.keyword_5 = "true";
                                                break;

                                            case 6:
                                                GlobalData.keyword_6 = "true";
                                                break;

                                            case 7:
                                                GlobalData.keyword_7 = "true";
                                                break;

                                            case 8:
                                                GlobalData.keyword_8 = "true";
                                                break;

                                            case 9:
                                                GlobalData.keyword_9 = "true";
                                                break;

                                            case 10:
                                                GlobalData.keyword_10 = "true";
                                                break;
                                        }
                                    }
                                }
                            }
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        #endregion

                        else
                        {
                            my_string = my_string + Keyword_SerialPort_E_temp_char;
                        }
                    }
                    else
                    {

                        if ((Keyword_SerialPort_E_temp_char == '\n'))
                        {
                            //textBox3.AppendText(my_string + '\n');
                            my_string = "";
                        }
                        else if ((Keyword_SerialPort_E_temp_char == '\r'))
                        {
                            //textBox3.AppendText(my_string + '\r');
                            my_string = "";
                        }
                        else
                        {
                            my_string = my_string + Keyword_SerialPort_E_temp_char;
                        }
                    }
                }
                Thread.Sleep(500);
            }
        }
        #endregion

        #region -- 迴車換行符號置換 --
        private void ReplaceNewLine(Mod_RS232 port, string columns_serial, string columns_switch)
        {
            List<string> originLineList = new List<string> {"\\r\\n", "\\n\\r", "\\r", "\\n"};
            List<string> newLineList = new List<string> {"\r\n", "\n\r", "\r", "\n"};
            var originAndNewLine = originLineList.Zip(newLineList, (o, n) => new { origin = o, newLine = n });
            foreach (var line in originAndNewLine)
            {
                if (columns_switch.Contains(line.origin))
                {
                    string stringToWrite = columns_serial + columns_switch.Replace(line.origin, line.newLine);
                    port.WriteDataOut(stringToWrite, stringToWrite.Length);
                    return;
                }
            }
        }
        #endregion

        #region -- 跑Schedule的指令集 --
        private void MyRunCamd()
        {
            int sRepeat = 0, stime = 0, SysDelay = 0;

            GlobalData.Loop_Number = 1;
            GlobalData.Break_Out_Schedule = 0;
            GlobalData.Pass_Or_Fail = "PASS";

            label_TestTime_Value.Text = "0d 0h 0m 0s 0ms";
            TestTime = 0;

            for (int l = 0; l <= GlobalData.Schedule_Loop; l++)
            {
                GlobalData.NGValue[l] = 0;
                GlobalData.NGRateValue[l] = 0;
            }

            #region -- 匯出比對結果到CSV & EXCEL --
            if (ini12.INIRead(MainSettingPath, "Record", "CompareChoose", "") == "1" && StartButtonPressed == true)
            {
                string compareFolder = ini12.INIRead(MainSettingPath, "Record", "VideoPath", "") + "\\" + "Schedule" + GlobalData.Schedule_Number + "_Original_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                if (Directory.Exists(compareFolder))
                {

                }
                else
                {
                    Directory.CreateDirectory(compareFolder);
                    ini12.INIWrite(MainSettingPath, "Record", "ComparePath", compareFolder);
                }
                // 匯出csv記錄檔
                string csvFile = ini12.INIRead(MainSettingPath, "Record", "ComparePath", "") + "\\SimilarityReport_" + GlobalData.Schedule_Number + ".csv";
                StreamWriter sw = new StreamWriter(csvFile, false, Encoding.UTF8);
                sw.WriteLine("Target, Source, Similarity, Sub-NG count, NGRate, Result");

                sw.Dispose();
                /*
                                #region Excel function
                                // 匯出excel記錄檔
                                GlobalData.excel_Num = 1;
                                string excelFile = ini12.INIRead(sPath, "Record", "ComparePath", "") + "\\SimilarityReport_" + GlobalData.Schedule_Num;

                                excelApp = new Excel.Application();
                                //excelApp.Visible = true;
                                excelApp.DisplayAlerts = false;
                                excelApp.Workbooks.Add(Type.Missing);
                                wBook = excelApp.Workbooks[1];
                                wBook.Activate();
                                excelstat = true;

                                try
                                {
                                    // 引用第一個工作表
                                    wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                                    // 命名工作表的名稱
                                    wSheet.Name = "全部測試資料";

                                    // 設定工作表焦點
                                    wSheet.Activate();

                                    excelApp.Cells[1, 1] = "All Data";

                                    // 設定第1列資料
                                    excelApp.Cells[1, 1] = "Target";
                                    excelApp.Cells[1, 2] = "Source";
                                    excelApp.Cells[1, 3] = "Similarity";
                                    excelApp.Cells[1, 4] = "Sub-NG count";
                                    excelApp.Cells[1, 5] = "NGRate";
                                    excelApp.Cells[1, 6] = "Result";
                                    // 設定第1列顏色
                                    wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 6]];
                                    wRange.Select();
                                    wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                                    wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
                                    wRange.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
                                }
                                #endregion
                */
            }
            #endregion

            for (int j = 1; j < GlobalData.Schedule_Loop + 1; j++)
            {
                GlobalData.caption_Num = 0;
                UpdateUI(j.ToString(), label_LoopNumber_Value);
                GlobalData.label_LoopNumber = j.ToString();
                ini12.INIWrite(MailPath, "Data Info", "CreateTime", string.Format("{0:R}", DateTime.Now));

                lock (this)
                {
                    for (GlobalData.Scheduler_Row = 0; GlobalData.Scheduler_Row < DataGridView_Schedule.Rows.Count - 1; GlobalData.Scheduler_Row++)
                    {
                        //Schedule All columns list
                        string columns_command = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[0].Value.ToString().Trim();
                        string columns_times = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[1].Value.ToString().Trim();
                        string columns_interval = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[2].Value.ToString().Trim();
                        string columns_comport = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[3].Value.ToString().Trim();
                        string columns_function = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[4].Value.ToString().Trim();
                        string columns_subFunction = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[5].Value.ToString().Trim();
                        string columns_serial = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[6].Value.ToString().Trim();
                        string columns_switch = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[7].Value.ToString().Trim();
                        string columns_wait = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[8].Value.ToString().Trim();
                        string columns_remark = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[9].Value.ToString().Trim();

                        IO_INPUT();                 //先讀取IO值，避免schedule第一行放IO CMD會出錯//
                        if (GlobalData.m_Arduino_Port.IsOpen() == true && GlobalData.Arduino_outputFlag == false)
                            Arduino_IO_INPUT();         //先讀取Arduino_IO值，避免schedule第一行放IO CMD會出錯//

                        GlobalData.Schedule_Step = GlobalData.Scheduler_Row;
                        if (StartButtonPressed == false)
                        {
                            j = GlobalData.Schedule_Loop;
                            UpdateUI(j.ToString(), label_LoopNumber_Value);
                            GlobalData.label_LoopNumber = j.ToString();
                            break;
                        }

                        Schedule_Time();
                        if (columns_wait != "")
                        {
                            if (columns_wait.Contains('m'))
                            {
                                //Console.WriteLine("Datagridview highlight.");
                                GridUI(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview highlight//
                                //Console.WriteLine("Datagridview scollbar.");
                                Gridscroll(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview scollbar//
                            }
                            else
                            {
                                if (int.Parse(columns_wait) > 500)  //DataGridView UI update 
                                {
                                    //Console.WriteLine("Datagridview highlight.");
                                    GridUI(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview highlight//
                                    //Console.WriteLine("Datagridview scollbar.");
                                    Gridscroll(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview scollbar//
                                }
                            }
                        }

                        if (columns_times != "" && int.TryParse(columns_times, out stime) == true)
                            stime = int.Parse(columns_times); // 次數
                        else
                            stime = 1;

                        if (columns_interval != "" && int.TryParse(columns_interval, out sRepeat) == true)
                            sRepeat = int.Parse(columns_interval); // 停止時間
                        else
                            sRepeat = 0;

                        if (columns_wait != "" && int.TryParse(columns_wait, out SysDelay) == true && columns_wait.Contains('m') == false)
                            SysDelay = int.Parse(columns_wait); // 指令停止時間(毫秒)
                        else if (columns_wait != "" && columns_wait.Contains('m') == true)
                            SysDelay = int.Parse(columns_wait.Replace('m', ' ').Trim()) * 60000; // 指令停止時間(分)
                        else
                            SysDelay = 0;

                        #region -- Record Schedule --
                        string delimiter_recordSch = ",";
                        string Schedule_log = "";
                        DateTime.Now.ToShortTimeString();
                        DateTime sch_dt = DateTime.Now;

                        log.Debug("Record Schedule");
                        Schedule_log = columns_command;
                        try
                        {
                            for (int i = 1; i < 10; i++)
                            {
                                Schedule_log = Schedule_log + delimiter_recordSch + DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[i].Value.ToString();
                            }
                        }
                        catch (Exception Ex)
                        {
                            MessageBox.Show(Ex.Message.ToString(), "The schedule length incorrect!");
                        }

                        string sch_log_text = "[Schedule] [" + sch_dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Schedule_log + "\r\n";
                        logDumpping.LogCat(ref logA_text, sch_log_text);        //log_process("A", sch_log_text);
                        logDumpping.LogCat(ref logB_text, sch_log_text);        //log_process("B", sch_log_text);
                        logDumpping.LogCat(ref logC_text, sch_log_text);        //log_process("C", sch_log_text);
                        logDumpping.LogCat(ref logD_text, sch_log_text);        //log_process("D", sch_log_text);
                        logDumpping.LogCat(ref logE_text, sch_log_text);        //log_process("E", sch_log_text);
                        logDumpping.LogCat(ref logAll_text, sch_log_text);        //log_process("All", sch_log_text);
                        logDumpping.LogCat(ref arduino_text, sch_log_text);        //log_process("Arduino", sch_log_text);
                        logDumpping.LogCat(ref canbus_text, sch_log_text);        //log_process("Canbus", sch_log_text);
                        logDumpping.LogCat(ref kline_text, sch_log_text);        //log_process("KlinePort", sch_log_text);
                        textBox_serial.AppendText(sch_log_text);
                        #endregion

                        #region -- _cmd --
                        if (columns_command == "_cmd")
                        {
                            #region -- AC SWITCH OLD --
                            if (columns_switch == "_on")
                            {
                                Console.WriteLine("AC SWITCH OLD: _on");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP0_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = true;
                                                pictureBox_AcPower.Image = Properties.Resources.ON;
                                                label_Command.Text = "AC ON";
                                            }
                                        }
                                    }
                                    if (PL2303_GP1_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = true;
                                                pictureBox_AcPower.Image = Properties.Resources.ON;
                                                label_Command.Text = "AC ON";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (columns_switch == "_off")
                            {
                                Console.WriteLine("AC SWITCH OLD: _off");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP0_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = false;
                                                pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                label_Command.Text = "AC OFF";
                                            }
                                        }
                                    }
                                    if (PL2303_GP1_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = false;
                                                pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                label_Command.Text = "AC OFF";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            #endregion

                            #region -- AC SWITCH --
                            if (columns_switch == "_AC1_ON")
                            {
                                Console.WriteLine("AC SWITCH: _AC1_ON");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP0_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = true;
                                                pictureBox_AcPower.Image = Properties.Resources.ON;
                                                label_Command.Text = "AC1 => POWER ON";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (columns_switch == "_AC1_OFF")
                            {
                                Console.WriteLine("AC SWITCH: _AC1_OFF");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP0_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = false;
                                                pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                label_Command.Text = "AC1 => POWER OFF";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }

                            if (columns_switch == "_AC2_ON")
                            {
                                Console.WriteLine("AC SWITCH: _AC2_ON");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP1_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = true;
                                                pictureBox_AcPower.Image = Properties.Resources.ON;
                                                label_Command.Text = "AC2 => POWER ON";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (columns_switch == "_AC2_OFF")
                            {
                                Console.WriteLine("AC SWITCH: _AC2_OFF");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP1_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                                        if (bSuccess)
                                        {
                                            {
                                                PowerState = false;
                                                pictureBox_AcPower.Image = Properties.Resources.OFF;
                                                label_Command.Text = "AC2 => POWER OFF";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Autobox Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            #endregion

                            #region -- USB SWITCH --
                            if (columns_switch == "_USB1_DUT")
                            {
                                Console.WriteLine("USB SWITCH: _USB1_DUT");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP2_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP2_SetValue(hCOM, val);
                                        if (bSuccess == true)
                                        {
                                            {
                                                USBState = false;
                                                label_Command.Text = "USB1 => DUT";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else if (columns_switch == "_USB1_PC")
                            {
                                Console.WriteLine("USB SWITCH: _USB1_PC");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP2_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP2_SetValue(hCOM, val);
                                        if (bSuccess == true)
                                        {
                                            {
                                                USBState = true;
                                                label_Command.Text = "USB1 => PC";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }

                            if (columns_switch == "_USB2_DUT")
                            {
                                Console.WriteLine("USB SWITCH: _USB2_DUT");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP3_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("1");
                                        bool bSuccess = PL2303_GP3_SetValue(hCOM, val);
                                        if (bSuccess == true)
                                        {
                                            {
                                                USBState = false;
                                                label_Command.Text = "USB2 => DUT";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else if (columns_switch == "_USB2_PC")
                            {
                                Console.WriteLine("USB SWITCH: _USB2_PC");
                                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                {
                                    if (PL2303_GP3_Enable(hCOM, 1) == true)
                                    {
                                        uint val = (uint)int.Parse("0");
                                        bool bSuccess = PL2303_GP3_SetValue(hCOM, val);
                                        if (bSuccess == true)
                                        {
                                            {
                                                USBState = true;
                                                label_Command.Text = "USB2 => PC";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please connect an AutoKit!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region -- 拍照 --
                        else if (columns_command == "_shot")
                        {
                            log.Debug("Take Picture: _shot_start");
                            if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                            {
                                GlobalData.caption_Num++;
                                if (GlobalData.Loop_Number == 1)
                                    GlobalData.caption_Sum = GlobalData.caption_Num;
                                Jes();
                                label_Command.Text = "Take Picture";
                            }
                            else
                            {
                                button_Start.PerformClick();
                                MessageBox.Show("Camera is not connected!\r\nPlease go to Settings to reload the device list.", "Connection Error");
                                //setStyle();
                            }
                            log.Debug("Take Picture: _shot_stop");
                        }
                        #endregion

                        #region -- 錄影 --
                        else if (columns_command == "_rec_start")
                        {
                            log.Debug("Take Record: _rec_start");
                            if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                            {
                                if (GlobalData.VideoRecording == false && columns_serial == "")
                                {
                                    Mysvideo(); // 開新檔
                                    GlobalData.VideoRecording = true;
                                    Thread oThreadC = new Thread(new ThreadStart(MySrtCamd));
                                    oThreadC.Start();
                                }
                                else if (GlobalData.VideoRecording == false && columns_serial == "wmv")
                                {
                                    Mywmvideo(); // 開新檔
                                    GlobalData.VideoRecording = true;
                                }
                                label_Command.Text = "Start Recording";
                            }
                            else
                            {
                                MessageBox.Show("Camera is not connected", "Camera Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                button_Start.PerformClick();
                            }
                        }

                        else if (columns_command == "_rec_stop")
                        {
                            log.Debug("Take Record: _rec_stop");
                            if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                            {
                                if (GlobalData.VideoRecording == true)       //判斷是不是正在錄影
                                {
                                    GlobalData.VideoRecording = false;
                                    Mysstop();      //先將先前的關掉
                                }
                                label_Command.Text = "Stop Recording";
                            }
                            else
                            {
                                MessageBox.Show("Camera is not connected", "Camera Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                button_Start.PerformClick();
                            }
                        }
                        #endregion

                        #region -- COM PORT --
                        /*
                        else if (columns_command == "_log1")
                        {
                            if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                            {
                                switch (columns_serial)
                                {
                                    case "_clear":
                                        textBox1 = string.empty; //清除textbox1
                                        break;

                                    case "_save":
                                        Rs232save(); //存檔rs232
                                        break;

                                    default:
                                        //byte[] data = Encoding.Unicode.GetBytes(DataGridView1.Rows[GlobalData.Scheduler_Row].Cells[5].Value.ToString());
                                        // string str = Convert.ToString(data);
                                        serialPort1.WriteLine(columns_serial); //發送數據 Rs232
                                        DateTime dt = DateTime.Now;
                                        string text = "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\n";
                                        textBox1.AppendText(text);
                                        break;
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }
                        }

                        else if (columns_command == "_log2")
                        {
                            if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1")
                            {
                                switch (columns_serial)
                                {
                                    case "_clear":
                                        textBox2.Clear(); //清除textbox2
                                        break;

                                    case "_save":
                                        ExtRs232save(); //存檔rs232
                                        break;

                                    default:
                                        //byte[] data = Encoding.Unicode.GetBytes(DataGridView1.Rows[GlobalData.Scheduler_Row].Cells[5].Value.ToString());
                                        // string str = Convert.ToString(data);
                                        serialPort2.WriteLine(columns_serial); //發送數據 Rs232
                                        DateTime dt = DateTime.Now;
                                        string text = "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\n";
                                        textBox2.AppendText(text);
                                        break;
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }
                        }

                        else if (columns_command == "_log3")
                        {
                            if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1")
                            {
                                switch (columns_serial)
                                {
                                    case "_clear":
                                        textBox3.Clear(); //清除textbox3
                                        break;

                                    case "_save":
                                        TriRs232save(); //存檔rs232
                                        break;

                                    default:
                                        //byte[] data = Encoding.Unicode.GetBytes(DataGridView1.Rows[GlobalData.Scheduler_Row].Cells[5].Value.ToString());
                                        // string str = Convert.ToString(data);
                                        serialPort3.WriteLine(columns_serial); //發送數據 Rs232
                                        DateTime dt = DateTime.Now;
                                        string text = "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\n";
                                        textBox3.AppendText(text);
                                        break;
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }
                        }*/
                        #endregion

                        #region -- Ascii --
                        else if (columns_command == "_ascii")
                        {
                            //if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" && columns_comport == "A")
							if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                            {
                                log.Debug("Ascii Log: _PortA");
                                if (columns_serial == "_save")
                                {
                                    Serialportsave("A"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logA_text = string.Empty; //清除logA_text
                                }
                                else if (columns_serial != "" || columns_switch != "")
                                {
                                    ReplaceNewLine(GlobalData.m_SerialPort_A, columns_serial, columns_switch);
                                }
                                else if (columns_serial == "" && columns_switch == "")
                                {
                                    MessageBox.Show("Ascii command is fail, please check the format.");
                                }
                                /*
                                DateTime dt = DateTime.Now;
                                string text = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\n\r";
                                textBox_serial.AppendText(dataValue);
                                log_process("A", dataValue);
                                log_process("All", dataValue);
                                */
                            }

                            if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                            {
                                log.Debug("Ascii Log: _PortB");
                                if (columns_serial == "_save")
                                {
                                    Serialportsave("B"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logB_text = string.Empty; //清除logB_text
                                }
                                else if (columns_serial != "" || columns_switch != "")
                                {
                                    ReplaceNewLine(GlobalData.m_SerialPort_B, columns_serial, columns_switch);
                                }
                                else if (columns_serial == "" && columns_switch == "")
                                {
                                    MessageBox.Show("Ascii command is fail, please check the format.");
                                }
                                /*
                                DateTime dt = DateTime.Now;
                                string text = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                log_process("B", dataValue);
                                log_process("All", dataValue);
                                */
                            }

                            if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                            {
                                log.Debug("Ascii Log: _PortC");
                                if (columns_serial == "_save")
                                {
                                    Serialportsave("C"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logC_text = string.Empty; //清除logC_text
                                }
                                else if (columns_serial != "" || columns_switch != "")
                                {
                                    ReplaceNewLine(GlobalData.m_SerialPort_C, columns_serial, columns_switch);
                                }
                                else if (columns_serial == "" && columns_switch == "")
                                {
                                    MessageBox.Show("Ascii command is fail, please check the format.");
                                }
                                /*
                                DateTime dt = DateTime.Now;
                                string text = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                log_process("C", dataValue);
                                log_process("All", dataValue);
                                */
                            }

                            if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                            {
                                log.Debug("Ascii Log: _PortD");
                                if (columns_serial == "_save")
                                {
                                    Serialportsave("D"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logD_text = string.Empty; //清除logD_text
                                }
                                else if (columns_serial != "" || columns_switch != "")
                                {
                                    ReplaceNewLine(GlobalData.m_SerialPort_D, columns_serial, columns_switch);
                                }
                                else if (columns_serial == "" && columns_switch == "")
                                {
                                    MessageBox.Show("Ascii command is fail, please check the format.");
                                }
                                /*
                                DateTime dt = DateTime.Now;
                                string text = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                log_process("D", dataValue);
                                log_process("All", dataValue);
                                */
                            }

                            if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                            {
                                log.Debug("Ascii Log: _PortE");
                                if (columns_serial == "_save")
                                {
                                    Serialportsave("E"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logE_text = string.Empty; //清除logE_text
                                }
                                else if (columns_serial != "" || columns_switch != "")
                                {
                                    ReplaceNewLine(GlobalData.m_SerialPort_E, columns_serial, columns_switch);
                                }
                                else if (columns_serial == "" && columns_switch == "")
                                {
                                    MessageBox.Show("Ascii command is fail, please check the format.");
                                }
                                /*
                                DateTime dt = DateTime.Now;
                                string text = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                log_process("E", dataValue);
                                log_process("All", dataValue);
                                */
                            }

                            if (columns_comport == "ALL")
                            {
                                log.Debug("Ascii Log: _All");
                                string[] serial_content = columns_serial.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                                string[] switch_content = columns_switch.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                                if (columns_serial == "_save")
                                {
                                    Serialportsave("All"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logAll_text = string.Empty; //清除logAll_text
                                }

                                if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "ALL" && serial_content[0] != "" && switch_content[0] != "")
                                {
                                    //不同的ReplaceNewLine Function
                                    //ReplaceNewLine(GlobalData.m_SerialPort_A, serial_content[0], switch_content[0]);
                                    logDumpping.ReplaceNewLine(serialPortA, serial_content[0], switch_content[0]);
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    //textBox_serial.AppendText(dataValue);
                                    logDumpping.LogCat(ref logA_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                    //log_process("A", dataValue);
                                    //log_process("All", dataValue);
                                }
                                if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "ALL" && serial_content[1] != "" && switch_content[1] != "")
                                {
                                    logDumpping.ReplaceNewLine(serialPortB, serial_content[1], switch_content[1]);
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    
                                    logDumpping.LogCat(ref logB_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "ALL" && serial_content[2] != "" && switch_content[2] != "")
                                {
                                    logDumpping.ReplaceNewLine(serialPortC, serial_content[2], switch_content[2]);
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    
                                    logDumpping.LogCat(logC_text, dataValue);
                                    logDumpping.LogCat(logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "ALL" && serial_content[3] != "" && switch_content[3] != "")
                                {
                                    logDumpping.ReplaceNewLine(serialPortD, serial_content[3], switch_content[3]);
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    
                                    logDumpping.LogCat(logD_text, dataValue);
                                    logDumpping.LogCat(logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "ALL" && serial_content[4] != "" && switch_content[4] != "")
                                {
                                    logDumpping.ReplaceNewLine(serialPortE, serial_content[4], switch_content[4]);
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    
                                    logDumpping.LogCat(logE_text, dataValue);
                                    logDumpping.LogCat(logAll_text, dataValue);
                                }
                            }

                            label_Command.Text = "(" + columns_command + ") " + columns_serial;
                        }
                        #endregion

                        #region -- Execute --
                        else if (columns_command == "_Execute")
                        {
                            //If voltage matched the expected value
                            if (PowerSupplyCheck)
                            {
                                IO_CMD();
                            }
                            PowerSupplyCheck = false;

                            if (columns_serial == "_pause")
                            {
                                foreach (Temperature_Data item in temperatureList)
                                {
                                    item.temperaturePause = true;
                                }
                            }
                            else if (columns_serial == "_shot")
                            {
                                foreach (Temperature_Data item in temperatureList)
                                {
                                    item.temperatureShot = true;
                                }
                            }
                            else if (columns_comport == "A" || columns_comport == "B" || columns_comport == "C" || columns_comport == "D" || columns_comport == "E")
                            {
                                foreach (Temperature_Data item in temperatureList)
                                {
                                    item.temperaturePort = columns_comport;
                                    item.temperatureLog = columns_serial;
                                    item.temperatureNewline = columns_switch;
                                }
                            }
                        }

                        #endregion

                        #region -- Condition_AND --
                        else if (columns_command == "_Condition_AND")
                        {
                            //if (columns_command.Substring(13) == "1")
                            //{
                            if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" ||
                                ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" ||
                                ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1" ||
                                ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1" ||
                                ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1")
                            {
                                if (columns_function == "start")
                                {
                                    ifStatementFlag = true;
                                    expectedVoltage = "";
                                    if (columns_serial != "")
                                    {
                                        columns_serial.Replace(" ", "");
                                        if (columns_serial.Contains("chamber_temp"))
                                        {
                                            timer_Chamber.Enabled = true;
                                            ChamberIsFound = true;
                                            Temperature_Data.initialTemperature = Int16.Parse(columns_serial.Substring(columns_serial.IndexOf("=") + 1, columns_serial.IndexOf("~") - columns_serial.IndexOf("=") - 1));
                                            Temperature_Data.finalTemperature = Int16.Parse(columns_serial.Substring(columns_serial.IndexOf("~") + 1, columns_serial.IndexOf("/") - columns_serial.IndexOf("~") - 1));
                                            if (columns_serial.Contains("/-"))
                                            {
                                                Temperature_Data.addTemperature = float.Parse("-" + columns_serial.Substring(columns_serial.IndexOf("-") + 1));
                                            }
                                            else
                                            {
                                                Temperature_Data.addTemperature = float.Parse(columns_serial.Substring(columns_serial.IndexOf("+") + 1));
                                            }
                                        }
                                        else if (columns_serial.Contains("PowerSupply_Voltage"))
                                        {
                                            PowerSupplyIsFound = true;
                                            expectedVoltage = columns_serial.Substring(columns_serial.IndexOf("=") + 1);

                                            string powerCommand = "MEASure1:ALL?"; //Read Power Supply information
                                            ReplaceNewLine(GlobalData.m_SerialPort_A, powerCommand, columns_switch);

                                            //Append Power Supply command to log
                                            DateTime dt = DateTime.Now;
                                            PowerSupplyCommandLog = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + powerCommand + "\r\n";
                                            logA_text = string.Concat(logA_text, PowerSupplyCommandLog);
                                        }
                                        else if (columns_serial.Contains("Temperature"))
                                        {
                                            try
                                            {
                                                //	lt：less than 小於
                                                //	le：less than or equal to 小於等於
                                                //	eq：equal to 等於
                                                //	ne：not equal to 不等於
                                                //	ge：greater than or equal to 大於等於
                                                //	gt：greater than 大於

                                                TemperatureIsFound = true;
                                                int symbel_equal_7e = columns_serial.IndexOf("~");
                                                int symbel_equal_28 = columns_serial.IndexOf("(");
                                                int symbel_equal_29 = columns_serial.IndexOf(")");
                                                int symbel_equal_6d29 = columns_serial.IndexOf("m)");
                                                int symbel_equal_3d = columns_serial.IndexOf("=");
                                                int symbel_equal_3c = columns_serial.IndexOf("<");
                                                int symbel_equal_3e = columns_serial.IndexOf(">");
                                                int symbel_equal_3d3d = columns_serial.IndexOf("==");
                                                int symbel_equal_3c3e = columns_serial.IndexOf("<>");
                                                int symbel_equal_3c3d = columns_serial.IndexOf("<=");
                                                int symbel_equal_3e3d = columns_serial.IndexOf(">=");
                                                int duringTimeInt = 0;
                                                int parameter_equal_Temperature = columns_serial.IndexOf("Temperature");

                                                if (columns_serial.Contains("~") && columns_serial.Contains("<>") && columns_serial.Contains("/") == false)
                                                {
                                                    MaxTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3e + 1, symbel_equal_7e - symbel_equal_3c3e - 1));
                                                    MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_7e + 1, symbel_equal_28 - symbel_equal_7e - 1));
                                                }
                                                else if (columns_serial.Contains("~") && columns_serial.Contains("=") && columns_serial.Contains("/") == false)
                                                {
                                                    MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3d + 1, symbel_equal_7e - symbel_equal_3d - 1));
                                                    MaxTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_7e + 1, symbel_equal_28 - symbel_equal_7e - 1));
                                                }
                                                else if (columns_serial.Contains("~") == false && columns_serial.Contains("/") == false)
                                                {
                                                    if (columns_serial.Contains("<"))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c + 1, symbel_equal_28 - symbel_equal_3c - 1));
                                                    else if (columns_serial.Contains("<="))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3d + 1, symbel_equal_28 - symbel_equal_3c3d - 1));
                                                    else if (columns_serial.Contains("=="))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3d3d + 1, symbel_equal_28 - symbel_equal_3d3d - 1));
                                                    else if (columns_serial.Contains("<>"))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3e + 1, symbel_equal_28 - symbel_equal_3c3e - 1));
                                                    else if (columns_serial.Contains(">"))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3e + 1, symbel_equal_28 - symbel_equal_3e - 1));
                                                    else if (columns_serial.Contains(">="))
                                                        MinTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3e3d + 1, symbel_equal_28 - symbel_equal_3e3d - 1));
                                                }
                                                
                                                string string_temperatureChannel = columns_serial.Substring(parameter_equal_Temperature + 11, symbel_equal_3d - parameter_equal_Temperature - 11);

                                                if (columns_serial.Contains("m)"))
                                                    duringTimeInt = Int16.Parse(columns_serial.Substring(symbel_equal_28 + 1, symbel_equal_6d29 - symbel_equal_28 - 1)) * 60000;
                                                else
                                                    duringTimeInt = Int16.Parse(columns_serial.Substring(symbel_equal_28 + 1, symbel_equal_29 - symbel_equal_28 - 1));

                                                byte temperatureChannel = Convert.ToByte(int.Parse(string_temperatureChannel) + 48);

                                                if (duringTimeInt > 0)
                                                {
                                                    // Create a timer and set a two second interval.
                                                    timer_duringShot.Interval = duringTimeInt;

                                                    // Start the timer
                                                    timer_duringShot.Start();
                                                }
                                            }
                                            catch (Exception Ex)
                                            {
                                                MessageBox.Show(Ex.Message.ToString(), "Temperature data parameter error!");
                                            }
                                        }
                                    }
                                }
                                else if (columns_function == "end")
                                {
                                    ifStatementFlag = false;

                                    PowerSupplyCheck = false;
                                    PowerSupplyIsFound = false;

                                    ChamberCheck = false;
                                    ChamberIsFound = false;
                                    timer_Chamber.Enabled = false;

                                    chamberTimer_IsTick = false;
                                    timer_duringShot.Stop();

                                    foreach (Temperature_Data item in temperatureList)
                                    {
                                        item.temperaturePause = false;
                                        item.temperatureShot = false;
                                    }

                                    StartButtonFlag = false;
                                }
                            }
                            //}
                        }
                        #endregion

                        #region -- Condition_OR --
                        else if (columns_command == "_Condition_OR")
                        {
                            //if (columns_command.Substring(13) == "1")
                            //{
                                if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" ||
                                    ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" ||
                                    ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1" ||
                                    ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1" ||
                                    ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1")
                                {
                                    if (columns_function == "start")
                                    {
                                        ifStatementFlag = true;
                                        expectedVoltage = "";
                                        if (columns_serial != "")
                                        {
                                            columns_serial.Replace(" ", "");
                                            if (columns_serial.Contains("chamber_temp"))
                                            {
                                                timer_Chamber.Enabled = true;
                                                ChamberIsFound = true;
                                                Temperature_Data.initialTemperature = Int16.Parse(columns_serial.Substring(columns_serial.IndexOf("=") + 1, columns_serial.IndexOf("~") - columns_serial.IndexOf("=") - 1));
                                                Temperature_Data.finalTemperature = Int16.Parse(columns_serial.Substring(columns_serial.IndexOf("~") + 1, columns_serial.IndexOf("/") - columns_serial.IndexOf("~") - 1));
                                                if (columns_serial.Contains("/-"))
                                                {
                                                    Temperature_Data.addTemperature = float.Parse("-" + columns_serial.Substring(columns_serial.IndexOf("-") + 1));
                                                }
                                                else
                                                {
                                                    Temperature_Data.addTemperature = float.Parse(columns_serial.Substring(columns_serial.IndexOf("+") + 1));
                                                }
                                            }
                                            else if (columns_serial.Contains("PowerSupply_Voltage"))
                                            {
                                                PowerSupplyIsFound = true;
                                                expectedVoltage = columns_serial.Substring(columns_serial.IndexOf("=") + 1);

                                                string powerCommand = "MEASure1:ALL?"; //Read Power Supply information
                                                ReplaceNewLine(GlobalData.m_SerialPort_A, powerCommand, columns_switch);

                                                //Append Power Supply command to log
                                                DateTime dt = DateTime.Now;
                                                PowerSupplyCommandLog = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + powerCommand + "\r\n";
                                                logA_text = string.Concat(logA_text, PowerSupplyCommandLog);
                                            }
                                            else if (columns_serial.Contains("Temperature"))
                                            {
                                                try
                                                {
                                                    //	lt：less than 小於
                                                    //	le：less than or equal to 小於等於
                                                    //	eq：equal to 等於
                                                    //	ne：not equal to 不等於
                                                    //	ge：greater than or equal to 大於等於
                                                    //	gt：greater than 大於

                                                    TemperatureIsFound = true;
                                                    temperatureList.Clear();
                                                    int symbel_equal_3d = columns_serial.IndexOf("=");
                                                    int symbel_equal_7e = columns_serial.IndexOf("~");
                                                    int symbel_equal_2f = columns_serial.IndexOf("/");
                                                    int symbel_equal_28 = columns_serial.IndexOf("(");
                                                    int symbel_equal_29 = columns_serial.IndexOf(")");
                                                    int symbel_equal_6d29 = columns_serial.IndexOf("m)");
                                                    int symbel_equal_3c = columns_serial.IndexOf("<");
                                                    int symbel_equal_3e = columns_serial.IndexOf(">");
                                                    int symbel_equal_3d3d = columns_serial.IndexOf("==");
                                                    int symbel_equal_3c3e = columns_serial.IndexOf("<>");
                                                    int symbel_equal_3c3d = columns_serial.IndexOf("<=");
                                                    int symbel_equal_3e3d = columns_serial.IndexOf(">=");
                                                    int duringTimeInt = 0;
                                                    int parameter_equal_Temperature = columns_serial.IndexOf("Temperature");
                                                    string initialTemperature = "", finalTemperature = "", addTemperature = "", symbel_equal_Math = "";
                                                    string temperatureChannel = columns_serial.Substring(parameter_equal_Temperature + 11, symbel_equal_3d - parameter_equal_Temperature - 11);

                                                    if (columns_serial.Contains("~") && columns_serial.Contains("/"))
                                                    {
                                                        initialTemperature = string.Format("{0:0.00}", columns_serial.Substring(symbel_equal_3d + 1, symbel_equal_7e - symbel_equal_3d - 1));
                                                        finalTemperature = string.Format("{0:0.00}", columns_serial.Substring(symbel_equal_7e + 1, symbel_equal_2f - symbel_equal_7e - 1));
                                                        addTemperature = string.Format("{0:0.00}", columns_serial.Substring(symbel_equal_2f + 1, symbel_equal_28 - symbel_equal_2f - 1));
                                                    }
                                                    else if (columns_serial.Contains("~") && columns_serial.Contains("<>") && columns_serial.Contains("/") == false)
                                                    {
                                                        initialTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3e + 1, symbel_equal_7e - symbel_equal_3c3e - 1));
                                                        finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_7e + 1, symbel_equal_28 - symbel_equal_7e - 1));
                                                        symbel_equal_Math = "<>";
                                                    }
                                                    else if (columns_serial.Contains("~") && columns_serial.Contains("=") && columns_serial.Contains("/") == false)
                                                    {
                                                        finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3d + 1, symbel_equal_7e - symbel_equal_3d - 1));
                                                        initialTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_7e + 1, symbel_equal_28 - symbel_equal_7e - 1));
                                                        symbel_equal_Math = "==";
                                                    }
                                                    else if (columns_serial.Contains("~") == false && columns_serial.Contains("/") == false)
                                                    {
                                                        if (columns_serial.Contains("<"))
                                                        {
                                                            symbel_equal_Math = "<";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c + 1, symbel_equal_28 - symbel_equal_3c - 1));
                                                        }
                                                        else if (columns_serial.Contains("<="))
                                                        {
                                                            symbel_equal_Math = "<=";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3d + 1, symbel_equal_28 - symbel_equal_3c3d - 1));
                                                        }
                                                        else if (columns_serial.Contains("=="))
                                                        {
                                                            symbel_equal_Math = "==";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3d3d + 1, symbel_equal_28 - symbel_equal_3d3d - 1));
                                                        }
                                                        else if (columns_serial.Contains("<>"))
                                                        {
                                                            symbel_equal_Math = "<>";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3c3e + 1, symbel_equal_28 - symbel_equal_3c3e - 1));
                                                        }
                                                        else if (columns_serial.Contains(">"))
                                                        {
                                                            symbel_equal_Math = ">";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3e + 1, symbel_equal_28 - symbel_equal_3e - 1));
                                                        }
                                                        else if (columns_serial.Contains(">="))
                                                        {
                                                            symbel_equal_Math = ">=";
                                                            finalTemperature = string.Format("{0:0.0}", columns_serial.Substring(symbel_equal_3e3d + 1, symbel_equal_28 - symbel_equal_3e3d - 1));
                                                        }
                                                    }

                                                    if (columns_serial.Contains("m)"))
                                                        duringTimeInt = Int16.Parse(columns_serial.Substring(symbel_equal_28 + 1, symbel_equal_6d29 - symbel_equal_28 - 1)) * 60000;
                                                    else
                                                        duringTimeInt = Int16.Parse(columns_serial.Substring(symbel_equal_28 + 1, symbel_equal_29 - symbel_equal_28 - 1));

                                                    if (columns_serial.Contains("~"))
                                                        Temperature_Data.initialTemperature = float.Parse(initialTemperature);
                                                    Temperature_Data.finalTemperature = float.Parse(finalTemperature);
                                                    Temperature_Data.temperatureChannel = Convert.ToByte(int.Parse(temperatureChannel) + 48);
                                                    if (columns_serial.Contains("~") && columns_serial.Contains("/"))
                                                    {
                                                        Temperature_Data.addTemperature = float.Parse(addTemperature);
                                                    }
                                                    float addTemperatureInt = Temperature_Data.addTemperature;

                                                    if (duringTimeInt > 0)
                                                    {
                                                        // Create a timer and set a two second interval.
                                                        timer_duringShot.Interval = duringTimeInt;

                                                        // Start the timer
                                                        timer_duringShot.Start();
                                                    }

                                                    if (addTemperatureInt < 0)
                                                    {
                                                        for (float i = Temperature_Data.initialTemperature; i >= Temperature_Data.finalTemperature; i += addTemperatureInt)
                                                        {
                                                            double conditionList = Convert.ToDouble(string.Format("{0:0.0}", i));
                                                            temperatureList.Add(new Temperature_Data(conditionList, false, false, "", "", ""));
                                                        }
                                                    }
                                                    else if (addTemperatureInt >= 0)
                                                    {
                                                        for (float i = Temperature_Data.initialTemperature; i <= Temperature_Data.finalTemperature; i += addTemperatureInt)
                                                        {
                                                            double conditionList = Convert.ToDouble(string.Format("{0:0.0}", i));
                                                            temperatureList.Add(new Temperature_Data(conditionList, false, false, "", "", ""));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        switch (symbel_equal_Math)
                                                        {
                                                            case "<":

                                                                break;
                                                            case "<=":

                                                                break;
                                                            case "==":

                                                                break;
                                                            case "<>":

                                                                break;
                                                            case ">":

                                                                break;
                                                            case "=>":

                                                                break;
                                                        }
                                                    }
                                                }
                                                catch(Exception Ex)
                                                {
                                                    MessageBox.Show(Ex.Message.ToString(), "Temperature data parameter error!");
                                                }
                                            }
                                        }
                                    }
                                    else if (columns_function == "end")
                                    {
                                        ifStatementFlag = false;

                                        PowerSupplyCheck = false;
                                        PowerSupplyIsFound = false;

                                        ChamberCheck = false;
                                        ChamberIsFound = false;
                                        timer_Chamber.Enabled = false;

                                        chamberTimer_IsTick = false;
                                        timer_duringShot.Stop();

                                        temperatureList.Clear();

                                        StartButtonFlag = false;
                                    }
                                }
                            //}
                        }
                        #endregion

                        #region -- Hex --
                        else if (columns_command == "_HEX")
                        {
                            Algorithm algorithm = new Algorithm();
                            string Outputstring = "";
                            //if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" && columns_comport == "A")
                            if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                            {
                                log.Debug("Hex Log: _PortA");
                                if (columns_serial == "_save")
                                {
                                    //Serialportsave("A"); //存檔rs232
                                    logDumpping.LogDumpToFile(serialPortConfig_A, GlobalData.portConfigGroup_A.portName, ref logA_text);
                                    Console.WriteLine("[YFC]HEX-logA_text: " + logA_text);
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logA_text = string.Empty; //清除logA_text
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                    columns_serial != "" && columns_function == "CRC16_Modbus")
                                {
                                    string original_data = columns_serial;
                                    string crc16_data = Crc16.PID_CRC16(original_data);
                                    Outputstring = original_data + crc16_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    //serialPortA.WriteDataOut(Outputbytes, Outputbytes.Length);
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length);
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "XOR8")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length);
                                    //serialPortA.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "MOD256")
                                {
                                    string orginal_data = columns_serial;
                                    byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(orginal_data);
                                    Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "")
                                {
                                    string hexValues = columns_serial;
                                    Outputstring = hexValues;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    //serialPortA.WriteDataOut(Outputbytes, Outputbytes.Length);
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length);
                                }

                                DateTime dt = DateTime.Now;
                                string dataValue = "[" + serialPortConfig_A + "(" + GlobalData.portConfigGroup_A.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
								
                                logDumpping.LogCat(ref logA_text, dataValue);
                                logDumpping.LogCat(ref logAll_text, dataValue);

                                //log_process("A", dataValue);
                                //log_process("All", dataValue);
                            }

                            if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                            {
                                log.Debug("Hex Log: _PortB");
                                if (columns_serial == "_save")
                                {
                                    //Serialportsave("B"); //存檔rs232
                                    logDumpping.LogDumpToFile(serialPortConfig_B, GlobalData.portConfigGroup_B.portName, ref logB_text);
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logB_text = string.Empty; //清除logB_text
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "CRC16_Modbus")
                                {
                                    string orginal_data = columns_serial;
                                    string crc16_data = Crc16.PID_CRC16(orginal_data);
                                    Outputstring = orginal_data + crc16_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    //serialPortB.WriteDataOut(Outputbytes, Outputbytes.Length);
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length);
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "XOR8")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "MOD256")
                                {
                                    string orginal_data = columns_serial;
                                    byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(orginal_data);
                                    Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "")
                                {
                                    string hexValues = columns_serial;
                                    byte[] Outputbytes = new byte[hexValues.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(hexValues);
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length);
                                }
                                DateTime dt = DateTime.Now;
                                string dataValue = "[" + serialPortConfig_B + "(" + GlobalData.portConfigGroup_B.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
								
                                logDumpping.LogCat(ref logB_text, dataValue);
                                logDumpping.LogCat(ref logAll_text, dataValue);
                            }

                            if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                            {
                                log.Debug("Hex Log: _PortC");
                                if (columns_serial == "_save")
                                {
                                    //Serialportsave("C"); //存檔rs232
                                    logDumpping.LogDumpToFile(serialPortConfig_C, GlobalData.portConfigGroup_C.portName, ref logC_text);
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logC_text = string.Empty; //清除logC_text
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "CRC16_Modbus")
                                {
                                    string orginal_data = columns_serial;
                                    string crc16_data = Crc16.PID_CRC16(orginal_data);
                                    Outputstring = orginal_data + crc16_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "XOR8")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "MOD256")
                                {
                                    string orginal_data = columns_serial;
                                    byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(orginal_data);
                                    Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "")
                                {
                                    string hexValues = columns_serial;
                                    byte[] Outputbytes = new byte[hexValues.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(hexValues);
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length);	//發送數據 Rs232
                                }
                                DateTime dt = DateTime.Now;
                                string dataValue = "[" + serialPortConfig_C + "(" + GlobalData.portConfigGroup_C.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                logDumpping.LogCat(ref logC_text, dataValue);
                                logDumpping.LogCat(ref logAll_text, dataValue);
                            }

                            if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                            {
                                log.Debug("Hex Log: _PortD");
                                if (columns_serial == "_save")
                                {
                                    //Serialportsave("D"); //存檔rs232
                                    logDumpping.LogDumpToFile(serialPortConfig_D, GlobalData.portConfigGroup_D.portName, ref logD_text);
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logD_text = string.Empty; //清除logD_text
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "CRC16_Modbus")
                                {
                                    string orginal_data = columns_serial;
                                    string crc16_data = Crc16.PID_CRC16(orginal_data);
                                    Outputstring = orginal_data + crc16_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length);	//發送數據 Rs232 + Crc16
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "XOR8")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "MOD256")
                                {
                                    string orginal_data = columns_serial;
                                    byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(orginal_data);
                                    Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "")
                                {
                                    string hexValues = columns_serial;
                                    byte[] Outputbytes = new byte[hexValues.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(hexValues);
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                }
                                DateTime dt = DateTime.Now;
                                string dataValue = "[" + serialPortConfig_D + "(" + GlobalData.portConfigGroup_D.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                logDumpping.LogCat(ref logD_text, dataValue);
                                logDumpping.LogCat(ref logAll_text, dataValue);
                            }

                            if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                            {
                                log.Debug("Hex Log: _PortE");
                                if (columns_serial == "_save")
                                {
                                    //Serialportsave("E"); //存檔rs232
                                    logDumpping.LogDumpToFile(serialPortConfig_E, GlobalData.portConfigGroup_E.portName, ref logE_text);
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logE_text = string.Empty; //清除logE_text
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "CRC16_Modbus")
                                {
                                    string orginal_data = columns_serial;
                                    string crc16_data = Crc16.PID_CRC16(orginal_data);
                                    Outputstring = orginal_data + crc16_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); ; //發送數據 Rs232 + Crc16
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "XOR8")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                }
                                else if (columns_serial != "_save" && columns_serial != "_clear" &&
                                         columns_serial != "" && columns_function == "MOD256")
                                {
                                    string orginal_data = columns_serial;
                                    byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(orginal_data);
                                    Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                }
                                else if (columns_serial != "_save" &&
                                         columns_serial != "_clear" &&
                                         columns_serial != "" &&
                                         columns_function == "")
                                {
                                    string hexValues = columns_serial;
                                    byte[] Outputbytes = new byte[hexValues.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(hexValues);
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                }
                                DateTime dt = DateTime.Now;
                                string dataValue = "[" + serialPortConfig_E + "(" + GlobalData.portConfigGroup_E.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                textBox_serial.AppendText(dataValue);
                                logDumpping.LogCat(ref logE_text, dataValue);
                                logDumpping.LogCat(ref logAll_text, dataValue);
                            }

                            if (columns_comport == "ALL")
                            {
                                log.Debug("Hex Log: _All");
                                string[] serial_content = columns_serial.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                                if (columns_serial == "_save")
                                {
                                    Serialportsave("All"); //存檔rs232
                                }
                                else if (columns_serial == "_clear")
                                {
                                    logAll_text = string.Empty; //清除logAll_text
                                }

                                if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "ALL" && serial_content[0] != "")
                                {
                                    string orginal_data = serial_content[0];
                                    if (columns_function == "CRC16_Modbus")
                                    {
                                        string crc16_data = Crc16.PID_CRC16(orginal_data);
                                        Outputstring = orginal_data + crc16_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                    }
                                    else if (columns_function == "XOR8")
                                    {
                                        string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                        Outputstring = orginal_data + xor8_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                    }
                                    else if (columns_function == "MOD256")
                                    {
                                        string mod256_data = columns_serial;
                                        byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(mod256_data);
                                        Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                        GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                    }
                                    else
                                    {
                                        Outputstring = orginal_data;
                                        byte[] Outputbytes = serial_content[0].Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                        GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[" + serialPortConfig_A + "(" + GlobalData.portConfigGroup_A.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    textBox_serial.AppendText(dataValue);
                                    logDumpping.LogCat(ref logA_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "ALL" && serial_content[1] != "")
                                {
                                    string orginal_data = serial_content[1];
                                    if (columns_function == "CRC16_Modbus")
                                    {
                                        string crc16_data = Crc16.PID_CRC16(orginal_data);
                                        Outputstring = orginal_data + crc16_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                    }
                                    else if (columns_function == "XOR8")
                                    {
                                        string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                        Outputstring = orginal_data + xor8_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                    }
                                    else if (columns_function == "MOD256")
                                    {
                                        string mod256_data = columns_serial;
                                        byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(mod256_data);
                                        Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                        GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                    }
                                    else
                                    {
                                        Outputstring = orginal_data;
                                        byte[] Outputbytes = serial_content[1].Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                        GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[" + serialPortConfig_B + "(" + GlobalData.portConfigGroup_B.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    textBox_serial.AppendText(dataValue);
                                    logDumpping.LogCat(ref logB_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "ALL" && serial_content[2] != "")
                                {
                                    string orginal_data = serial_content[2];
                                    if (columns_function == "CRC16_Modbus")
                                    {
                                        string crc16_data = Crc16.PID_CRC16(orginal_data);
                                        Outputstring = orginal_data + crc16_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                    }
                                    else if (columns_function == "XOR8")
                                    {
                                        string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                        Outputstring = orginal_data + xor8_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                    }
                                    else if (columns_function == "MOD256")
                                    {
                                        string mod256_data = columns_serial;
                                        byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(mod256_data);
                                        Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                        GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                    }
                                    else
                                    {
                                        Outputstring = orginal_data;
                                        byte[] Outputbytes = serial_content[2].Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                        GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[" + serialPortConfig_C + "(" + GlobalData.portConfigGroup_C.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    textBox_serial.AppendText(dataValue);
                                    logDumpping.LogCat(ref logC_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "ALL" && serial_content[3] != "")
                                {
                                    string orginal_data = serial_content[3];
                                    if (columns_function == "CRC16_Modbus")
                                    {
                                        string crc16_data = Crc16.PID_CRC16(orginal_data);
                                        Outputstring = orginal_data + crc16_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                    }
                                    else if (columns_function == "XOR8")
                                    {
                                        string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                        Outputstring = orginal_data + xor8_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                    }
                                    else if (columns_function == "MOD256")
                                    {
                                        string mod256_data = columns_serial;
                                        byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(mod256_data);
                                        Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                        GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                    }
                                    else
                                    {
                                        Outputstring = orginal_data;
                                        byte[] Outputbytes = serial_content[3].Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                        GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[" + serialPortConfig_D + "(" + GlobalData.portConfigGroup_D.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    textBox_serial.AppendText(dataValue);
									logDumpping.LogCat(ref logD_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "ALL" && serial_content[4] != "")
                                {
                                    string orginal_data = serial_content[4];
                                    if (columns_function == "CRC16_Modbus")
                                    {
                                        string crc16_data = Crc16.PID_CRC16(orginal_data);
                                        Outputstring = orginal_data + crc16_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Crc16
                                    }
                                    else if (columns_function == "XOR8")
                                    {
                                        string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                        Outputstring = orginal_data + xor8_data;
                                        byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                        Outputbytes = HexConverter.StrToByte(Outputstring);
                                        GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Xor8
                                    }
                                    else if (columns_function == "MOD256")
                                    {
                                        string mod256_data = columns_serial;
                                        byte[] Outputbytes = algorithm.MOD256_BytesWithChksum(mod256_data);
                                        Outputstring = BitConverter.ToString(Outputbytes).Replace("-", " ");
                                        GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232 + Mod256
                                    }
                                    else
                                    {
                                        Outputstring = orginal_data;
                                        byte[] Outputbytes = serial_content[4].Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                        GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[" + serialPortConfig_E + "(" + GlobalData.portConfigGroup_E.portName + ")] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    textBox_serial.AppendText(dataValue);
									logDumpping.LogCat(ref logE_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }
                            label_Command.Text = "(" + columns_command + ") " + Outputstring;
                        }
                        #endregion

                        #region -- Minolta --
                        else if (columns_command == "_OPM")
                        {
                            if (columns_comport == "None")
                            {
                                if (columns_function == "GetDUTSensor")
                                {
                                    log.Debug("DUT sensor control: DUT sensor start");
                                    string dataValue = "";
                                    dataValue = rk2797.GetDUTSensor(columns_remark);

                                    logDumpping.LogCat(ref minolta_csv_report, dataValue);
                                    logDumpping.LogCat(ref minolta_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                    log.Debug("DUT sensor control: DUT sensor end");
                                }
                                else if (columns_serial == "_save")
                                {
                                    saveCA210csv(columns_remark); //存檔ca210
                                }
                                else if (columns_serial == "_clear")
                                {
                                    minolta_csv_report = "Sx, Sy, Lv, T, duv, Display mode, X, Y, Z, Date, Time, Scenario, Now measure count, Target measure count, Backlight sensor, Thanmal sensor, \r\n";
                                }
                            }
                            else
                            {
                                if (CA210.Status() == true)
                                {
                                    if (columns_function == "Measure")
                                    {
                                        int mtimes = 0, mRepeat = 0;
                                        string dataValue = "";
                                        Stopwatch sw = new Stopwatch();
                                        sw.Start();
                                        log.Debug("CA210 control: Measure start");
                                        if (columns_times != "" && int.TryParse(columns_times, out mtimes) == true && columns_interval != "" && int.TryParse(columns_interval, out mRepeat) == true)
                                        {
                                            mtimes = int.Parse(columns_times); // 量測次數
                                            mRepeat = int.Parse(columns_interval); // 量測時間
                                            dataValue = CA210.Measure_Multi(mtimes, mRepeat, columns_remark);
                                        }
                                        else
                                            dataValue = CA210.Measure_Once(columns_remark);

                                        logDumpping.LogCat(ref minolta_csv_report, dataValue);
                                        logDumpping.LogCat(ref minolta_text, dataValue);
                                        logDumpping.LogCat(ref logAll_text, dataValue);
                                        sw.Stop();
                                        log.Debug($"Minolta Measure: { sw.ElapsedMilliseconds}ms");
                                        log.Debug("CA210 control: Measure stop");
                                    }
                                    else if (columns_function == "DisplayMode")
                                    {
                                        Stopwatch sw = new Stopwatch();
                                        sw.Start();
                                        log.Debug("CA210 control: DisplayMode start");
                                        if (columns_times != "" && int.TryParse(columns_times, out stime) == true)
                                            stime = int.Parse(columns_times); // 模式切換
                                        else
                                            stime = 0;

                                        CA210.DisplayMode(stime);
                                        sw.Stop();
                                        log.Debug($"Minolta DisplayMode: { sw.ElapsedMilliseconds}ms");
                                        log.Debug("CA210 control: DisplayMode end");
                                    }
                                    else if (columns_function == "CalZero")
                                    {
                                        Stopwatch sw = new Stopwatch();
                                        sw.Start();
                                        log.Debug("CA210 control: Zero-calibrates the device start");
                                        CA210.CalZero();
                                        sw.Stop();
                                        log.Debug($"Minolta CalZero: { sw.ElapsedMilliseconds}ms");
                                        log.Debug("CA210 control: Zero-calibrates the device end");
                                    }

                                    if (columns_serial == "_save")
                                    {
                                        saveCA210csv(columns_remark); //存檔ca210
                                    }
                                    else if (columns_serial == "_clear")
                                    {
                                        minolta_csv_report = "Sx, Sy, Lv, T, duv, Display mode, X, Y, Z, Date, Time, Scenario, Now measure count, Target measure count, Backlight sensor, Thanmal sensor, \r\n";
                                    }
                                }
                                else if (CA210.Status() == false)
                                {
                                    MessageBox.Show("Minolta is not connected!\r\nPlease restart the OPTT to reload the device.", "Connection Error");
                                }
                            }
                        }
                        #endregion

                        #region -- K-Line --
                        else if (columns_command == "_K_ABS")
                        {
                            log.Debug("K-line control: _K_ABS");
                            try
                            {
                                // K-lite ABS指令檔案匯入
                                string xmlfile = ini12.INIRead(MainSettingPath, "Record", "Generator", "");
                                if (System.IO.File.Exists(xmlfile) == true)
                                {
                                    var allDTC = XDocument.Load(xmlfile).Root.Element("ABS_ErrorCode").Elements("DTC");
                                    foreach (var ErrorCode in allDTC)
                                    {
                                        if (ErrorCode.Attribute("Name").Value == "_ABS")
                                        {
                                            if (columns_serial == ErrorCode.Element("DTC_D").Value)
                                            {
                                                UInt16 int_abs_code = Convert.ToUInt16(ErrorCode.Element("DTC_C").Value, 16);
                                                byte abs_code_high = Convert.ToByte(int_abs_code >> 8);
                                                byte abs_code_low = Convert.ToByte(int_abs_code & 0xff);
                                                byte abs_code_status = Convert.ToByte(ErrorCode.Element("DTC_S").Value, 16);
                                                ABS_error_list.Add(new DTC_Data(abs_code_high, abs_code_low, abs_code_status));
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Content includes other error code", "ABS code Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("DTC code file does not exist", "File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "Kline_ABS library error!");
                            }
                        }
                        else if (columns_command == "_K_OBD")
                        {
                            log.Debug("K-line control: _K_OBD");
                            try
                            {
                                // K-lite OBD指令檔案匯入
                                string xmlfile = ini12.INIRead(MainSettingPath, "Record", "Generator", "");
                                if (System.IO.File.Exists(xmlfile) == true)
                                {
                                    var allDTC = XDocument.Load(xmlfile).Root.Element("OBD_ErrorCode").Elements("DTC");
                                    foreach (var ErrorCode in allDTC)
                                    {
                                        if (ErrorCode.Attribute("Name").Value == "_OBD")
                                        {
                                            if (columns_serial == ErrorCode.Element("DTC_D").Value)
                                            {
                                                UInt16 obd_code_int16 = Convert.ToUInt16(ErrorCode.Element("DTC_C").Value, 16);
                                                byte obd_code_high = Convert.ToByte(obd_code_int16 >> 8);
                                                byte obd_code_low = Convert.ToByte(obd_code_int16 & 0xff);
                                                byte obd_code_status = Convert.ToByte(ErrorCode.Element("DTC_S").Value, 16);
                                                OBD_error_list.Add(new DTC_Data(obd_code_high, obd_code_low, obd_code_status));
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Content includes other error code", "OBD code Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("DTC code file does not exist", "File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "Kline_OBD library error !");
                            }
                        }
                        else if (columns_command == "_K_SEND")
                        {
                            kline_send = 1;
                        }
                        else if (columns_command == "_K_CLEAR")
                        {
                            kline_send = 0;
                            ABS_error_list.Clear();
                            OBD_error_list.Clear();
                        }
                        #endregion

                        #region -- I2C Read --
                        else if (columns_command == "_TX_I2C_Read")
                        {
                            if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                            {
                                log.Debug("I2C Read Log: _TX_I2C_Read_PortA");
                                if (columns_times != "" && columns_function != "")
                                {
                                    string orginal_data = columns_times + " " + columns_function + " " + "20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6D " + columns_times + " 06 " + columns_function + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logA_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                            {
                                log.Debug("I2C Read Log: _TX_I2C_Read_PortB");
                                if (columns_times != "" && columns_function != "")
                                {
                                    string orginal_data = columns_times + " " + columns_function + " " + "20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6D " + columns_times + " 06 " + columns_function + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logB_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                            {
                                log.Debug("I2C Read Log: _TX_I2C_Read_PortC");
                                if (columns_times != "" && columns_function != "")
                                {
                                    string orginal_data = columns_times + " " + columns_function + " " + "20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6D " + columns_times + " 06 " + columns_function + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logC_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                            {
                                log.Debug("I2C Read Log: _TX_I2C_Read_PortD");
                                if (columns_times != "" && columns_function != "")
                                {
                                    string orginal_data = columns_times + " " + columns_function + " " + "20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6D " + columns_times + " 06 " + columns_function + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logD_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                            {
                                log.Debug("I2C Read Log: _TX_I2C_Read_PortE");
                                if (columns_times != "" && columns_function != "")
                                {
                                    string orginal_data = columns_times + " " + columns_function + " " + "20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6D " + columns_times + " 06 " + columns_function + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logE_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }
                        }
                        #endregion

                        #region -- I2C Write --
                        else if (columns_command == "_TX_I2C_Write")
                        {
                            if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                            {
                                log.Debug("I2C Write Log: _TX_I2C_Write_PortA");
                                if (columns_function != "" && columns_subFunction != "")
                                {
                                    int Data_length = columns_subFunction.Split(' ').Count();
                                    string orginal_data = (Data_length + 1).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6C " + (Data_length + 1).ToString("X2") + " " + (Data_length + 6).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_A.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logA_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                            {
                                log.Debug("I2C Write Log: _TX_I2C_Write_PortB");
                                if (columns_function != "" && columns_subFunction != "")
                                {
                                    int Data_length = columns_subFunction.Split(' ').Count();
                                    string orginal_data = (Data_length + 1).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6C " + (Data_length + 1).ToString("X2") + " " + (Data_length + 6).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_B.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logB_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                            {
                                log.Debug("I2C Write Log: _TX_I2C_Write_PortC");
                                if (columns_function != "" && columns_subFunction != "")
                                {
                                    int Data_length = columns_subFunction.Split(' ').Count();
                                    string orginal_data = (Data_length + 1).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6C " + (Data_length + 1).ToString("X2") + " " + (Data_length + 6).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_C.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logC_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                            {
                                log.Debug("I2C Write Log: _TX_I2C_Write_PortD");
                                if (columns_function != "" && columns_subFunction != "")
                                {
                                    int Data_length = columns_subFunction.Split(' ').Count();
                                    string orginal_data = (Data_length + 1).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6C " + (Data_length + 1).ToString("X2") + " " + (Data_length + 6).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_D.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logD_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }

                            if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                            {
                                log.Debug("I2C Write Log: _TX_I2C_Write_PortE");
                                if (columns_function != "" && columns_subFunction != "")
                                {
                                    int Data_length = columns_subFunction.Split(' ').Count();
                                    string orginal_data = (Data_length + 1).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20";
                                    string crc32_data = Crc32.I2C_CRC32(orginal_data);
                                    string Outputstring = "79 6C " + (Data_length + 1).ToString("X2") + " " + (Data_length + 6).ToString("X2") + " " + columns_function + " " + columns_subFunction + " 20 " + crc32_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    GlobalData.m_SerialPort_E.WriteDataOut(Outputbytes, Outputbytes.Length); //發送數據 Rs232
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref logE_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                            }
                        }
                        #endregion

                        #region -- FTDI Read/Write --
                        else if (columns_command == "_FTDI")
                        {
                            if (portinfo.ftStatus == FtResult.Ok)
                            {
                                log.Debug("FTDI Write Log: _FTDI_Read_Write");
                                Algorithm algorithm = new Algorithm();
                                string Outputstring = "";
                                if (columns_serial == "_save")
                                {
                                    logDumpping.LogDumpToFile("FTDI", "FTDI Port", ref ftdi_text);
                                    Outputstring = "Save FTDI log";
                                }
                                else if (columns_serial == "_clear")
                                {
                                    ftdi_text = string.Empty; //清除ftdi_text
                                    Outputstring = "Clear FTDI log";
                                }
                                else if (columns_function == "Write" && columns_serial != "")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    byte DeviceAddr = Outputbytes[0];
                                    byte[] DeviceData = new byte[Outputstring.Split(' ').Count() - 1];
                                    int i;
                                    for (i = 0; i < DeviceData.Length; i++)
                                    {
                                        DeviceData[i] = Outputbytes[i + 1];
                                    }
                                    Ftdi_lib.I2C_SEQ_Write(portinfo.ftHandle, DeviceAddr, DeviceData); //發送數據 Rs232 + Xor8
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Ftdi_Port_Send] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref ftdi_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                }
                                else if (columns_function == "Read" && columns_serial != "")
                                {
                                    string orginal_data = columns_serial;
                                    string xor8_data = algorithm.Medical_XOR8(orginal_data);
                                    byte[] Readbytes = new byte[128];
                                    byte Readbyte;
                                    Outputstring = orginal_data + xor8_data;
                                    byte[] Outputbytes = new byte[Outputstring.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(Outputstring);
                                    byte DeviceAddr = Outputbytes[0];
                                    byte[] DeviceData = new byte[Outputstring.Split(' ').Count() - 1];
                                    int i;
                                    for (i = 0; i < DeviceData.Length; i++)
                                    {
                                        DeviceData[i] = Outputbytes[i + 1];
                                    }
                                    DateTime dt = DateTime.Now;
                                    string dataValue = "[Ftdi_Port_Send] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(ref ftdi_text, dataValue);
                                    logDumpping.LogCat(ref logAll_text, dataValue);
                                    if (Ftdi_lib.I2C_SEQ_Read(portinfo.ftHandle, DeviceAddr, DeviceData, Readbytes, out Readbyte) == FtResult.Ok)
                                    {
                                        byte[] Getbytes = new byte[Readbyte];
                                        byte checksum;
                                        Array.Copy(Readbytes, Getbytes, Getbytes.Length);
                                        checksum = Getbytes[Getbytes.Length - 1];
                                        Getbytes[Getbytes.Length - 1] = 0x50;
                                        if (checksum == algorithm.XOR_Byte(Getbytes, Getbytes.Length))
                                        {
                                            Getbytes[Getbytes.Length - 1] = checksum;
                                            string inputstring = BitConverter.ToString(Getbytes).Replace("-", " ");
                                            dt = DateTime.Now;
                                            dataValue = "[Ftdi_Port_Receive] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + inputstring + "\r\n";
                                            logDumpping.LogCat(ref ftdi_text, dataValue);
                                            logDumpping.LogCat(ref logAll_text, dataValue);
                                        }
                                        else
                                            MessageBox.Show("Ftdi_Port_Receive no data !!", "The checksum error!!");
                                    }
                                }
                                label_Command.Text = "(" + columns_command + ") " + Outputstring;
                            }
                        }
                        #endregion

                        #region -- Canbus Send --
                        else if (columns_command == "_Canbus_Send")
                        {
                            if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "UsbCAN")
                            {
                                if (columns_times != "" && columns_interval == "" && columns_serial != "")
                                {
                                    log.Debug("Canbus Send (Event): _Canbus_Send");
                                    byte[] Outputdata = new byte[columns_serial.Split(' ').Count()];
                                    Outputdata = HexConverter.StrToByte(columns_serial);
                                    Can_Usb2C.TransmitData(Convert.ToUInt32(columns_times), Outputdata);

                                    string Outputstring = "ID: 0x";
                                    Outputstring += columns_times + " Data: " + columns_serial;
                                    DateTime dt = DateTime.Now;
                                    string canbus_log_text = "[Send_Canbus] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(canbus_text, canbus_log_text);
                                    logDumpping.LogCat(logAll_text, canbus_log_text);
                                }
                                else if (columns_times != "" && columns_interval != "" && columns_serial != "")
                                {
                                    set_timer_rate = true;
                                    can_id = System.Convert.ToUInt16("0x" + columns_times, 16);
                                    byte[] Outputdata = new byte[columns_serial.Split(' ').Count()];
                                    Outputdata = HexConverter.StrToByte(columns_serial);
                                    if (can_rate.Count > 0)
                                    {
                                        foreach (var OneItem in can_rate)
                                        {
                                            if (can_rate.ContainsKey(can_id))
                                            {
                                                can_rate[can_id] = Convert.ToUInt32(columns_interval);
                                                can_data[can_id] = Outputdata;
                                                break;
                                            }
                                            else
                                            {
                                                can_rate.Add(can_id, Convert.ToUInt32(columns_interval));
                                                can_data.Add(can_id, Outputdata);
                                                UsbCAN_Count = 0;
                                                UsbCAN_Delay(Convert.ToInt16(columns_interval));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        can_rate.Add(can_id, Convert.ToUInt32(columns_interval));
                                        can_data.Add(can_id, Outputdata);
                                        UsbCAN_Count = 0;
                                        UsbCAN_Delay(Convert.ToInt16(columns_interval));
                                    }
                                }
                            }
                            else if (ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "Vector")
                            {
                                if (columns_times != "" && columns_interval == "" && columns_serial != "")
                                {
                                    log.Debug("Canbus Send: Vector_Canbus_once");
                                    byte[] Outputdata = new byte[columns_serial.Split(' ').Count()];
                                    Outputdata = HexConverter.StrToByte(columns_serial);
                                    Can_1630A.LoopCANTransmit(Convert.ToUInt32(columns_times), Convert.ToUInt32(columns_interval), Outputdata);

                                    string Outputstring = "ID: 0x";
                                    Outputstring += columns_times + " Data: " + columns_serial;
                                    DateTime dt = DateTime.Now;
                                    string canbus_log_text = "[Send_Canbus] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                                    logDumpping.LogCat(canbus_text, canbus_log_text);
                                    logDumpping.LogCat(logAll_text, canbus_log_text);
                                }
                                else if (columns_times != "" && columns_interval != "" && columns_serial != "")
                                {
                                    log.Debug("Canbus Send: Vector_Canbus_loop");
                                    set_timer_rate = true;
                                    can_id = System.Convert.ToUInt16("0x" + columns_times, 16);
                                    byte[] Outputdata = new byte[columns_serial.Split(' ').Count()];
                                    Outputdata = HexConverter.StrToByte(columns_serial);
                                    if (can_rate.Count > 0)
                                    {
                                        foreach (var OneItem in can_rate)
                                        {
                                            if (can_rate.ContainsKey(can_id))
                                            {
                                                can_rate[can_id] = Convert.ToUInt32(columns_interval);
                                                can_data[can_id] = Outputdata;
                                                break;
                                            }
                                            else
                                            {
                                                can_rate.Add(can_id, Convert.ToUInt32(columns_interval));
                                                can_data.Add(can_id, Outputdata);
                                                //VectorCAN_Count = 0;
                                                //VectorCAN_Delay(Convert.ToInt16(columns_interval));
                                                Thread CanSetTimeRate = new Thread(new ThreadStart(vectorcanloop));
                                                CanSetTimeRate.Start();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        can_rate.Add(can_id, Convert.ToUInt32(columns_interval));
                                        can_data.Add(can_id, Outputdata);
                                        //VectorCAN_Count = 0;
                                        //VectorCAN_Delay(Convert.ToInt16(columns_interval));
                                        Thread CanSetTimeRate = new Thread(new ThreadStart(vectorcanloop));
                                        CanSetTimeRate.Start();
                                    }
                                }
                            }

                            label_Command.Text = "(" + columns_command + ") " + columns_serial;
                        }
                        #endregion

                        #region -- Canbus Queue --
                        else if (columns_command == "_Canbus_Queue")
                        {
                            if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "UsbCAN")
                            {
                                if (columns_times != "" && columns_interval != "" && columns_serial != "")
                                {
                                    log.Debug("Canbus Write: UsbCAN_Canbus_Queue_data");
                                    byte[] Outputbytes = new byte[columns_serial.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(columns_serial);
                                    can_data_list.Add(new USB_CAN2C.CAN_Data(System.Convert.ToUInt16("0x" + columns_times, 16), System.Convert.ToUInt32(columns_interval), Outputbytes, Convert.ToByte(columns_serial.Split(' ').Count())));
                                }
                                else if (columns_function == "send")
                                {
                                    log.Debug("Canbus Write: UsbCAN_Canbus_Queue_send");
                                    can_send = 1;
                                }
                                else if (columns_function == "clear")
                                {
                                    log.Debug("Canbus Write: UsbCAN_Canbus_Queue_clean");
                                    can_send = 0;
                                    can_data_list.Clear();
                                }
                            }
                            else if (ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "Vector")
                            {
                                if (columns_times != "" && columns_interval != "" && columns_serial != "")
                                {
                                    log.Debug("Canbus Write: Vector_Canbus_Queue_data");
                                    byte[] Outputbytes = new byte[columns_serial.Split(' ').Count()];
                                    Outputbytes = HexConverter.StrToByte(columns_serial);
                                    can_data_list.Add(new USB_CAN2C.CAN_Data(System.Convert.ToUInt16("0x" + columns_times, 16), System.Convert.ToUInt32(columns_interval), Outputbytes, Convert.ToByte(columns_serial.Split(' ').Count())));
                                }
                                else if (columns_function == "send")
                                {
                                    log.Debug("Canbus Write: Vector_Canbus_Queue_send");
                                    can_send = 1;
                                }
                                else if (columns_function == "clear")
                                {
                                    log.Debug("Canbus Write: Vector_Canbus_Queue_clean");
                                    can_send = 0;
                                    can_data_list.Clear();
                                }
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_serial;
                        }
                        #endregion

                        #region -- Astro Timing --
                        else if (columns_command == "_astro")
                        {
                            log.Debug("Astro control: _astro");
                            try
                            {
                                // Astro指令
                                byte[] startbit = new byte[7] { 0x05, 0x24, 0x20, 0x02, 0xfd, 0x24, 0x20 };
                                PortA.Write(startbit, 0, 7);

                                // Astro指令檔案匯入
                                string xmlfile = ini12.INIRead(MainSettingPath, "Record", "Generator", "");
                                if (System.IO.File.Exists(xmlfile) == true)
                                {
                                    var allTiming = XDocument.Load(xmlfile).Root.Element("Generator").Elements("Device");
                                    foreach (var generator in allTiming)
                                    {
                                        if (generator.Attribute("Name").Value == "_astro")
                                        {
                                            if (columns_function == generator.Element("Timing").Value)
                                            {
                                                string[] timestrs = generator.Element("Signal").Value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                                                byte[] timebit1 = Encoding.ASCII.GetBytes(timestrs[0]);
                                                byte[] timebit2 = Encoding.ASCII.GetBytes(timestrs[1]);
                                                byte[] timebit3 = Encoding.ASCII.GetBytes(timestrs[2]);
                                                byte[] timebit4 = Encoding.ASCII.GetBytes(timestrs[3]);
                                                byte[] timebit = new byte[4] { timebit1[1], timebit2[1], timebit3[1], timebit4[1] };
                                                PortA.Write(timebit, 0, 4);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Content include other signal", "Astro Signal Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Signal Generator not exist", "File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                                byte[] endbit = new byte[3] { 0x2c, 0x31, 0x03 };
                                PortA.Write(endbit, 0, 3);
                                label_Command.Text = "(" + columns_command + ") " + columns_switch;
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "Transmit the Astro command fail !");
                            }
                        }
                        #endregion

                        #region -- Quantum Timing --
                        else if (columns_command == "_quantum")
                        {
                            log.Debug("Quantum control: _quantum");
                            try
                            {
                                // Quantum指令檔案匯入
                                string xmlfile = ini12.INIRead(MainSettingPath, "Record", "Generator", "");
                                if (System.IO.File.Exists(xmlfile) == true)
                                {
                                    var allTiming = XDocument.Load(xmlfile).Root.Element("Generator").Elements("Device");
                                    foreach (var generator in allTiming)
                                    {
                                        if (generator.Attribute("Name").Value == "_quantum")
                                        {
                                            if (columns_function == generator.Element("Timing").Value)
                                            {
                                                PortA.WriteLine(generator.Element("Signal").Value + "\r");
                                                PortA.WriteLine("ALLU" + "\r");
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Content include other signal", "Quantum Signal Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Signal Generator not exist", "File Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                                switch (columns_subFunction)
                                {
                                    case "RGB":
                                        // RGB mode
                                        PortA.WriteLine("AVST 0" + "\r");
                                        PortA.WriteLine("DVST 10" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "YCbCr":
                                        // YCbCr mode
                                        PortA.WriteLine("AVST 0" + "\r");
                                        PortA.WriteLine("DVST 14" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "xvYCC":
                                        // xvYCC mode
                                        PortA.WriteLine("AVST 0" + "\r");
                                        PortA.WriteLine("DVST 17" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "4:4:4":
                                        // 4:4:4
                                        PortA.WriteLine("DVSM 4" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "4:2:2":
                                        // 4:2:2
                                        PortA.WriteLine("DVSM 2" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "8bits":
                                        // 8bits
                                        PortA.WriteLine("NBPC 8" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "10bits":
                                        // 10bits
                                        PortA.WriteLine("NBPC 10" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    case "12bits":
                                        // 12bits
                                        PortA.WriteLine("NBPC 12" + "\r");
                                        PortA.WriteLine("FMTU" + "\r");
                                        break;
                                    default:
                                        break;
                                }
                                label_Command.Text = "(" + columns_command + ") " + columns_switch + columns_remark;
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "Transmit the Quantum command fail !");
                            }
                        }
                        #endregion

                        #region -- Dektec --
                        else if (columns_command == "_dektec")
                        {
                            if (columns_switch == "_start")
                            {
                                log.Debug("Dektec control: _start");
                                string StreamName = columns_serial;
                                string TvSystem = columns_function;
                                string Freq = columns_subFunction;
                                string arguments = Application.StartupPath + @"\\DektecPlayer\\" + StreamName + " " +
                                                   "-mt " + TvSystem + " " +
                                                   "-mf " + Freq + " " +
                                                   "-r 0 " +
                                                   "-l 0";

                                Console.WriteLine(arguments);
                                System.Diagnostics.Process Dektec = new System.Diagnostics.Process();
                                Dektec.StartInfo.FileName = Application.StartupPath + @"\\DektecPlayer\\DtPlay.exe";
                                Dektec.StartInfo.UseShellExecute = false;
                                Dektec.StartInfo.RedirectStandardInput = true;
                                Dektec.StartInfo.RedirectStandardOutput = true;
                                Dektec.StartInfo.RedirectStandardError = true;
                                Dektec.StartInfo.CreateNoWindow = true;

                                Dektec.StartInfo.Arguments = arguments;
                                Dektec.Start();
                                label_Command.Text = "(" + columns_command + ") " + columns_serial;
                            }

                            if (columns_switch == "_stop")
                            {
                                log.Debug("Dektec control: _stop");
                                CloseDtplay();
                            }
                        }
                        #endregion

                        #region -- 命令提示 --
                        else if (columns_command == "_DOS")
                        {
                            log.Debug("DOS command: _DOS");
                            if (columns_serial != "")
                            {
                                string Command = columns_serial;

                                System.Diagnostics.Process p = new Process();
                                p.StartInfo.FileName = "cmd.exe";
                                p.StartInfo.WorkingDirectory = ini12.INIRead(MainSettingPath, "Device", "DOS", "");
                                p.StartInfo.UseShellExecute = false;
                                p.StartInfo.RedirectStandardInput = true;
                                p.StartInfo.RedirectStandardOutput = true;
                                p.StartInfo.RedirectStandardError = true;
                                p.StartInfo.CreateNoWindow = true; //不跳出cmd視窗
                                string strOutput = null;

                                try
                                {
                                    p.Start();
                                    p.StandardInput.WriteLine(Command);
                                    label_Command.Text = "DOS CMD_" + columns_serial;
                                    //p.StandardInput.WriteLine("exit");
                                    //strOutput = p.StandardOutput.ReadToEnd();//匯出整個執行過程
                                    //p.WaitForExit();
                                    //p.Close();
                                }
                                catch (Exception e)
                                {
                                    strOutput = e.Message;
                                }
                            }
                        }
                        #endregion

                        #region -- GPIO_INPUT_OUTPUT --
                        else if (columns_command == "_IO_Input")
                        {
                            log.Debug("GPIO control: _IO_Input");
                            IO_INPUT();
                        }

                        else if (columns_command == "_IO_Output")
                        {
                            log.Debug("GPIO control: _IO_Output");
                            //string GPIO = "01010101";
                            string GPIO = columns_times;
                            byte GPIO_B = Convert.ToByte(GPIO, 2);
                            MyBlueRat.Set_GPIO_Output(GPIO_B);
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }

                        else if (columns_command == "_Arduino_Input")
                        {
                            log.Debug("GPIO control: _Arduino_Input");
                            GlobalData.Arduino_outputFlag = false;
                            Arduino_IO_INPUT(SysDelay);
                        }

                        else if (columns_command == "_Arduino_Output")
                        {
                            log.Debug("GPIO control: _Arduino_Output");
                            //string GPIO = "01010101";
                            string GPIO = columns_times;
                            byte GPIO_B = Convert.ToByte(GPIO, 2);
                            Arduino_Set_GPIO_Output(GPIO_B, SysDelay);
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }

                        else if (columns_command == "_Arduino_Command")
                        {
                            log.Debug("GPIO control: _Arduino_Command");
                            if (columns_serial != "")
                            {
                                GlobalData.Arduino_outputFlag = true;
                                Arduino_Set_Value(columns_serial, SysDelay);

                            }
                            else
                                MessageBox.Show("Please check the Arduino command.", "Arduino command Error!");
                            label_Command.Text = "(" + columns_command + ") " + columns_serial;
                        }
                        #endregion

                        #region -- Extend_GPIO_OUTPUT --
                        else if (columns_command == "_WaterTemp")
                        {
                            log.Debug("Extend GPIO control: _WaterTemp");
                            string GPIO = columns_times; // GPIO = "010101010";
                            if (GPIO.Length == 9)
                            {
                                for (int i = 0; i < 9; i++)
                                {
                                    MyBlueRat.Set_IO_Extend_Set_Pin(Convert.ToByte(i), Convert.ToByte(GPIO.Substring(8 - i, 1)));
                                    Thread.Sleep(50);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Please check the value equal nine.");
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }

                        else if (columns_command == "_FuelDisplay")
                        {
                            log.Debug("Extend GPIO control: _FuelDisplay");
                            string GPIO = columns_times;
                            if (GPIO.Length == 9)
                            {
                                for (int i = 0; i < 9; i++)
                                {
                                    MyBlueRat.Set_IO_Extend_Set_Pin(Convert.ToByte(i + 16), Convert.ToByte(GPIO.Substring(8 - i, 1)));
                                    Thread.Sleep(50);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Please check the value equal nine.");
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }

                        else if (columns_command == "_Temperature")
                        {
                            log.Debug("Extend GPIO control: _Temperature");
                            //string GPIO = "01010101";
                            string GPIO = columns_serial;
                            int GPIO_B = int.Parse(GPIO);
                            if (GPIO_B >= -20 && GPIO_B <= 50)
                            {
                                if (GPIO_B >= -20 && GPIO_B < -17)
                                    MyBlueRat.Set_MCP42xxx(224);
                                else if (GPIO_B >= -17 && GPIO_B < -12)
                                    MyBlueRat.Set_MCP42xxx(172);
                                else if (GPIO_B >= -12 && GPIO_B < -7)
                                    MyBlueRat.Set_MCP42xxx(130);
                                else if (GPIO_B >= -7 && GPIO_B < -2)
                                    MyBlueRat.Set_MCP42xxx(101);
                                else if (GPIO_B >= -2 && GPIO_B < 3)
                                    MyBlueRat.Set_MCP42xxx(78);
                                else if (GPIO_B >= 3 && GPIO_B < 8)
                                    MyBlueRat.Set_MCP42xxx(61);
                                else if (GPIO_B >= 8 && GPIO_B < 13)
                                    MyBlueRat.Set_MCP42xxx(47);
                                else if (GPIO_B >= 13 && GPIO_B < 18)
                                    MyBlueRat.Set_MCP42xxx(36);
                                else if (GPIO_B >= 18 && GPIO_B < 23)
                                    MyBlueRat.Set_MCP42xxx(29);
                                else if (GPIO_B >= 23 && GPIO_B < 28)
                                    MyBlueRat.Set_MCP42xxx(23);
                                else if (GPIO_B >= 28 && GPIO_B < 33)
                                    MyBlueRat.Set_MCP42xxx(19);
                                else if (GPIO_B >= 33 && GPIO_B < 38)
                                    MyBlueRat.Set_MCP42xxx(15);
                                else if (GPIO_B >= 38 && GPIO_B < 43)
                                    MyBlueRat.Set_MCP42xxx(12);
                                else if (GPIO_B >= 43 && GPIO_B < 48)
                                    MyBlueRat.Set_MCP42xxx(10);
                                else if (GPIO_B >= 48 && GPIO_B <= 50)
                                    MyBlueRat.Set_MCP42xxx(8);
                                Thread.Sleep(50);
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }
                        #endregion

                        #region -- Push_Release_Function--
                        else if (columns_command == "_FuncKey")
                        {
                            try
                            {
                                for (int k = 0; k < stime; k++)
                                {
                                    log.Debug("Extend GPIO control: _FuncKey:" + k + " times");
                                    label_Command.Text = "(Push CMD)" + columns_serial;
                                    if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                                    {
                                        if (columns_serial == "_save")
                                        {
                                            Serialportsave("A"); //存檔rs232
                                            //logDumpping_A.SerialPortLogDump("PortA");
                                        }
                                        else if (columns_serial == "_clear")
                                        {
                                            logA_text = string.Empty; //清除textbox1
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_A, columns_serial, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logA_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                                    {
                                        if (columns_serial == "_save")
                                        {
                                            Serialportsave("B"); //存檔rs232
                                        }
                                        else if (columns_serial == "_clear")
                                        {
                                            logB_text = string.Empty; //清除logB_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_B, columns_serial, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logB_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                                    {
                                        if (columns_serial == "_save")
                                        {
                                            Serialportsave("C"); //存檔rs232
                                        }
                                        else if (columns_serial == "_clear")
                                        {
                                            logC_text = string.Empty; //清除logC_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_C, columns_serial, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logC_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                                    {
                                        if (columns_serial == "_save")
                                        {
                                            Serialportsave("D"); //存檔rs232
                                        }
                                        else if (columns_serial == "_clear")
                                        {
                                            logD_text = string.Empty; //清除logD_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_D, columns_serial, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logD_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                                    {
                                        if (columns_serial == "_save")
                                        {
                                            Serialportsave("E"); //存檔rs232
                                        }
                                        else if (columns_serial == "_clear")
                                        {
                                            logE_text = string.Empty; //清除logE_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_E, columns_serial, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logE_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    //label_Command.Text = "(" + columns_command + ") " + columns_serial;
                                    log.Debug("Extend GPIO control: _FuncKey Delay:" + sRepeat + " ms");

                                    RedRatDBViewer_Delay(sRepeat);
                                    int length = columns_serial.Length;
                                    string status = columns_serial.Substring(length - 1, 1);
                                    string reverse = "";
                                    if (status == "0")
                                        reverse = columns_serial.Substring(0, length - 1) + "1";
                                    else if (status == "1")
                                        reverse = columns_serial.Substring(0, length - 1) + "0";
                                    label_Command.Text = "(Release CMD)" + reverse;

                                    if (GlobalData.portConfigGroup_A.checkedValue == true && columns_comport == "A")
                                    {
                                        if (reverse == "_save")
                                        {
                                            Serialportsave("A"); //存檔rs232
                                            //logDumpping_A.SerialPortLogDump("Port A");
                                        }
                                        else if (reverse == "_clear")
                                        {
                                            logA_text = string.Empty; //清除textbox1
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_A, reverse, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + reverse + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logA_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_B.checkedValue == true && columns_comport == "B")
                                    {
                                        if (reverse == "_save")
                                        {
                                            Serialportsave("B"); //存檔rs232
                                        }
                                        else if (reverse == "_clear")
                                        {
                                            logB_text = string.Empty; //清除logB_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_B, reverse, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_B] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + reverse + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logB_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_C.checkedValue == true && columns_comport == "C")
                                    {
                                        if (reverse == "_save")
                                        {
                                            Serialportsave("C"); //存檔rs232
                                        }
                                        else if (reverse == "_clear")
                                        {
                                            logC_text = string.Empty; //清除logC_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_C, reverse, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_C] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + reverse + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logC_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_D.checkedValue == true && columns_comport == "D")
                                    {
                                        if (reverse == "_save")
                                        {
                                            Serialportsave("D"); //存檔rs232
                                        }
                                        else if (reverse == "_clear")
                                        {
                                            logD_text = string.Empty; //清除logD_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_D, reverse, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_D] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + reverse + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logD_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    else if (GlobalData.portConfigGroup_E.checkedValue == true && columns_comport == "E")
                                    {
                                        if (reverse == "_save")
                                        {
                                            Serialportsave("E"); //存檔rs232
                                        }
                                        else if (reverse == "_clear")
                                        {
                                            logE_text = string.Empty; //清除logE_text
                                        }
                                        else if (columns_serial != "" || columns_switch != "")
                                        {
                                            ReplaceNewLine(GlobalData.m_SerialPort_E, reverse, columns_switch);
                                        }
                                        else if (columns_serial == "" && columns_switch == "")
                                        {
                                            MessageBox.Show("Command is fail, please check the format.");
                                        }
                                        DateTime dt = DateTime.Now;
                                        string dataValue = "[Send_Port_E] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + reverse + "\r\n";
                                        textBox_serial.AppendText(dataValue);
                                        logDumpping.LogCat(logE_text, dataValue);
                                        logDumpping.LogCat(logAll_text, dataValue);
                                    }
                                    //label_Command.Text = "(" + columns_command + ") " + columns_serial;
                                    RedRatDBViewer_Delay(500);
                                }
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "SerialPort content fail !");
                            }
                        }
                        #endregion

                        #region -- MonkeyTest --
                        else if (columns_command == "_MonkeyTest")
                        {
                            log.Debug("Android control: _MonkeyTest");
                            Add_ons MonkeyTest = new Add_ons();
                            MonkeyTest.MonkeyTest();
                            MonkeyTest.CreateExcelFile();
                        }
                        #endregion

                        #region -- Factory Command 控制 --
                        /*
                        else if (columns_command == "_SXP")
                        {
                            if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" &&
                                columns_serial == "_save")
                            {
                                string fName = "";

                                fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                                string t = fName + "\\_Log2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";

                                StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                                MYFILE.WriteLine(textBox2.Text);
                                MYFILE.Close();

                                Txtbox2("", textBox2);
                            }

                            if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" &&
                                columns_serial != "_save")
                            {
                                try
                                {
                                    string str = columns_serial;
                                    byte[] bytes = str.Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
                                    label_Command.Text = "(SXP CMD)" + columns_serial;
                                    serialPort2.Write(bytes, 0, bytes.Length);
                                    label_Command.Text = "(" + columns_command + ") " + columns_serial;
                                    // DateTime dt = DateTime.Now;
                                    // string text = "[" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + columns_serial + "\r\n";
                                    str = str.Replace(" ", "");
                                    string text = str + "\r\n";
                                    textBox2.AppendText(text);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Check your SerialPort2 setting.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Question);
                                GlobalData.Break_Out_Schedule = 1;
                            }
                        }*/
                        #endregion

                        #region -- IO CMD --
                        else if (columns_command == "_Pin" && columns_comport.Length >= 7 && columns_comport.Substring(0, 3) == "_PA" ||
                                 columns_command == "_Pin" && columns_comport.Length >= 7 && columns_comport.Substring(0, 3) == "_PB")
                        {
                            {
                                switch (columns_comport.Substring(3, 2))
                                {
                                    #region -- PA10 --
                                    case "10":
                                        log.Debug("IO CMD: PA10");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(10, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA10_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(10, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA10_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- PA11 --
                                    case "11":
                                        log.Debug("IO CMD: PA11");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(8, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA11_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(8, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA11_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- PA14 --
                                    case "14":
                                        log.Debug("IO CMD: PA14");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(6, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA14_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(6, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA14_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- PA15 --
                                    case "15":
                                        log.Debug("IO CMD: PA15");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(4, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA15_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(4, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PA15_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- PB01 --
                                    case "01":
                                        log.Debug("IO CMD: PB01");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(2, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PB1_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }

                                            else
                                                IO_CMD();
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(2, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PB1_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- PB07 --
                                    case "07":
                                        log.Debug("IO CMD: PB07");
                                        if (columns_comport.Substring(6, 1) == "0" &&
                                            GlobalData.IO_INPUT.Substring(0, 1) == "0")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PB7_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else if (columns_comport.Substring(6, 1) == "1" &&
                                            GlobalData.IO_INPUT.Substring(0, 1) == "1")
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_PB7_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                                IO_CMD();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                        #endregion
                                }
                            }
                        }
                        #endregion

                        #region -- Arduino IO CMD --
                        else if (columns_command == "_Arduino_Pin" && columns_comport.Length >= 6 && columns_comport.Substring(0, 3) == "_P0")
                        {
                            {
                                if (GlobalData.Arduino_outputFlag)
                                {
                                    GlobalData.Arduino_outputFlag = false;
                                    Arduino_IO_INPUT();
                                }
                                switch (columns_comport.Substring(3, 1))
                                {
                                    #region -- P02 --
                                    case "2":
                                        log.Debug("Arduino IO CMD: P02");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x01) == 0x00)
                                        {
                                            // if ( GlobalData.Arduino_IO_INPUT_value >= 0x100 )
                                            // { // to do 
                                            // }
                                            // else
                                            // {
                                            //   if ((GlobalData.Arduino_IO_INPUT_value & 0x01U)==0x00U))
                                            //   {
                                            //      if (columns_comport.Substring(5, 1) == "0")
                                            //      {
                                            //            // GPIO 0
                                            //      }
                                            //   }
                                            //   else // reading 1 != 00
                                            //   {
                                            //      if (columns_comport.Substring(5, 1) == "1")
                                            //      {
                                            //            // GPIO 1
                                            //      }
                                            //   }
                                            // } 

                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino2_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x01) == 0x01)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino2_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P03 --
                                    case "3":
                                        log.Debug("Arduino IO CMD: P03");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x02) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino3_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x02) == 0x02)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino3_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P04 --
                                    case "4":
                                        log.Debug("Arduino IO CMD: P04");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x04) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino4_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x04) == 0x04)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino4_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P05 --
                                    case "5":
                                        log.Debug("Arduino IO CMD: P05");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x08) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino5_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x08) == 0x08)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino5_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P06 --
                                    case "6":
                                        log.Debug("Arduino IO CMD: P06");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x10) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino6_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x10) == 0x10)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino6_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P07 --
                                    case "7":
                                        log.Debug("Arduino IO CMD: P07");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x20) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino7_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x20) == 0x20)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino7_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P08 --
                                    case "8":
                                        log.Debug("Arduino IO CMD: P08");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x40) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino8_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x40) == 0x40)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino8_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                    #endregion

                                    #region -- P09 --
                                    case "9":
                                        log.Debug("Arduino IO CMD: P09");
                                        if (columns_comport.Substring(5, 1) == "0" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x80) == 0x00)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino9_0_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else if (columns_comport.Substring(5, 1) == "1" &&
                                            GlobalData.Arduino_IO_INPUT_value < 0x100 &&
                                            Convert.ToByte(GlobalData.Arduino_IO_INPUT_value & 0x80) == 0x80)
                                        {
                                            if (columns_serial == "_accumulate")
                                            {
                                                GlobalData.IO_Arduino9_1_COUNT++;
                                                label_Command.Text = "IO CMD_ACCUMULATE";
                                            }
                                            else
                                            {
                                                IO_CMD();
                                            }
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        break;
                                        #endregion
                                }
                            }
                        }
                        #endregion

                        #region -- NI IO Input --
                        /*
                        else if (columns_command.Length >= 13 && columns_command.Substring(0, 11) == "_EXT_Input_")
                        {
                            switch (columns_command.Substring(11, 2))
                            {
                                case "P0":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[0].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                                case "P1":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[1].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                                case "P2":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[2].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }
                        #endregion

                        #region -- NI IO Output --
                        else if (columns_command.Length >= 14 && columns_command.Substring(0, 12) == "_EXT_Output_")
                        {
                            switch (columns_command.Substring(12, 2))
                            {
                                case "P0":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[0].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                                case "P1":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[1].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                                case "P2":
                                    try
                                    {
                                        using (Task digitalWriteTask = new Task())
                                        {
                                            //  Create an Digital Output channel and name it.
                                            digitalWriteTask.DOChannels.CreateChannel(DaqSystem.Local.GetPhysicalChannels(PhysicalChannelTypes.DOPort, PhysicalChannelAccess.External)[2].ToString(), "port0",
                                                ChannelLineGrouping.OneChannelForAllLines);

                                            //  Write digital port data. WriteDigitalSingChanSingSampPort writes a single sample
                                            //  of digital data on demand, so no timeout is necessary.
                                            DigitalSingleChannelWriter writer = new DigitalSingleChannelWriter(digitalWriteTask.Stream);
                                            writer.WriteSingleSamplePort(true, (UInt32)Convert.ToUInt32(columns_times));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                    break;
                            }
                            label_Command.Text = "(" + columns_command + ") " + columns_times;
                        }*/
                        #endregion

                        #region -- Audio Debounce --
                        else if (columns_command == "_audio_debounce")
                        {
                            log.Debug("Audio Detect: _audio_debounce");
                            bool Debounce_Time_PB1, Debounce_Time_PB7;
                            if (columns_interval != "")
                            {
                                MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB1(Convert.ToUInt16(columns_interval));
                                MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB7(Convert.ToUInt16(columns_interval));
                                Debounce_Time_PB1 = MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB1(Convert.ToUInt16(columns_interval));
                                Debounce_Time_PB7 = MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB7(Convert.ToUInt16(columns_interval));
                            }
                            else
                            {
                                MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB1();
                                MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB7();
                                Debounce_Time_PB1 = MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB1();
                                Debounce_Time_PB7 = MyBlueRat.Set_Input_GPIO_Low_Debounce_Time_PB7();
                            }
                        }
                        #endregion

                        #region -- Keyword Search --
                        else if (columns_command == "_keyword")
                        {
                            switch (columns_times)
                            {
                                case "1":
                                    log.Debug("Keyword Search: 1");
                                    if (GlobalData.keyword_1 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_1 = "false";
                                    break;

                                case "2":
                                    log.Debug("Keyword Search: 2");
                                    if (GlobalData.keyword_2 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_2 = "false";
                                    break;

                                case "3":
                                    log.Debug("Keyword Search: 3");
                                    if (GlobalData.keyword_3 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_3 = "false";
                                    break;

                                case "4":
                                    log.Debug("Keyword Search: 4");
                                    if (GlobalData.keyword_4 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_4 = "false";
                                    break;

                                case "5":
                                    log.Debug("Keyword Search: 5");
                                    if (GlobalData.keyword_5 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_5 = "false";
                                    break;

                                case "6":
                                    log.Debug("Keyword Search: 6");
                                    if (GlobalData.keyword_6 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_6 = "false";
                                    break;

                                case "7":
                                    log.Debug("Keyword Search: 7");
                                    if (GlobalData.keyword_7 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_7 = "false";
                                    break;

                                case "8":
                                    log.Debug("Keyword Search: 8");
                                    if (GlobalData.keyword_8 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_8 = "false";
                                    break;

                                case "9":
                                    log.Debug("Keyword Search: 9");
                                    if (GlobalData.keyword_9 == "true")
                                    {
                                        KeywordCommand();
                                    }
                                    else
                                    {
                                        SysDelay = 0;
                                    }
                                    GlobalData.keyword_9 = "false";
                                    break;

                                default:
                                    log.Debug("Keyword Search: 10");
                                    if (columns_times == "10")
                                    {
                                        if (GlobalData.keyword_10 == "true")
                                        {
                                            KeywordCommand();
                                        }
                                        else
                                        {
                                            SysDelay = 0;
                                        }
                                        GlobalData.keyword_10 = "false";
                                    }
                                    log.Debug("keyword not found_schedule");
                                    break;

                            }
                        }
                        #endregion

                        #region -- PWM1 --
                        else if (columns_command == "_pwm1")
                        {
                            log.Debug("PWM Control: _pwm1");
                            if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                            {
                                string pwm_output;
                                int result = 0;
                                if (columns_serial == "off")
                                {
                                    pwm_output = "set pwm_output 0";
                                    PortA.WriteLine(pwm_output);
                                }
                                else if (columns_serial == "on")
                                {
                                    pwm_output = "set pwm_output 1";
                                    PortA.WriteLine(pwm_output);
                                }
                                else if (int.TryParse(columns_serial, out result) == true)
                                {
                                    if (int.Parse(columns_serial) >= 0 && int.Parse(columns_serial) <= 100)
                                    {
                                        pwm_output = "set pwm_percent " + columns_serial;
                                        PortA.WriteLine(pwm_output);
                                    }
                                }
                                else
                                {
                                    pwm_output = columns_serial;
                                    PortA.WriteLine(pwm_output);
                                }
                            }
                        }
                        #endregion

                        #region -- PWM2 --
                        else if (columns_command == "_pwm2")
                        {
                            log.Debug("PWM Control: _pwm2");
                            if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1")
                            {
                                string pwm_output;
                                int result = 0;
                                if (columns_serial == "off")
                                {
                                    pwm_output = "set pwm_output 0";
                                    PortB.WriteLine(pwm_output);
                                }
                                else if (columns_serial == "on")
                                {
                                    pwm_output = "set pwm_output 1";
                                    PortB.WriteLine(pwm_output);
                                }
                                else if (int.TryParse(columns_serial, out result) == true)
                                {
                                    if (int.Parse(columns_serial) >= 0 && int.Parse(columns_serial) <= 100)
                                    {
                                        pwm_output = "set pwm_percent " + columns_serial;
                                        PortB.WriteLine(pwm_output);
                                    }
                                }
                                else
                                {
                                    pwm_output = columns_serial;
                                    PortB.WriteLine(pwm_output);
                                }
                            }
                        }
                        #endregion

                        #region -- PWM3 --
                        else if (columns_command == "_pwm3")
                        {
                            log.Debug("PWM Control: _pwm3");
                            if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1")
                            {
                                string pwm_output;
                                int result = 0;
                                if (columns_serial == "off")
                                {
                                    pwm_output = "set pwm_output 0";
                                    PortB.WriteLine(pwm_output);
                                }
                                else if (columns_serial == "on")
                                {
                                    pwm_output = "set pwm_output 1";
                                    PortB.WriteLine(pwm_output);
                                }
                                else if (int.TryParse(columns_serial, out result) == true)
                                {
                                    if (int.Parse(columns_serial) >= 0 && int.Parse(columns_serial) <= 100)
                                    {
                                        pwm_output = "set pwm_percent " + columns_serial;
                                        PortB.WriteLine(pwm_output);
                                    }
                                }
                                else
                                {
                                    pwm_output = columns_serial;
                                    PortB.WriteLine(pwm_output);
                                }
                            }
                        }
                        #endregion

                        #region -- 遙控器指令 --
                        else
                        {
                            try
                            {
                                log.Debug("Remote Control: TV_rc_key");
                                for (int k = 0; k < stime; k++)
                                {
                                    label_Command.Text = columns_command;
                                    if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
                                    {
                                        //執行小紅鼠指令
                                        Autocommand_RedRat("Form1", columns_command);
                                    }
                                    else if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                                    {
                                        //執行小藍鼠指令
                                        Autocommand_BlueRat("Form1", columns_command);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Please connect AutoKit or RedRat!", "Redrat Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        button_Start.PerformClick();
                                    }
                                    videostring = columns_command;
                                    RedRatDBViewer_Delay(sRepeat);
                                }
                            }
                            catch (Exception Ex)
                            {
                                MessageBox.Show(Ex.Message.ToString(), "RCDB library fail !");
                            }
                        }
                        #endregion

                        #region -- Remark --
                        if (columns_remark != "")
                        {
                            label_Remark.Invoke((MethodInvoker)(() => label_Remark.Text = columns_remark));
                            //label_Remark.Text = columns_remark;
                        }
                        else
                        {
                            label_Remark.Text = "";
                        }
                        #endregion

                        //Thread MyExportText = new Thread(new ThreadStart(MyExportCamd));
                        //MyExportText.Start();
                        log.Debug("CloseTime record.");
                        ini12.INIWrite(MailPath, "Data Info", "CloseTime", string.Format("{0:R}", DateTime.Now));


                        if (GlobalData.Break_Out_Schedule == 1)//定時器時間到跳出迴圈//
                        {
                            log.Debug("Break schedule.");
                            j = GlobalData.Schedule_Loop;
                            UpdateUI(j.ToString(), label_LoopNumber_Value);
                            GlobalData.label_LoopNumber = j.ToString();
                            break;
                        }

                        Nowpoint = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Index;
                        log.Debug("Nowpoint record: " + Nowpoint);

                        if (Breakfunction == true)
                        {
                            log.Debug("Breakfunction.");
                            if (Breakpoint == Nowpoint)
                            {
                                log.Debug("Breakpoint = Nowpoint");
                                button_Pause.PerformClick();
                            }
                        }

                        if (Pause == true)//如果按下暫停鈕//
                        {
                            timer_countdown.Stop();
                            SchedulePause.WaitOne();
                            log.Debug("SchedulePause_WaitOne");
                        }
                        else
                        {
                            RedRatDBViewer_Delay(SysDelay);
                            log.Debug("RedRatDBViewer_Delay: " + SysDelay);
                        }

                        #region -- 足跡模式 --
                        //假如足跡模式打開則會append足跡上去
                        if (ini12.INIRead(MainSettingPath, "Record", "Footprint Mode", "") == "1" && SysDelay != 0)
                        {
                            string fName = ini12.INIRead(GlobalData.MainSettingPath, "Record", "LogPath", "");
                            log.Debug("IO Input Footprint Mode Start");
                            //檔案不存在則加入標題
                            if (File.Exists(fName + @"\StepRecord.csv") == false && ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                            {
                                File.AppendAllText(fName + @"\StepRecord.csv", "LOOP,TIME,COMMAND,PB07_Status,PB01_Status,PA15_Status,PA14_Status,PA11_Status,PA10_Status," +
                                    "PA10_0,PA10_1," +
                                    "PA11_0,PA11_1," +
                                    "PA14_0,PA14_1," +
                                    "PA15_0,PA15_1," +
                                    "PB1_0,PB1_1," +
                                    "PB7_0,PB7_1," +
                                    Environment.NewLine);

                                File.AppendAllText(fName + @"\StepRecord.csv",
                                GlobalData.Loop_Number + "," + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + label_Command.Text + "," + GlobalData.IO_INPUT +
                                "," + GlobalData.IO_PA10_0_COUNT + "," + GlobalData.IO_PA10_1_COUNT +
                                "," + GlobalData.IO_PA11_0_COUNT + "," + GlobalData.IO_PA11_1_COUNT +
                                "," + GlobalData.IO_PA14_0_COUNT + "," + GlobalData.IO_PA14_1_COUNT +
                                "," + GlobalData.IO_PA15_0_COUNT + "," + GlobalData.IO_PA15_1_COUNT +
                                "," + GlobalData.IO_PB1_0_COUNT + "," + GlobalData.IO_PB1_1_COUNT +
                                "," + GlobalData.IO_PB7_0_COUNT + "," + GlobalData.IO_PB7_1_COUNT + Environment.NewLine);
                            }
                            else
                            {
                                File.AppendAllText(fName + @"\StepRecord.csv",
                                GlobalData.Loop_Number + "," + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + label_Command.Text + "," + GlobalData.IO_INPUT +
                                "," + GlobalData.IO_PA10_0_COUNT + "," + GlobalData.IO_PA10_1_COUNT +
                                "," + GlobalData.IO_PA11_0_COUNT + "," + GlobalData.IO_PA11_1_COUNT +
                                "," + GlobalData.IO_PA14_0_COUNT + "," + GlobalData.IO_PA14_1_COUNT +
                                "," + GlobalData.IO_PA15_0_COUNT + "," + GlobalData.IO_PA15_1_COUNT +
                                "," + GlobalData.IO_PB1_0_COUNT + "," + GlobalData.IO_PB1_1_COUNT +
                                "," + GlobalData.IO_PB7_0_COUNT + "," + GlobalData.IO_PB7_1_COUNT + Environment.NewLine);
                            }
                            log.Debug("IO Input Footprint Mode Stop");

                            log.Debug("Arduino IO Input Footprint Mode Start");
                            //檔案不存在則加入標題
                            if (File.Exists(fName + @"\Arduino_StepRecord.csv") == false && GlobalData.m_Arduino_Port.IsOpen() == true)
                            {
                                File.AppendAllText(fName + @"\Arduino_StepRecord.csv", "LOOP,TIME,COMMAND,P09_Status,P08_Status,P07_Status,P06_Status,P05_Status,P04_Status,P03_Status,P02_Status," +
                                    "P09_0,P09_1," +
                                    "P08_0,P08_1," +
                                    "P07_0,P07_1," +
                                    "P06_0,P06_1," +
                                    "P05_0,P05_1," +
                                    "P04_0,P04_1," +
                                    "P03_0,P03_1," +
                                    "P02_0,P02_1," +
                                    Environment.NewLine);

                                File.AppendAllText(fName + @"\Arduino_StepRecord.csv",
                                GlobalData.Loop_Number + "," + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + label_Command.Text + "," + GlobalData.Arduino_IO_INPUT +
                                GlobalData.IO_Arduino9_0_COUNT + "," + GlobalData.IO_Arduino9_1_COUNT + "," +
                                GlobalData.IO_Arduino8_0_COUNT + "," + GlobalData.IO_Arduino8_1_COUNT + "," +
                                GlobalData.IO_Arduino7_0_COUNT + "," + GlobalData.IO_Arduino7_1_COUNT + "," +
                                GlobalData.IO_Arduino6_0_COUNT + "," + GlobalData.IO_Arduino6_1_COUNT + "," +
                                GlobalData.IO_Arduino5_0_COUNT + "," + GlobalData.IO_Arduino5_1_COUNT + "," +
                                GlobalData.IO_Arduino4_0_COUNT + "," + GlobalData.IO_Arduino4_1_COUNT + "," +
                                GlobalData.IO_Arduino3_0_COUNT + "," + GlobalData.IO_Arduino3_1_COUNT + "," +
                                GlobalData.IO_Arduino2_0_COUNT + "," + GlobalData.IO_Arduino2_1_COUNT + "," + Environment.NewLine);
                            }
                            else
                            {
                                File.AppendAllText(fName + @"\Arduino_StepRecord.csv",
                                GlobalData.Loop_Number + "," + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + label_Command.Text + "," + GlobalData.Arduino_IO_INPUT +
                                GlobalData.IO_Arduino9_0_COUNT + "," + GlobalData.IO_Arduino9_1_COUNT + "," +
                                GlobalData.IO_Arduino8_0_COUNT + "," + GlobalData.IO_Arduino8_1_COUNT + "," +
                                GlobalData.IO_Arduino7_0_COUNT + "," + GlobalData.IO_Arduino7_1_COUNT + "," +
                                GlobalData.IO_Arduino6_0_COUNT + "," + GlobalData.IO_Arduino6_1_COUNT + "," +
                                GlobalData.IO_Arduino5_0_COUNT + "," + GlobalData.IO_Arduino5_1_COUNT + "," +
                                GlobalData.IO_Arduino4_0_COUNT + "," + GlobalData.IO_Arduino4_1_COUNT + "," +
                                GlobalData.IO_Arduino3_0_COUNT + "," + GlobalData.IO_Arduino3_1_COUNT + "," +
                                GlobalData.IO_Arduino2_0_COUNT + "," + GlobalData.IO_Arduino2_1_COUNT + "," + Environment.NewLine);
                            }
                            log.Debug("Arduino IO Input Footprint Mode Stop");
                        }
                        #endregion
                        log.Debug("End.");
                    }

                    #region -- Import database --
                    if (ini12.INIRead(MainSettingPath, "Record", "ImportDB", "") == "1")
                    {
                        string SQLServerURL = "server=192.168.56.2\\ATMS;database=Autobox;uid=AS;pwd=AS";

                        SqlConnection conn = new SqlConnection(SQLServerURL);
                        conn.Open();
                        SqlCommand s_com = new SqlCommand
                        {
                            //s_com.CommandText = "select * from Autobox.dbo.testresult";
                            CommandText = "insert into Autobox.dbo.testresult (ab_p_id, ab_result, ab_create, ab_time, ab_loop, ab_loop_time, ab_loop_step, ab_root, ab_user) values ('" + label_LoopNumber_Value.Text + "', 'Pass', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + label_LoopNumber_Value.Text + "', 1, 21000, 2, 0, 'Joseph')",
                            //s_com.CommandText = "update Autobox.dbo.testresult (ab_result, ab_close, ab_time, ab_loop, ab_root, ab_user) values ('Pass', '" + DateTime.Now.ToString("HH:mm:ss") + "', '" + label1.Text + "', 1, 21000, 'Joseph')";
                            //s_com.CommandText = "Update Autobox.dbo.testresult set ab_result='Pass', ab_close='2014/5/21 15:49:35', ab_time=600000, ab_loop=25, ab_root=0 where ab_num=2";
                            //s_com.CommandText = "Update Autobox.dbo.testresult set ab_result='NG', ab_close='2014/5/21 15:59:35', ab_time=1200000, ab_loop=50, ab_root=1 where ab_num=3";

                            Connection = conn
                        };

                        SqlDataReader s_read = s_com.ExecuteReader();
                        try
                        {
                            while (s_read.Read())
                            {
                                Console.WriteLine("Log> Find {0}", s_read["ab_p_id"].ToString());
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                        s_read.Close();

                        conn.Close();
                    }
                    #endregion
                }

                log.Debug("Loop_Number: " + GlobalData.Loop_Number);

				DisposeRam();
                GlobalData.Loop_Number++;
            }

            #region -- Each video record when completed the schedule --
            if (ini12.INIRead(MainSettingPath, "Record", "EachVideo", "") == "1")
            {
                if (StartButtonPressed == true)
                {
                    if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                    {
                        if (GlobalData.VideoRecording == false)
                        {
                            label_Command.Text = "Record Video...";
                            Thread.Sleep(1500);
                            Mysvideo(); // 開新檔
                            GlobalData.VideoRecording = true;
                            Thread oThreadC = new Thread(new ThreadStart(MySrtCamd));
                            oThreadC.Start();
                            Thread.Sleep(60000); // 錄影60秒

                            GlobalData.VideoRecording = false;
                            Mysstop();
                            oThreadC.Abort();
                            Thread.Sleep(1500);
                            label_Command.Text = "Vdieo recording completely.";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Camera not exist", "Camera Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            #endregion

            #region Excel function
            /*
            if (ini12.INIRead(sPath, "Record", "CompareChoose", "") == "1" && excelstat == true)
            {
                string excelFile = ini12.INIRead(sPath, "Record", "ComparePath", "") + "\\SimilarityReport_" + GlobalData.Schedule_Num;

                try
                {
                    //另存活頁簿
                    wBook.SaveAs(excelFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + excelFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }

                //關閉活頁簿
                //wBook.Close(false, Type.Missing, Type.Missing);

                //關閉Excel
                excelApp.Quit();

                //釋放Excel資源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
                wBook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wSheet);
                wSheet = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wRange);
                wRange = null;

                GC.Collect();
                excelstat = false;

                //Console.Read();

                CloseExcel();
            }*/
            #endregion

            #region -- schedule 切換 --
            if (StartButtonPressed != false)
            {
                if (GlobalData.Schedule_2_Exist == 1 && GlobalData.Schedule_Number == 1)
                {
                    if (ini12.INIRead(MainSettingPath, "Schedule2", "OnTimeStart", "") == "1" && StartButtonPressed == true)       //定時器時間未到進入等待<<<<<<<<<<<<<<
                    {
                        if (GlobalData.Break_Out_Schedule == 0)
                        {
                            while (ini12.INIRead(MainSettingPath, "Schedule2", "Timer", "") != TimeLabel2.Text)
                            {
                                ScheduleWait.WaitOne();
                            }
                            ScheduleWait.Set();
                        }
                    }       //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ini12.INIWrite(MainSettingPath, "Schedule1", "OnTimeStart", "0");
                    button_Schedule2.PerformClick();
                    timer_duringShot.Stop();
                    MyRunCamd();
                }
                else if (
                    GlobalData.Schedule_3_Exist == 1 && GlobalData.Schedule_Number == 1 ||
                    GlobalData.Schedule_3_Exist == 1 && GlobalData.Schedule_Number == 2)
                {
                    if (ini12.INIRead(MainSettingPath, "Schedule3", "OnTimeStart", "") == "1" && StartButtonPressed == true)
                    {
                        if (GlobalData.Break_Out_Schedule == 0)
                        {
                            while (ini12.INIRead(MainSettingPath, "Schedule3", "Timer", "") != TimeLabel2.Text)
                            {
                                ScheduleWait.WaitOne();
                            }
                            ScheduleWait.Set();
                        }
                    }
                    ini12.INIWrite(MainSettingPath, "Schedule2", "OnTimeStart", "0");
                    button_Schedule3.PerformClick();
                    timer_duringShot.Stop();
                    MyRunCamd();
                }
                else if (
                    GlobalData.Schedule_4_Exist == 1 && GlobalData.Schedule_Number == 1 ||
                    GlobalData.Schedule_4_Exist == 1 && GlobalData.Schedule_Number == 2 ||
                    GlobalData.Schedule_4_Exist == 1 && GlobalData.Schedule_Number == 3)
                {
                    if (ini12.INIRead(MainSettingPath, "Schedule4", "OnTimeStart", "") == "1" && StartButtonPressed == true)
                    {
                        if (GlobalData.Break_Out_Schedule == 0)
                        {
                            while (ini12.INIRead(MainSettingPath, "Schedule4", "Timer", "") != TimeLabel2.Text)
                            {
                                ScheduleWait.WaitOne();
                            }
                            ScheduleWait.Set();
                        }
                    }
                    ini12.INIWrite(MainSettingPath, "Schedule3", "OnTimeStart", "0");
                    button_Schedule4.PerformClick();
                    timer_duringShot.Stop();
                    MyRunCamd();
                }
                else if (
                    GlobalData.Schedule_5_Exist == 1 && GlobalData.Schedule_Number == 1 ||
                    GlobalData.Schedule_5_Exist == 1 && GlobalData.Schedule_Number == 2 ||
                    GlobalData.Schedule_5_Exist == 1 && GlobalData.Schedule_Number == 3 ||
                    GlobalData.Schedule_5_Exist == 1 && GlobalData.Schedule_Number == 4)
                {
                    if (ini12.INIRead(MainSettingPath, "Schedule5", "OnTimeStart", "") == "1" && StartButtonPressed == true)
                    {
                        if (GlobalData.Break_Out_Schedule == 0)
                        {
                            while (ini12.INIRead(MainSettingPath, "Schedule5", "Timer", "") != TimeLabel2.Text)
                            {
                                ScheduleWait.WaitOne();
                            }
                            ScheduleWait.Set();
                        }
                    }
                    ini12.INIWrite(MainSettingPath, "Schedule4", "OnTimeStart", "0");
                    button_Schedule5.PerformClick();
                    timer_duringShot.Stop();
                    MyRunCamd();
                }
            }
            #endregion

            //全部schedule跑完或是按下stop鍵以後會跑以下這段/////////////////////////////////////////
            if (StartButtonPressed == false)//按下STOP讓schedule結束//
            {
                GlobalData.Break_Out_MyRunCamd = 1;
                ini12.INIWrite(MailPath, "Data Info", "CloseTime", string.Format("{0:R}", DateTime.Now));
                UpdateUI("START", button_Start);
                button_Start.Enabled = true;
                button_Setting.Enabled = true;
                button_Pause.Enabled = false;
                button_SaveSchedule.Enabled = true;
                //setStyle();

                if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                {
                    _captureInProgress = false;
                    OnOffCamera();
                    //button_VirtualRC.Enabled = true;
                }

                /*
                if (Directory.Exists(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + GlobalData.Schedule_Num + "_Original") == true)
                {
                    DirectoryInfo DIFO = new DirectoryInfo(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + GlobalData.Schedule_Num + "_Original");
                    DIFO.Delete(true);
                }
                */
            }
            else//schedule自動跑完//
            {
                StartButtonPressed = false;
                UpdateUI("START", button_Start);
                button_Setting.Enabled = true;
                button_Pause.Enabled = false;
                button_SaveSchedule.Enabled = true;
                button_Start.Enabled = true;

                if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                {
                    _captureInProgress = false;
                    OnOffCamera();
                }

                GlobalData.Total_Test_Time = GlobalData.Schedule_1_TestTime + GlobalData.Schedule_2_TestTime + GlobalData.Schedule_3_TestTime + GlobalData.Schedule_4_TestTime + GlobalData.Schedule_5_TestTime;
                ConvertToRealTime(GlobalData.Total_Test_Time);
                if (ini12.INIRead(MailPath, "Send Mail", "value", "") == "1")
                {
                    GlobalData.Loop_Number = GlobalData.Loop_Number - 1;
                    FormMail FormMail = new FormMail();
                    FormMail.send();
                }
            }

            label_Command.Text = "Completed!";
            label_Remark.Text = "";
            ini12.INIWrite(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "OnTimeStart", "0");
            button_Schedule1.PerformClick();
            timer_countdown.Stop();
            timer_duringShot.Stop();
            CloseDtplay();

            //如果serialport開著則先關閉//
            /*
            if (PortA.IsOpen == true)
            {
                CloseSerialPort("A");
            }
            */
            if (GlobalData.m_SerialPort_A.IsOpen())
                //serialPortA.ClosePort().Handle
                GlobalData.m_SerialPort_A.ClosePort();
            if (GlobalData.m_SerialPort_B.IsOpen())
                GlobalData.m_SerialPort_B.ClosePort();
            if (GlobalData.m_SerialPort_C.IsOpen())
                GlobalData.m_SerialPort_C.ClosePort();
            if (GlobalData.m_SerialPort_D.IsOpen())
                GlobalData.m_SerialPort_D.ClosePort();
            if (GlobalData.m_SerialPort_E.IsOpen())
                GlobalData.m_SerialPort_E.ClosePort();
            if (MySerialPort.IsPortOpened() == true)
            {
                //CloseSerialPort("kline");
                MySerialPort.ClosePort();
            }
            if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1")
            {
                if (Can_Usb2C.Connect() == 1)
                {
                    Can_Usb2C.StopCAN();
                    Can_Usb2C.Disconnect();
                }
            }

            timeCount = GlobalData.Schedule_1_TestTime;
            ConvertToRealTime(timeCount);
            GlobalData.Scheduler_Row = 0;
            //setStyle();
        }
        #endregion

        bool PauseFlag = false;
        bool ShotFlag = false;
        #region -- IO CMD 指令集 --
        private void IO_CMD()
        {
            string columns_serial = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[6].Value.ToString();
            string columns_switch = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[7].Value.ToString();
            if (columns_serial == "_pause")
            {
                PauseFlag = true;
                button_Pause.PerformClick();
                label_Command.Text = "IO CMD_PAUSE";
            }
            else if (columns_serial == "_stop")
            {
                button_Start.PerformClick();
                label_Command.Text = "IO CMD_STOP";
            }
            else if (columns_serial == "_ac_restart")
            {
                GP0_GP1_AC_OFF_ON();
                Thread.Sleep(10);
                GP0_GP1_AC_OFF_ON();
                label_Command.Text = "IO CMD_AC_RESTART";
            }
            else if (columns_serial == "_shot")
            {
                ShotFlag = true;
                GlobalData.caption_Num++;
                if (GlobalData.Loop_Number == 1)
                    GlobalData.caption_Sum = GlobalData.caption_Num;
                Jes();
                label_Command.Text = "IO CMD_SHOT";
            }
            else if (columns_serial == "_mail")
            {
                if (ini12.INIRead(MailPath, "Send Mail", "value", "") == "1")
                {
                    GlobalData.Pass_Or_Fail = "NG";
                    FormMail FormMail = new FormMail();
                    FormMail.send();
                    label_Command.Text = "IO CMD_MAIL";
                }
                else
                {
                    MessageBox.Show("Please enable Mail Function in Settings.");
                }
            }
            else if (columns_serial.Substring(0, 3) == "_rc")
            {
                String rc_key = columns_serial;
                int startIndex = 4;
                int length = rc_key.Length - 4;
                String rc_key_substring = rc_key.Substring(startIndex, length);

                if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
                {
                    Autocommand_RedRat("Form1", rc_key_substring);
                    label_Command.Text = rc_key_substring;
                }
                else if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                {
                    Autocommand_BlueRat("Form1", rc_key_substring);
                    label_Command.Text = rc_key_substring;
                }
            }
            else if (columns_serial.Substring(0, 7) == "_logcmd")
            {
                String log_cmd = columns_serial;
                String log_newline = columns_switch;
                int startIndex = 10;
                int length = log_cmd.Length - 10;
                String log_cmd_substring = log_cmd.Substring(startIndex, length);
                String log_cmd_serialport = log_cmd.Substring(8, 1);

                if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" && log_cmd_serialport == "A")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_A, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" && log_cmd_serialport == "B")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_B, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1" && log_cmd_serialport == "C")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_C, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1" && log_cmd_serialport == "D")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_D, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1" && log_cmd_serialport == "E")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_E, log_cmd_substring, log_newline);
                }
                else if (log_cmd_serialport == "O")
                {
                    if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_A, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_B, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_C, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_D, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_E, log_cmd_substring, log_newline);
                }
            }
        }
        #endregion

        #region -- KEYWORD 指令集 --
        private void KeywordCommand()
        {
            string columns_serial = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[6].Value.ToString();
            string columns_switch = DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[7].Value.ToString();

            if (columns_serial == "_pause")
            {
                button_Pause.PerformClick();
                label_Command.Text = "KEYWORD_PAUSE";
            }
            else if (columns_serial == "_stop")
            {
                button_Start.PerformClick();
                label_Command.Text = "KEYWORD_STOP";
            }
            else if (columns_serial == "_ac_restart")
            {
                GP0_GP1_AC_OFF_ON();
                Thread.Sleep(10);
                GP0_GP1_AC_OFF_ON();
                label_Command.Text = "KEYWORD_AC_RESTART";
            }
            else if (columns_serial == "_shot")
            {
                GlobalData.caption_Num++;
                if (GlobalData.Loop_Number == 1)
                    GlobalData.caption_Sum = GlobalData.caption_Num;
                Jes();
                label_Command.Text = "KEYWORD_SHOT";
            }
            else if (columns_serial == "_mail")
            {
                if (ini12.INIRead(MailPath, "Send Mail", "value", "") == "1")
                {
                    GlobalData.Pass_Or_Fail = "NG";
                    FormMail FormMail = new FormMail();
                    FormMail.send();
                    label_Command.Text = "KEYWORD_MAIL";
                }
                else
                {
                    MessageBox.Show("Please enable Mail Function in Settings.");
                }
            }
            else if (columns_serial.Substring(0, 7) == "_savelog")
            {
                string fName = "";
                fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                String savelog_serialport = columns_serial.Substring(9, 1);
                if (savelog_serialport == "A")
                {
                    string t = fName + "\\_SaveLogA_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logA_text);
                    MYFILE.Close();
                }
                else if (savelog_serialport == "B")
                {
                    string t = fName + "\\_SaveLogB_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logB_text);
                    MYFILE.Close();
                }
                else if (savelog_serialport == "C")
                {
                    string t = fName + "\\_SaveLogC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logC_text);
                    MYFILE.Close();
                }
                else if (savelog_serialport == "D")
                {
                    string t = fName + "\\_SaveLogD_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logD_text);
                    MYFILE.Close();
                }
                else if (savelog_serialport == "E")
                {
                    string t = fName + "\\_SaveLogE_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logE_text);
                    MYFILE.Close();
                }
                else if (savelog_serialport == "O")
                {
                    string t = fName + "\\_SaveLogAll_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logAll_text);
                    MYFILE.Close();
                }
                label_Command.Text = "KEYWORD_SAVELOG";
            }
            else if (columns_serial.Substring(0, 3) == "_rc")
            {
                String rc_key = columns_serial;
                int startIndex = 4;
                int length = rc_key.Length - 4;
                String rc_key_substring = rc_key.Substring(startIndex, length);

                if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
                {
                    Autocommand_RedRat("Form1", rc_key_substring);
                    label_Command.Text = rc_key_substring;
                }
                else if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                {
                    Autocommand_BlueRat("Form1", rc_key_substring);
                    label_Command.Text = rc_key_substring;
                }
            }
            else if (columns_serial.Substring(0, 7) == "_logcmd")
            {
                String log_cmd = columns_serial;
                String log_newline = columns_switch;
                int startIndex = 10;
                int length = log_cmd.Length - 10;
                String log_cmd_substring = log_cmd.Substring(startIndex, length);
                String log_cmd_serialport = log_cmd.Substring(8, 1);

                if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" && log_cmd_serialport == "A")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_A, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" && log_cmd_serialport == "B")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_B, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1" && log_cmd_serialport == "C")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_C, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1" && log_cmd_serialport == "D")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_D, log_cmd_substring, log_newline);
                }
                else if (ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1" && log_cmd_serialport == "E")
                {
                    ReplaceNewLine(GlobalData.m_SerialPort_E, log_cmd_substring, log_newline);
                }
                else if (log_cmd_serialport == "O")
                {
                    if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_A, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_B, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_C, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port D", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_D, log_cmd_substring, log_newline);
                    if (ini12.INIRead(MainSettingPath, "Port E", "Checked", "") == "1")
                        ReplaceNewLine(GlobalData.m_SerialPort_E, log_cmd_substring, log_newline);
                }
                label_Command.Text = "KEYWORD_LOGCMD";
            }
        }
        #endregion

        #region -- 圖片比對 --
        private void MyCompareCamd()
        {
            //String fNameAll = "";
            //String fNameNG = "";
            /*            
            int i, j = 1;
            int TotalDelay = 0;

            switch (GlobalData.Schedule_Num)
            {
                case 1:
                    TotalDelay = (Convert.ToInt32(GlobalData.Schedule_Num1_Time) / GlobalData.Schedule_Loop);
                    break;
                case 2:
                    TotalDelay = (Convert.ToInt32(GlobalData.Schedule_Num2_Time) / GlobalData.Schedule_Loop);
                    break;
                case 3:
                    TotalDelay = (Convert.ToInt32(GlobalData.Schedule_Num3_Time) / GlobalData.Schedule_Loop);
                    break;
                case 4:
                    TotalDelay = (Convert.ToInt32(GlobalData.Schedule_Num4_Time) / GlobalData.Schedule_Loop);
                    break;
                case 5:
                    TotalDelay = (Convert.ToInt32(GlobalData.Schedule_Num5_Time) / GlobalData.Schedule_Loop);
                    break;
            }       //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


            //float[,] ReferenceResult = new float[GlobalData.Schedule_Loop, GlobalData.caption_Sum + 1];
            //float[] MeanValue = new float[GlobalData.Schedule_Loop];
            //int[] TotalValue = new int[GlobalData.Schedule_Loop];
            */
            //string ngPath = ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + GlobalData.Schedule_Num + "_NG\\";
            string comparePath = ini12.INIRead(MainSettingPath, "Record", "ComparePath", "") + "\\";
            string csvFile = comparePath + "SimilarityReport_" + GlobalData.Schedule_Number + ".csv";

            //Console.WriteLine("Loop Number: " + GlobalData.loop_Num);

            // 讀取ini中的路徑
            //fNameNG = ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + GlobalData.Schedule_Num + "_NG\\";

            string pathCompare1 = comparePath + "cf-" + GlobalData.Loop_Number + "_" + GlobalData.caption_Num + ".png";
            string pathCompare2 = comparePath + "cf-" + (GlobalData.Loop_Number - 1) + "_" + GlobalData.caption_Num + ".png";
            if (GlobalData.caption_Num == 0)
            {
                Console.WriteLine("Path Compare1: " + pathCompare1);
                Console.WriteLine("Path Compare2: " + pathCompare2);
            }
            if (System.IO.File.Exists(pathCompare1) && System.IO.File.Exists(pathCompare2))
            {
                string oHashCode = ImageHelper.produceFingerPrint(pathCompare1);
                string nHashCode = ImageHelper.produceFingerPrint(pathCompare2);
                int difference = ImageHelper.hammingDistance(oHashCode, nHashCode);
                int differenceNum = Convert.ToInt32(ini12.INIRead(MainSettingPath, "Record", "CompareDifferent", ""));
                string differencePercent = "";

                if (difference == 0)
                {
                    differencePercent = "100%";
                }
                else if (difference <= 10)
                {
                    differencePercent = "90%";
                }
                else if (difference <= 20)
                {
                    differencePercent = "80%";
                }
                else if (difference <= 30)
                {
                    differencePercent = "70%";
                }
                else if (difference <= 40)
                {
                    differencePercent = "60%";
                }
                else if (difference <= 50)
                {
                    differencePercent = "50%";
                }
                else if (difference <= 60)
                {
                    differencePercent = "40%";
                }
                else if (difference <= 70)
                {
                    differencePercent = "30%";
                }
                else if (difference <= 80)
                {
                    differencePercent = "20%";
                }
                else if (difference <= 90)
                {
                    differencePercent = "10%";
                }
                else
                {
                    differencePercent = "0%";
                }
                // 匯出csv記錄檔
                StreamWriter sw = new StreamWriter(csvFile, true);

                // 比對值設定
                GlobalData.excel_Num++;
                if (difference > differenceNum)
                {
                    GlobalData.NGValue[GlobalData.caption_Num]++;
                    GlobalData.NGRateValue[GlobalData.caption_Num] = (float)GlobalData.NGValue[GlobalData.caption_Num] / (GlobalData.Loop_Number - 1);

                    /*
                                        string[] FileList = System.IO.Directory.GetFiles(fNameAll, "cf-" + GlobalData.loop_Num + "_" + GlobalData.caption_Num + ".png");
                                        foreach (string File in FileList)
                                        {
                                            System.IO.FileInfo fi = new System.IO.FileInfo(File);
                                            fi.CopyTo(fNameNG + fi.Name);
                                        }
                    */

                    GlobalData.NGRateValue[GlobalData.caption_Num] = (float)GlobalData.NGValue[GlobalData.caption_Num] / (GlobalData.Loop_Number - 1);

                    /*
                    #region Excel function
                    try
                    {
                        // 引用第一個工作表
                        wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                        // 命名工作表的名稱
                        wSheet.Name = "全部測試資料";

                        // 設定工作表焦點
                        wSheet.Activate();

                        // 設定第n列資料
                        excelApp.Cells[GlobalData.excel_Num, 1] = " " + (GlobalData.loop_Num - 1) + "-" + GlobalData.caption_Num;
                        wSheet.Hyperlinks.Add(excelApp.Cells[GlobalData.excel_Num, 1], "cf-" + (GlobalData.loop_Num - 1) + "_" + GlobalData.caption_Num + ".png", Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Cells[GlobalData.excel_Num, 2] = " " + (GlobalData.loop_Num) + "-" + GlobalData.caption_Num;
                        wSheet.Hyperlinks.Add(excelApp.Cells[GlobalData.excel_Num, 2], "cf-" + (GlobalData.loop_Num) + "_" + GlobalData.caption_Num + ".png", Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Cells[GlobalData.excel_Num, 3] = differencePercent;
                        excelApp.Cells[GlobalData.excel_Num, 4] = GlobalData.NGValue[GlobalData.caption_Num];
                        excelApp.Cells[GlobalData.excel_Num, 5] = GlobalData.NGRateValue[GlobalData.caption_Num];
                        excelApp.Cells[GlobalData.excel_Num, 6] = "NG";

                        // 設定第n列顏色
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 1], wSheet.Cells[GlobalData.excel_Num, 2]];
                        wRange.Select();
                        wRange.Font.Color = ColorTranslator.ToOle(Color.Blue);
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 3], wSheet.Cells[GlobalData.excel_Num, 6]];
                        wRange.Select();
                        wRange.Font.Color = ColorTranslator.ToOle(Color.Red);

                        // 自動調整欄寬
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 1], wSheet.Cells[GlobalData.excel_Num, 6]];
                        wRange.EntireRow.AutoFit();
                        wRange.EntireColumn.AutoFit();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
                    }
                    #endregion
                    */

                    sw.Write("=hyperlink(\"cf-" + (GlobalData.Loop_Number - 1) + "_" + GlobalData.caption_Num + ".png\"，\"" + (GlobalData.Loop_Number - 1) + "-" + GlobalData.caption_Num + "\")" + ",");
                    sw.Write("=hyperlink(\"cf-" + (GlobalData.Loop_Number) + "_" + GlobalData.caption_Num + ".png\"，\"" + (GlobalData.Loop_Number) + "-" + GlobalData.caption_Num + "\")" + ",");
                    sw.Write(differencePercent + ",");
                    sw.Write(GlobalData.NGValue[GlobalData.caption_Num] + ",");
                    sw.Write(GlobalData.NGRateValue[GlobalData.caption_Num] + ",");
                    sw.WriteLine("NG");
                }
                else
                {
                    GlobalData.NGRateValue[GlobalData.caption_Num] = (float)GlobalData.NGValue[GlobalData.caption_Num] / (GlobalData.Loop_Number - 1);

                    /*
                    #region Excel function
                    try
                    {
                        // 引用第一個工作表
                        wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                        // 命名工作表的名稱
                        wSheet.Name = "全部測試資料";

                        // 設定工作表焦點
                        wSheet.Activate();

                        // 設定第n列資料
                        excelApp.Cells[GlobalData.excel_Num, 1] = " " + (GlobalData.loop_Num - 1) + "-" + GlobalData.caption_Num;
                        wSheet.Hyperlinks.Add(excelApp.Cells[GlobalData.excel_Num, 1], "cf-" + (GlobalData.loop_Num - 1) + "_" + GlobalData.caption_Num + ".png", Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Cells[GlobalData.excel_Num, 2] = " " + (GlobalData.loop_Num) + "-" + GlobalData.caption_Num;
                        wSheet.Hyperlinks.Add(excelApp.Cells[GlobalData.excel_Num, 2], "cf-" + (GlobalData.loop_Num) + "_" + GlobalData.caption_Num + ".png", Type.Missing, Type.Missing, Type.Missing);
                        excelApp.Cells[GlobalData.excel_Num, 3] = differencePercent;
                        excelApp.Cells[GlobalData.excel_Num, 4] = GlobalData.NGValue[GlobalData.caption_Num];
                        excelApp.Cells[GlobalData.excel_Num, 5] = GlobalData.NGRateValue[GlobalData.caption_Num];
                        excelApp.Cells[GlobalData.excel_Num, 6] = "Pass";

                        // 設定第n列顏色
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 1], wSheet.Cells[GlobalData.excel_Num, 2]];
                        wRange.Select();
                        wRange.Font.Color = ColorTranslator.ToOle(Color.Blue);
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 3], wSheet.Cells[GlobalData.excel_Num, 6]];
                        wRange.Select();
                        wRange.Font.Color = ColorTranslator.ToOle(Color.Green);

                        // 自動調整欄寬
                        wRange = wSheet.Range[wSheet.Cells[GlobalData.excel_Num, 1], wSheet.Cells[GlobalData.excel_Num, 6]];
                        wRange.EntireRow.AutoFit();
                        wRange.EntireColumn.AutoFit();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
                    }
                    #endregion
                    */

                    sw.Write("=hyperlink(\"cf-" + (GlobalData.Loop_Number - 1) + "_" + GlobalData.caption_Num + ".png\"，\"" + (GlobalData.Loop_Number - 1) + "-" + GlobalData.caption_Num + "\")" + ",");
                    sw.Write("=hyperlink(\"cf-" + (GlobalData.Loop_Number) + "_" + GlobalData.caption_Num + ".png\"，\"" + (GlobalData.Loop_Number) + "-" + GlobalData.caption_Num + "\")" + ",");
                    sw.Write(differencePercent + ",");
                    sw.Write(GlobalData.NGValue[GlobalData.caption_Num] + ",");
                    sw.Write(GlobalData.NGRateValue[GlobalData.caption_Num] + ",");
                    sw.WriteLine("Pass");
                }
                sw.Close();

                /*
                Bitmap picCompare1 = (Bitmap)Image.FromFile(pathCompare1);
                Bitmap picCompare2 = (Bitmap)Image.FromFile(pathCompare2);
                float CompareValue = Similarity(picCompare1, picCompare2);
                ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num] = CompareValue;
                Console.WriteLine("Reference(" + (GlobalData.loop_Num - 1) + "," + GlobalData.caption_Num + ") = " + ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num]);

                GlobalData.SumValue[GlobalData.caption_Num] = GlobalData.SumValue[GlobalData.caption_Num] + ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num];
                Console.WriteLine("SumValue" + GlobalData.caption_Num + " = " + GlobalData.SumValue[GlobalData.caption_Num]);

                MeanValue[GlobalData.caption_Num] = GlobalData.SumValue[GlobalData.caption_Num] / (GlobalData.loop_Num - 1);
                Console.WriteLine("MeanValue" + GlobalData.caption_Num + " = " + MeanValue[GlobalData.caption_Num]);

                for (i = GlobalData.loop_Num - 11; i < GlobalData.loop_Num - 1; i++)
                {
                    for (j = 1; j < GlobalData.caption_Sum + 1; j++)
                    {
                        string pathCompare1 = fNameAll + "cf-" + i + "_" + j + ".png";
                        string pathCompare2 = fNameAll + "cf-" + (i - 1) + "_" + j + ".png";
                        Bitmap picCompare1 = (Bitmap)Image.FromFile(pathCompare1);
                        Bitmap picCompare2 = (Bitmap)Image.FromFile(pathCompare2);
                        float CompareValue = Similarity(picCompare1, picCompare2);
                        ReferenceResult[i, j] = CompareValue;
                        Console.WriteLine("Reference(" + i + "," + j + ") = " + ReferenceResult[i, j]);

                        //int[] GetHisogram1 = GetHisogram(picCompare1);
                        //int[] GetHisogram2 = GetHisogram(picCompare2);
                        //float CompareResult = GetResult(GetHisogram1, GetHisogram2);

                        //long[] GetHistogram1 = GetHistogram(picCompare1);
                        //long[] GetHistogram2 = GetHistogram(picCompare2);
                        //float CompareResult = GetResult(GetHistogram1, GetHistogram2);

                    }
                    //Thread.Sleep(TotalDelay);
                }

                for (j = 1; j < GlobalData.caption_Sum + 1; j++)
                {
                    for (i = 1; i < GlobalData.loop_Num - 1; i++)
                    {
                        SumValue[j] = SumValue[j] + ReferenceResult[i, j];
                        TotalValue[j]++;
                        //Console.WriteLine("SumValue" + j + " = " + SumValue[j]);
                    }
                    //Thread.Sleep(TotalDelay);
                    MeanValue[j] = SumValue[j] / (GlobalData.loop_Num - 2);
                    //Console.WriteLine("MeanValue" + j + " = " + MeanValue[j]);
                }

                StreamWriter sw = new StreamWriter(csvFile, true);
                if (GlobalData.loop_Num == 2 && GlobalData.caption_Num == 1)
                    sw.WriteLine("Point(X), Point(Y), MeanValue, Reference, NGValue, TotalValue, NGRate, Test Result");

                if (ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num] > (MeanValue[GlobalData.caption_Num] + 0.5) || ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num] < (MeanValue[GlobalData.caption_Num] - 0.5))
                {
                    GlobalData.NGValue[GlobalData.caption_Num]++;
                    GlobalData.NGRateValue[GlobalData.caption_Num] = (float)GlobalData.NGValue[GlobalData.caption_Num] / GlobalData.loop_Num;
                    string[] FileList = System.IO.Directory.GetFiles(fNameAll, "cf-" + GlobalData.loop_Num + "_" + GlobalData.caption_Num + ".png");
                    foreach (string File in FileList)
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(fNameNG + fi.Name);
                    }
                    sw.Write((GlobalData.loop_Num - 1) + ", " + GlobalData.caption_Num + ", ");
                    sw.Write(MeanValue[GlobalData.caption_Num] + ", ");
                    sw.Write(ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num] + ", ");
                    sw.Write(GlobalData.NGValue[GlobalData.caption_Num] + ", ");
                    sw.Write(GlobalData.loop_Num + ", ");
                    sw.Write(GlobalData.NGRateValue[GlobalData.caption_Num] + ", ");
                    sw.WriteLine("NG");
                }
                else
                {
                    GlobalData.NGRateValue[GlobalData.caption_Num] = (float)GlobalData.NGValue[GlobalData.caption_Num] / GlobalData.loop_Num;
                    sw.Write((GlobalData.loop_Num - 1) + ", " + GlobalData.caption_Num + ", ");
                    sw.Write(MeanValue[GlobalData.caption_Num] + ", ");
                    sw.Write(ReferenceResult[(GlobalData.loop_Num - 1), GlobalData.caption_Num] + ", ");
                    sw.Write(GlobalData.NGValue[GlobalData.caption_Num] + ", ");
                    sw.Write(GlobalData.loop_Num + ", ");
                    sw.Write(GlobalData.NGRateValue[GlobalData.caption_Num] + ", ");
                    sw.WriteLine("Pass");
                }
                sw.Close();

                RedratLable.Text = "End Compare Picture.";
                */
            }
        }
        #endregion

        #region -- 拍照 --
        private void Jes() => Invoke(new EventHandler(delegate { Myshot(); }));

        private void Myshot()
        {
            log.Debug("Myshot: Start");
            button_Start.Enabled = false;
            if (TakePicture == true && TakePictureError <= 3)
                TakePictureError++;
            else if (TakePicture == true && TakePictureError > 3)
                MessageBox.Show("Please check the Camera!", "Camera take picture Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                TakePicture = true;
                //setStyle();
                capture.FrameEvent2 += new Capture.HeFrame(CaptureDone);
                capture.GrapImg();
                log.Debug("Myshot: Myshot");
            }
        }

        // 複製原始圖片
        protected Bitmap CloneBitmap(Bitmap source)
        {
            log.Debug("Myshot: CloneBitmap");
            return new Bitmap(source);
        }

        private void CaptureDone(System.Drawing.Bitmap e)
        {
            log.Debug("Myshot: CaptureDone");
            capture.FrameEvent2 -= new Capture.HeFrame(CaptureDone);
            string fName = ini12.INIRead(MainSettingPath, "Record", "VideoPath", "");
            //string ngFolder = "Schedule" + GlobalData.Schedule_Num + "_NG";

            //圖片印字
            Bitmap newBitmap = CloneBitmap(e);
            //newBitmap = CloneBitmap(e);
            pictureBox4.Image = newBitmap;

            if (ini12.INIRead(MainSettingPath, "Record", "CompareChoose", "") == "1")
            {
                // Create Compare folder
                string comparePath = ini12.INIRead(MainSettingPath, "Record", "ComparePath", "");
                //string ngPath = fName + "\\" + ngFolder;
                string compareFile = comparePath + "\\" + "cf-" + GlobalData.Loop_Number + "_" + GlobalData.caption_Num + ".png";
                if (GlobalData.caption_Num == 0)
                    GlobalData.caption_Num++;
                /*
                if (Directory.Exists(ngPath))
                {

                }
                else
                {
                    Directory.CreateDirectory(ngPath);
                }
                */
                // 圖片比較

                /*
                newBitmap = CloneBitmap(e);
                newBitmap = RGB2Gray(newBitmap);
                newBitmap = ConvertTo1Bpp2(newBitmap);
                newBitmap = SobelEdgeDetect(newBitmap);                
                this.pictureBox4.Image = newBitmap;
                */
                pictureBox4.Image.Save(compareFile, ImageFormat.Png);
                if (GlobalData.Loop_Number < 2)
                {

                }
                else
                {
                    Thread MyCompareThread = new Thread(new ThreadStart(MyCompareCamd));
                    MyCompareThread.Start();
                }
            }

            Graphics bitMap_g = Graphics.FromImage(pictureBox4.Image);//底圖
            Font Font = new Font("Microsoft JhengHei Light", 16, FontStyle.Bold);
            Brush FontColor = new SolidBrush(Color.Red);
            string[] Resolution = ini12.INIRead(MainSettingPath, "Camera", "Resolution", "").Split('*');
            int YPoint = int.Parse(Resolution[1]);

            //照片印上現在步驟//
            if (DataGridView_Schedule.Rows[GlobalData.Schedule_Step].Cells[9].Value.ToString() != "")           //含有remark value
            {
                bitMap_g.DrawString(DataGridView_Schedule.Rows[GlobalData.Schedule_Step].Cells[9].Value.ToString(),
                                Font,
                                FontColor,
                                new PointF(5, YPoint - 120));
                bitMap_g.DrawString(DataGridView_Schedule.Rows[GlobalData.Schedule_Step].Cells[0].Value.ToString() + "  ( " + label_Command.Text + " )",
                                Font,
                                FontColor,
                                new PointF(5, YPoint - 80));
            }
            else
            {
                bitMap_g.DrawString(DataGridView_Schedule.Rows[GlobalData.Schedule_Step].Cells[0].Value.ToString() + "  ( " + label_Command.Text + " )",
                Font,
                FontColor,
                new PointF(5, YPoint - 80));
            }
            //照片印上現在時間//
            bitMap_g.DrawString(TimeLabel.Text,
                                Font,
                                FontColor,
                                new PointF(5, YPoint - 40));

            Font.Dispose();
            FontColor.Dispose();
            bitMap_g.Dispose();

            string t = fName + "\\" + "pic-" + DateTime.Now.ToString("yyyyMMddHHmmss") + "(" + label_LoopNumber_Value.Text + "-" + GlobalData.caption_Num + ").jpeg";
            pictureBox4.Image.Save(t, ImageFormat.Jpeg);
            log.Debug("Save the CaptureDone Picture");
            button_Start.Enabled = true;
            TakePicture = false;
            TakePictureError = 0;
            //setStyle();
            log.Debug("Stop the CaptureDone function");
        }
        #endregion

        #region -- 字幕 --
        private void MySrtCamd()
        {
            int count = 1;
            string starttime = "0:0:0";
            TimeSpan time_start = TimeSpan.Parse(DateTime.Now.ToString("HH:mm:ss"));

            while (GlobalData.VideoRecording)
            {
                System.Threading.Thread.Sleep(1000);
                TimeSpan time_end = TimeSpan.Parse(DateTime.Now.ToString("HH:mm:ss")); //計時結束 取得目前時間
                //後面的時間減前面的時間後 轉型成TimeSpan即可印出時間差
                string endtime = (time_end - time_start).Hours.ToString() + ":" + (time_end - time_start).Minutes.ToString() + ":" + (time_end - time_start).Seconds.ToString();
                StreamWriter srtWriter = new StreamWriter(srtstring, true);
                srtWriter.WriteLine(count);

                srtWriter.WriteLine(starttime + ",001" + " --> " + endtime + ",000");
                srtWriter.WriteLine(label_Command.Text + "     " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                srtWriter.WriteLine(label_Remark.Text);
                srtWriter.WriteLine("");
                srtWriter.WriteLine("");
                srtWriter.Close();
                count++;
                starttime = endtime;
            }
        }
        #endregion

        #region -- 錄影 --
        private void Mysvideo() => Invoke(new EventHandler(delegate { Savevideo("avi"); }));//開始錄影//

        private void Mywmvideo() => Invoke(new EventHandler(delegate { Savevideo("wmv"); }));//開始錄影//

        private void Mysstop() => Invoke(new EventHandler(delegate//停止錄影//
        {
            capture.Stop();
            capture.Dispose();
            Camstart();
        }));

        private void Savevideo(string format)//儲存影片//
        {
            string fName = ini12.INIRead(MainSettingPath, "Record", "VideoPath", "");
            string t;
            if (format == "wmv")
                t = fName + "\\" + "_rec" + DateTime.Now.ToString("yyyyMMddHHmmss") + "__" + label_LoopNumber_Value.Text + ".wmv";
            else
            {
                t = fName + "\\" + "_rec" + DateTime.Now.ToString("yyyyMMddHHmmss") + "__" + label_LoopNumber_Value.Text + ".avi";
                srtstring = fName + "\\" + "_rec" + DateTime.Now.ToString("yyyyMMddHHmmss") + "__" + label_LoopNumber_Value.Text + ".srt";
            }

            if (!capture.Cued)
                capture.Filename = t;

            if (format == "wmv")
                capture.RecFileMode = DirectX.Capture.Capture.RecFileModeType.Wmv; //宣告我要wmv檔格式
            else
                capture.RecFileMode = DirectX.Capture.Capture.RecFileModeType.Avi; //宣告我要avi檔格式

            capture.Cue(); // 創一個檔
            capture.Start(); // 開始錄影

            /*
            double chd; //檢查HD 空間 小於100M就停止錄影s
            chd = ImageOpacity.ChDisk(ImageOpacity.Dkroot(fName));
            if (chd < 0.1)
            {
                Vread = false;
                MessageBox.Show("Check the HD Capacity!", "HD Capacity Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/
        }
        #endregion

        private void OnOffCamera()//啟動攝影機//
        {
            if (_captureInProgress == true)
            {
                Camstart();
            }

            if (_captureInProgress == false && capture != null)
            {
                capture.Stop();
                capture.Dispose();
            }
        }

        private void Camstart()
        {
            try
            {
                Filters filters = new Filters();
                Filter f;

                List<string> video = new List<string> { };
                for (int c = 0; c < filters.VideoInputDevices.Count; c++)
                {
                    f = filters.VideoInputDevices[c];
                    video.Add(f.Name);
                }

                List<string> audio = new List<string> { };
                for (int j = 0; j < filters.AudioInputDevices.Count; j++)
                {
                    f = filters.AudioInputDevices[j];
                    audio.Add(f.Name);
                }

                int scam, saud, VideoNum, AudioNum = 0;
                if (ini12.INIRead(MainSettingPath, "Camera", "VideoIndex", "") == "")
                    scam = 0;
                else
                    scam = int.Parse(ini12.INIRead(MainSettingPath, "Camera", "VideoIndex", ""));

                if (ini12.INIRead(MainSettingPath, "Camera", "AudioIndex", "") == "")
                    saud = 0;
                else
                    saud = int.Parse(ini12.INIRead(MainSettingPath, "Camera", "AudioIndex", ""));

                if (ini12.INIRead(MainSettingPath, "Camera", "VideoNumber", "") == "")
                    VideoNum = 0;
                else
                    VideoNum = int.Parse(ini12.INIRead(MainSettingPath, "Camera", "VideoNumber", ""));

                if (ini12.INIRead(MainSettingPath, "Camera", "AudioNumber", "") == "")
                    AudioNum = 0;
                else
                    AudioNum = int.Parse(ini12.INIRead(MainSettingPath, "Camera", "AudioNumber", ""));

                if (filters.VideoInputDevices.Count < VideoNum ||
                    filters.AudioInputDevices.Count < AudioNum)
                {
                    MessageBox.Show("Please reset video or audio device and select OK.", "Camera Status Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    button_Setting.PerformClick();
                }
                else
                {
                    capture = new Capture(filters.VideoInputDevices[scam], filters.AudioInputDevices[saud]);
                    try
                    {
                        capture.FrameSize = new Size(2304, 1296);
                        ini12.INIWrite(MainSettingPath, "Camera", "Resolution", "2304*1296");
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message.ToString(), "Webcam does not support 2304*1296!\n\r");
                        try
                        {
                            capture.FrameSize = new Size(1920, 1080);
                            ini12.INIWrite(MainSettingPath, "Camera", "Resolution", "1920*1080");
                        }
                        catch (Exception ex1)
                        {
                            Console.Write(ex1.Message.ToString(), "Webcam does not support 1920*1080!\n\r");
                            try
                            {
                                capture.FrameSize = new Size(1280, 720);
                                ini12.INIWrite(MainSettingPath, "Camera", "Resolution", "1280*720");
                            }
                            catch (Exception ex2)
                            {
                                Console.Write(ex2.Message.ToString(), "Webcam does not support 1280*720!\n\r");
                                try
                                {
                                    capture.FrameSize = new Size(640, 480);
                                    ini12.INIWrite(MainSettingPath, "Camera", "Resolution", "640*480");
                                }
                                catch (Exception ex3)
                                {
                                    Console.Write(ex3.Message.ToString(), "Webcam does not support 640*480!\n\r");
                                    try
                                    {
                                        capture.FrameSize = new Size(320, 240);
                                        ini12.INIWrite(MainSettingPath, "Camera", "Resolution", "320*240");
                                    }
                                    catch (Exception ex4)
                                    {
                                        Console.Write(ex4.Message.ToString(), "Webcam does not support 320*240!\n\r");
                                    }
                                }
                            }
                        }
                    }
                    capture.CaptureComplete += new EventHandler(OnCaptureComplete);
                }

                if (capture.PreviewWindow == null)
                {
                    try
                    {
                        capture.PreviewWindow = panelVideo;
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message.ToString(), "Please set the supported resolution!\n\r");
                    }
                }
                else
                {
                    capture.PreviewWindow = null;
                }
            }
            catch (NotSupportedException)
            {
                MessageBox.Show("Camera is disconnected unexpectedly!\r\nPlease go to Settings to reload the device list.", "Connection Error");
                button_Start.PerformClick();
            };
        }

        #region -- 讀取RC DB並填入combobox --
        private void LoadRCDB()
        {
            RedRatData.RedRatLoadSignalDB(ini12.INIRead(MainSettingPath, "RedRat", "DBFile", ""));
            RedRatData.RedRatSelectDevice(ini12.INIRead(MainSettingPath, "RedRat", "Brands", ""));

            DataGridViewComboBoxColumn RCDB = (DataGridViewComboBoxColumn)DataGridView_Schedule.Columns[0];

            string devicename = ini12.INIRead(MainSettingPath, "RedRat", "Brands", "");
            if (RedRatData.RedRatSelectDevice(devicename))
            {
                RCDB.Items.AddRange(RedRatData.RedRatGetRCNameList().ToArray());
                GlobalData.RcList = RedRatData.RedRatGetRCNameList();
                GlobalData.Rc_Number = RedRatData.RedRatGetRCNameList().Count;
            }
            else
            {
                Console.WriteLine("Select Device Error: " + devicename);
            }
            RCDB.Items.Add("_Execute");
            //RCDB.Items.Add("_Condition_AND");
            RCDB.Items.Add("_Condition_OR");
            RCDB.Items.Add("------------------------");
            RCDB.Items.Add("_HEX");
            RCDB.Items.Add("_FTDI");
            RCDB.Items.Add("_ascii");
            RCDB.Items.Add("------------------------");
            RCDB.Items.Add("_FuncKey");
            RCDB.Items.Add("_K_ABS");
            RCDB.Items.Add("_K_OBD");
            RCDB.Items.Add("_K_SEND");
            RCDB.Items.Add("_K_CLEAR");
            RCDB.Items.Add("_WaterTemp");
            RCDB.Items.Add("_FuelDisplay");
            RCDB.Items.Add("_Temperature");
            RCDB.Items.Add("------------------------");
            RCDB.Items.Add("_TX_I2C_Read");
            RCDB.Items.Add("_TX_I2C_Write");
            RCDB.Items.Add("_Canbus_Send");
            RCDB.Items.Add("_Canbus_Queue");
            RCDB.Items.Add("------------------------");
            RCDB.Items.Add("_shot");
            RCDB.Items.Add("_rec_start");
            RCDB.Items.Add("_rec_stop");
            RCDB.Items.Add("_cmd");
            RCDB.Items.Add("_DOS");
            RCDB.Items.Add("------------------------");
            RCDB.Items.Add("_IO_Output");
            RCDB.Items.Add("_IO_Input");
            RCDB.Items.Add("_Pin");
            RCDB.Items.Add("_Arduino_Output");
            RCDB.Items.Add("_Arduino_Input");
            RCDB.Items.Add("_Arduino_Pin");
            RCDB.Items.Add("_Arduino_Command");
            if (ini12.INIRead(MainSettingPath, "Device", "Software", "") == "All")
            {
                RCDB.Items.Add("_audio_debounce");
                RCDB.Items.Add("_keyword");
                RCDB.Items.Add("------------------------");
                RCDB.Items.Add("_quantum");
                RCDB.Items.Add("_astro");
                RCDB.Items.Add("_dektec");
                RCDB.Items.Add("_OPM");
            }
            RCDB.Items.Add("------------------------");
            //RCDB.Items.Add("------------------------");
            //RCDB.Items.Add("_SXP");
            //RCDB.Items.Add("_log1");
            //RCDB.Items.Add("_log2");
            //RCDB.Items.Add("_log3");
            //RCDB.Items.Add("_pwm1");
            //RCDB.Items.Add("_pwm2");
            //RCDB.Items.Add("_pwm3");
            //RCDB.Items.Add("------------------------");
            //RCDB.Items.Add("_EXT_Output_P0");
            //RCDB.Items.Add("_EXT_Output_P1");
            //RCDB.Items.Add("_EXT_Output_P2");
            //RCDB.Items.Add("_EXT_Input_P0");
            //RCDB.Items.Add("_EXT_Input_P1");
            //RCDB.Items.Add("_EXT_Input_P2");
            //RCDB.Items.Add("------------------------");
            //RCDB.Items.Add("_MonkeyTest");
        }
        #endregion

        #region -- 讀取RC DB並填入Virtual RC Panel --
        Button[] Buttons;
        private void LoadVirtualRC()
        {
            //根據dpi調整按鍵寬度//
            Graphics graphics = CreateGraphics();
            float dpiX = graphics.DpiX;
            float dpiY = graphics.DpiY;
            int width, height;
            if (dpiX == 96 && dpiY == 96)
            {
                width = 75;
                height = 25;
            }
            else
            {
                width = 90;
                height = 25;
            }

            Buttons = new Button[GlobalData.Rc_Number];

            for (int i = 0; i < Buttons.Length; i++)
            {
                Buttons[i] = new Button
                {
                    Size = new Size(width, height),
                    Text = GlobalData.RcList[i],
                    AutoSize = false,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink
                };

                if (i <= 11)
                {
                    Buttons[i].Location = new Point(0 + (i * width), 5);
                }
                else if (i > 11 && i <= 23)
                {
                    Buttons[i].Location = new Point(0 + ((i - 12) * width), 45);
                }
                else if (i > 23 && i <= 35)
                {
                    Buttons[i].Location = new Point(0 + ((i - 24) * width), 85);
                }
                else if (i > 35 && i <= 47)
                {
                    Buttons[i].Location = new Point(0 + ((i - 36) * width), 125);
                }
                else if (i > 47 && i <= 59)
                {
                    Buttons[i].Location = new Point(0 + ((i - 48) * width), 165);
                }
                else if (i > 59 && i <= 71)
                {
                    Buttons[i].Location = new Point(0 + ((i - 60) * width), 205);
                }
                else if (i > 71 && i <= 83)
                {
                    Buttons[i].Location = new Point(0 + ((i - 72) * width), 245);
                }
                else if (i > 83 && i <= 95)
                {
                    Buttons[i].Location = new Point(0 + ((i - 84) * width), 285);
                }

                int index = i;
                Buttons[i].Click += (sender1, ex) => Sand_Key(index + 1);
                panel_VirtualRC.Controls.Add(Buttons[i]);
            }
        }
        #endregion

        private void Sand_Key(int i)
        {
            if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
            {
                if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
                {
                    Autocommand_RedRat("Form1", Buttons[i - 1].Text);
                }
                else if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "2")
                {
                    Autocommand_BlueRat("Form1", Buttons[i - 1].Text);
                }
            }
            else if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
            {
                Autocommand_RedRat("Form1", Buttons[i - 1].Text);
            }
        }

        void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ComboBox cmb = e.Control as ComboBox;
            if (cmb != null)
            {
                cmb.DropDown -= new EventHandler(cmb_DropDown);
                cmb.DropDown += new EventHandler(cmb_DropDown);
            }
        }

        //自動調整ComboBox寬度//
        void cmb_DropDown(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            int width = cmb.DropDownWidth;
            Graphics g = cmb.CreateGraphics();
            Font font = cmb.Font;
            int vertScrollBarWidth = 0;
            if (cmb.Items.Count > cmb.MaxDropDownItems)
            {
                vertScrollBarWidth = SystemInformation.VerticalScrollBarWidth;
            }

            int maxWidth;
            foreach (string s in cmb.Items)
            {
                maxWidth = (int)g.MeasureString(s, font).Width + vertScrollBarWidth;
                if (width < maxWidth)
                {
                    width = maxWidth;
                }
            }

            DataGridViewComboBoxColumn c =
                DataGridView_Schedule.Columns[0] as DataGridViewComboBoxColumn;
            if (c != null)
            {
                c.DropDownWidth = width;
            }
        }

        private bool StartButtonFlag = false;
        private DateTime startTime;
        long delayTime = 0;
        long repeatTime = 0;
        long timeCountUpdated = 0;
        private void StartBtn_Click(object sender, EventArgs e)
        {
            byte[] val = new byte[2];
            val[0] = 0;
            bool AutoBox_Status;

            GlobalData.IO_PA10_0_COUNT = 0;
            GlobalData.IO_PA10_1_COUNT = 0;
            GlobalData.IO_PA11_0_COUNT = 0;
            GlobalData.IO_PA11_1_COUNT = 0;
            GlobalData.IO_PA14_0_COUNT = 0;
            GlobalData.IO_PA14_1_COUNT = 0;
            GlobalData.IO_PA15_0_COUNT = 0;
            GlobalData.IO_PA15_1_COUNT = 0;
            GlobalData.IO_PB1_0_COUNT = 0;
            GlobalData.IO_PB1_1_COUNT = 0;
            GlobalData.IO_PB7_0_COUNT = 0;
            GlobalData.IO_PB7_1_COUNT = 0;

            GlobalData.IO_Arduino2_0_COUNT = 0;
            GlobalData.IO_Arduino2_1_COUNT = 0;
            GlobalData.IO_Arduino3_0_COUNT = 0;
            GlobalData.IO_Arduino3_1_COUNT = 0;
            GlobalData.IO_Arduino4_0_COUNT = 0;
            GlobalData.IO_Arduino4_1_COUNT = 0;
            GlobalData.IO_Arduino5_0_COUNT = 0;
            GlobalData.IO_Arduino5_1_COUNT = 0;
            GlobalData.IO_Arduino6_0_COUNT = 0;
            GlobalData.IO_Arduino6_1_COUNT = 0;
            GlobalData.IO_Arduino7_0_COUNT = 0;
            GlobalData.IO_Arduino7_1_COUNT = 0;
            GlobalData.IO_Arduino8_0_COUNT = 0;
            GlobalData.IO_Arduino8_1_COUNT = 0;
            GlobalData.IO_Arduino9_0_COUNT = 0;
            GlobalData.IO_Arduino9_1_COUNT = 0;

            delayTime = 0;
            repeatTime = 0;

            AutoBox_Status = ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1" ? true : false;

            if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
            {
                if (!_captureInProgress)
                {
                    _captureInProgress = true;
                    OnOffCamera();
                }
            }

            Thread MainThread = new Thread(new ThreadStart(MyRunCamd));
            Thread LogThread1 = new Thread(new ThreadStart(MyLog1Camd));
            Thread LogThread2 = new Thread(new ThreadStart(MyLog2Camd));
            Thread LogThread3 = new Thread(new ThreadStart(MyLog3Camd));
            Thread LogThread4 = new Thread(new ThreadStart(MyLog4Camd));
            Thread LogThread5 = new Thread(new ThreadStart(MyLog5Camd));
            Thread LogAThread = new Thread(new ThreadStart(logA_analysis));
            Thread LogBThread = new Thread(new ThreadStart(logB_analysis));
            Thread LogCThread = new Thread(new ThreadStart(logC_analysis));
            Thread LogDThread = new Thread(new ThreadStart(logD_analysis));
            Thread LogEThread = new Thread(new ThreadStart(logE_analysis));

            Thread RK2797A_Thread = new Thread(new ThreadStart(logA_RK2797));
            Thread RK2797B_Thread = new Thread(new ThreadStart(logB_RK2797));
            Thread RK2797C_Thread = new Thread(new ThreadStart(logC_RK2797));
            Thread RK2797D_Thread = new Thread(new ThreadStart(logD_RK2797));
            Thread RK2797E_Thread = new Thread(new ThreadStart(logE_RK2797));

            startTime = DateTime.Now;

            if (AutoBox_Status)//如果電腦有接上AutoKit//
            {
                button_Schedule1.PerformClick();

                //Thread Log1Data = new Thread(new ThreadStart(Log1_Receiving_Task));
                //Thread Log2Data = new Thread(new ThreadStart(Log2_Receiving_Task));

                if (StartButtonPressed == true)//按下STOP//
                {
                    GlobalData.Break_Out_MyRunCamd = 1;//跳出倒數迴圈//
                    MainThread.Abort();//停止執行緒//
                    timer_countdown.Stop();//停止倒數//
                    CloseDtplay();//關閉DtPlay//
                    duringTimer.Enabled = false;
                    if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1")
                    {
                        can_send = 0;
                        set_timer_rate = false;
                        can_rate.Clear();
                        can_data.Clear();
                    }

                    if (GlobalData.portConfigGroup_A.checkedValue)
                    {
                        LogAThread.Abort();
                        RK2797A_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread1.Abort();
                            //Log1Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_B.checkedValue)
                    {
                        LogBThread.Abort();
                        RK2797B_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread2.Abort();
                            //Log2Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_C.checkedValue)
                    {
                        LogCThread.Abort();
                        RK2797C_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread3.Abort();
                            //Log3Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_D.checkedValue)
                    {
                        LogDThread.Abort();
                        RK2797D_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread4.Abort();
                            //Log4Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_E.checkedValue)
                    {
                        LogEThread.Abort();
                        RK2797E_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread5.Abort();
                            //Log5Data.Abort();
                        }
                    }

                    StartButtonPressed = false;
                    button_Start.Enabled = false;
                    button_Setting.Enabled = false;
                    button_SaveSchedule.Enabled = false;
                    button_Pause.Enabled = true;
                    //setStyle();
                    label_Command.Text = "Please wait...";
                }
                else     //按下START//
                {
                    /*
                    for (int i = 1; i < 6; i++)
                    {
                        if (Directory.Exists(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + i + "_Original") == true)
                        {
                            DirectoryInfo DIFO = new DirectoryInfo(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + i + "_Original");
                            DIFO.Delete(true);
                        }

                        if (Directory.Exists(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + i + "_NG") == true)
                        {
                            DirectoryInfo DIFO = new DirectoryInfo(ini12.INIRead(sPath, "Record", "VideoPath", "") + "\\" + "Schedule" + i + "_NG");
                            DIFO.Delete(true);
                        }                
                    }
                    */
                    GlobalData.Break_Out_MyRunCamd = 0;
                    if (CA210.Status() == true)
                        createCA210folder();
                    ini12.INIWrite(MainSettingPath, "LogSearch", "StartTime", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
                    MainThread.Start();       // 啟動執行緒
                    timer_countdown.Start();     //開始倒數
                    button_Start.Text = "STOP";

                    StartButtonPressed = true;
                    StartButtonFlag = true;
                    button_Setting.Enabled = false;
                    button_Pause.Enabled = true;
                    button_SaveSchedule.Enabled = false;
                    //setStyle();

                    //if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                    if (GlobalData.portConfigGroup_A.checkedValue)
                    {
                        //serialPortConfig_A = Form_Setting.checkBox_SerialPort1.Text;                                          //PortA
                        //serialPortName_A = Form_Setting.comboBox_SerialPort1_PortName_Value.Text;         //e.g. COM10
                        //serialPortBR_A = Form_Setting.comboBox_SerialPort1_BaudRate_Value.Text;               //e.g. 9600
                        //serialPortA.OpenSerialPort(GlobalData.portConfigGroup_A.portName, GlobalData.portConfigGroup_A.portBR);
                        GlobalData.m_SerialPort_A.OpenSerialPort(GlobalData.portConfigGroup_A.portName, GlobalData.portConfigGroup_A.portBR);
                        //OpenSerialPort("A");      //previous implementation with no DrvRS232 support

                        //logDumpping.LogDataReceiving(GlobalData.m_SerialPort_A, GlobalData.portConfigGroup_A.portConfig, ref logA_text);// InitlogConfig(port_Label_A, serialPortName, ref logA_text);
                        LogAThread.Start();
                        RK2797A_Thread.Start();
                        textBox_serial.Clear();
                        //LogAThread.Start();
                        //textBox1.Text = string.Empty;//清空serialport1//
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport1", "") == "1")
                        {
                            LogThread1.IsBackground = true;
                            LogThread1.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_B.checkedValue)
                    {
                        GlobalData.m_SerialPort_B.OpenSerialPort(GlobalData.portConfigGroup_B.portName, GlobalData.portConfigGroup_B.portBR);
                        //logDumpping.LogDataReceiving(GlobalData.m_SerialPort_B, GlobalData.portConfigGroup_B.portConfig, ref logB_text);// InitlogConfig(port_Label_A, serialPortName, ref logA_text);
                        LogBThread.Start();
                        RK2797B_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport2", "") == "1")
                        {
                            LogThread2.IsBackground = true;
                            LogThread2.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_C.checkedValue)
                    {
                        GlobalData.m_SerialPort_C.OpenSerialPort(GlobalData.portConfigGroup_C.portName, GlobalData.portConfigGroup_C.portBR);
                        //logDumpping.LogDataReceiving(GlobalData.m_SerialPort_C, GlobalData.portConfigGroup_C.portConfig, ref logC_text);// InitlogConfig(port_Label_A, serialPortName, ref logA_text);
                        LogCThread.Start();
                        RK2797C_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport3", "") == "1")
                        {
                            LogThread3.IsBackground = true;
                            LogThread3.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_D.checkedValue)
                    {
                        GlobalData.m_SerialPort_D.OpenSerialPort(GlobalData.portConfigGroup_D.portName, GlobalData.portConfigGroup_D.portBR);
                        //logDumpping.LogDataReceiving(GlobalData.m_SerialPort_D, GlobalData.portConfigGroup_D.portConfig, ref logD_text);// InitlogConfig(port_Label_A, serialPortName, ref logA_text);
                        LogDThread.Start();
                        RK2797D_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport4", "") == "1")
                        {
                            LogThread4.IsBackground = true;
                            LogThread4.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_E.checkedValue)
                    {
                        GlobalData.m_SerialPort_E.OpenSerialPort(GlobalData.portConfigGroup_E.portName, GlobalData.portConfigGroup_E.portBR);
                        //logDumpping.LogDataReceiving(GlobalData.m_SerialPort_E, GlobalData.portConfigGroup_E.portConfig, ref logE_text);// InitlogConfig(port_Label_A, serialPortName, ref logA_text);
                        RK2797E_Thread.Start();
                        LogEThread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport5", "") == "1")
                        {
                            LogThread5.IsBackground = true;
                            LogThread5.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_Kline.checkedValue)
                    {
                        //serialPortK.OpenSerialPort("Port kline");
                        MySerialPort.OpenPort("kline");
                        textBox_serial.Text = "";   //清空kline//

                        if (MySerialPort.IsPortOpened())
                        {
                            //BlueRat_UART_Exception_status = false;
                            timer_kline.Enabled = true;
                        }
                        else
                        {
                            timer_kline.Enabled = false;
                        }
                    }

                    label_Command.Text = "";
                }
            }
            else     //如果沒接AutoKit//
            {
                if (StartButtonPressed == true)     //按下STOP//
                {
                    GlobalData.Break_Out_MyRunCamd = 1;    //跳出倒數迴圈
                    MainThread.Abort(); //停止執行緒
                    timer_countdown.Stop();  //停止倒數
                    CloseDtplay();
                    duringTimer.Enabled = false;
                    if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1")
                    {
                        can_send = 0;
                        set_timer_rate = false;
                        can_rate.Clear();
                        can_data.Clear();
                    }

                    if (GlobalData.portConfigGroup_A.checkedValue)
                    {
                        LogAThread.Abort();
                        RK2797A_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread1.Abort();
                            //Log1Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_B.checkedValue)
                    {
                        LogBThread.Abort();
                        RK2797B_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread2.Abort();
                            //Log2Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_C.checkedValue)
                    {
                        LogCThread.Abort();
                        RK2797C_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread3.Abort();
                            //Log3Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_D.checkedValue)
                    {
                        LogDThread.Abort();
                        RK2797D_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread4.Abort();
                            //Log4Data.Abort();
                        }
                    }

                    if (GlobalData.portConfigGroup_E.checkedValue)
                    {
                        LogEThread.Abort();
                        RK2797E_Thread.Abort();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0")
                        {
                            LogThread5.Abort();
                            //Log4Data.Abort();
                        }
                    }

                    StartButtonPressed = false;
                    button_Start.Enabled = false;
                    button_Setting.Enabled = false;
                    button_Pause.Enabled = true;
                    button_SaveSchedule.Enabled = false;
                    //setStyle();

                    label_Command.Text = "Please wait...";
                }
                else     //按下START//
                {
                    GlobalData.Break_Out_MyRunCamd = 0;
                    if (CA210.Status() == true)
                        createCA210folder();
                    MainThread.Start();// 啟動執行緒
                    timer_countdown.Start();     //開始倒數
                    StartButtonPressed = true;
                    StartButtonFlag = true;
                    button_Setting.Enabled = false;
                    button_Pause.Enabled = true;
                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                    button_Start.Text = "STOP";
                    //setStyle();

                    //if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
                    if (GlobalData.portConfigGroup_A.checkedValue)
                    {
                        GlobalData.m_SerialPort_A.OpenSerialPort(GlobalData.portConfigGroup_A.portName, GlobalData.portConfigGroup_A.portBR);
                        //serialPortA.OpenSerialPort(GlobalData.portConfigGroup_A.portName, GlobalData.portConfigGroup_A.portBR);
                        textBox_serial.Clear();
                        LogAThread.Start();
                        RK2797A_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport1", "") == "1")
                        {
                            LogThread1.IsBackground = true;
                            LogThread1.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_B.checkedValue)
                    {
                        GlobalData.m_SerialPort_B.OpenSerialPort(GlobalData.portConfigGroup_B.portName, GlobalData.portConfigGroup_B.portBR);
                        LogBThread.Start();
                        RK2797B_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport2", "") == "1")
                        {
                            LogThread2.IsBackground = true;
                            LogThread2.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_C.checkedValue)
                    {
                        GlobalData.m_SerialPort_C.OpenSerialPort(GlobalData.portConfigGroup_C.portName, GlobalData.portConfigGroup_C.portBR);
                        LogCThread.Start();
                        RK2797C_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport3", "") == "1")
                        {
                            LogThread3.IsBackground = true;
                            LogThread3.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_D.checkedValue)
                    {
                        GlobalData.m_SerialPort_D.OpenSerialPort(GlobalData.portConfigGroup_D.portName, GlobalData.portConfigGroup_D.portBR);
                        LogDThread.Start();
                        RK2797D_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport4", "") == "1")
                        {
                            LogThread4.IsBackground = true;
                            LogThread4.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_E.checkedValue)
                    {
                        GlobalData.m_SerialPort_E.OpenSerialPort(GlobalData.portConfigGroup_E.portName, GlobalData.portConfigGroup_E.portBR);
                        LogEThread.Start();
                        RK2797E_Thread.Start();
                        if (ini12.INIRead(MainSettingPath, "LogSearch", "TextNum", "") != "0" && ini12.INIRead(MainSettingPath, "LogSearch", "Comport5", "") == "1")
                        {
                            LogThread5.IsBackground = true;
                            LogThread5.Start();
                        }
                    }

                    if (GlobalData.portConfigGroup_Kline.checkedValue)
                    {
                        //serialPortK.OpenSerialPort("Port kline");
                        MySerialPort.OpenPort("kline");
                        textBox_serial.Text = "";    //清空kline//
                    }

                    label_Command.Text = "";
                }
            }
        }

        private void SettingBtn_Click(object sender, EventArgs e)
        {
            FormTabControl FormTabControl = new FormTabControl();
            GlobalData.RCDB = ini12.INIRead(MainSettingPath, "RedRat", "Brands", "");
            if (GlobalData.m_Arduino_Port.IsOpen() == true)
            {
                GlobalData.m_Arduino_Port.ClosePort();
                pictureBox_ext_board.Image = Properties.Resources.OFF;
            }

            if (ini12.INIRead(MainSettingPath, "Device", "FTDIExist", "") == "1" && portinfo.ftStatus == FtResult.Ok)
            {
                DisconnectFtdi();
                pictureBox_ftdi.Image = Properties.Resources.OFF;
            }

            if (CA210.Status() == true)
            {
                CA210.DisConnect();
                pictureBox_Minolta.Image = Properties.Resources.OFF;
            }

            //關閉SETTING以後會讀這段>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            if (FormTabControl.ShowDialog() == DialogResult.OK)
            {
                if (ini12.INIRead(MainSettingPath, "RedRat", "Brands", "") != GlobalData.RCDB)
                {
                    DataGridViewComboBoxColumn RCDB = (DataGridViewComboBoxColumn)DataGridView_Schedule.Columns[0];
                    RCDB.Items.Clear();
                    LoadRCDB();

                    //panel_VirtualRC.Controls.Clear();
                    //LoadVirtualRC();
                }

                if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1")
                {
                    if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "1")
                    {
                        ConnectAutoBox1();
                    }

                    if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "2" && ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "0" && ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "0")
                    {
                        ConnectAutoBox2();
                    }

                    pictureBox_BlueRat.Image = Properties.Resources.ON;
                    if (ini12.INIRead(MainSettingPath, "Device", "AutoboxExist", "") == "1" && ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "0" && ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "0")
                    {
                        GP0_GP1_AC_ON();
                        GP2_GP3_USB_PC();
                    }
                }
                else
                {
                    pictureBox_BlueRat.Image = Properties.Resources.OFF;
                    pictureBox_AcPower.Image = Properties.Resources.OFF;
                    button_AcUsb.Enabled = false;
                    PowerState = false;
                    MyBlueRat.Disconnect(); //Prevent from System.ObjectDisposedException
                }

                if (ini12.INIRead(MainSettingPath, "Device", "ArduinoExist", "") == "1")
                {
                    if (ini12.INIRead(MainSettingPath, "Device", "ArduinoPort", "") != "")
                    {
                        GlobalData.m_Arduino_Port.OpenSerialPort(GlobalData.Arduino_Comport, GlobalData.Arduino_Baudrate, true);
                        pictureBox_ext_board.Image = Properties.Resources.ON;
                    }
                    else
                    {
                        pictureBox_ext_board.Image = Properties.Resources.OFF;
                    }
                }

                if (ini12.INIRead(MainSettingPath, "Device", "FTDIExist", "") == "1")
                {
                    ConnectFtdi();
                    if (portinfo.ftStatus == FtResult.Ok)
                        pictureBox_ftdi.Image = Properties.Resources.ON;
                    else
                        pictureBox_ftdi.Image = Properties.Resources.OFF;
                }
                else
                {
                    pictureBox_ftdi.Image = Properties.Resources.OFF;
                }

                if (ini12.INIRead(MainSettingPath, "Device", "RedRatExist", "") == "1")
                {
                    OpenRedRat3();
                }
                else
                {
                    pictureBox_RedRat.Image = Properties.Resources.OFF;
                    ini12.INIWrite(MainSettingPath, "Device", "RedRatExist", "0");
                }

                if (ini12.INIRead(MainSettingPath, "Device", "CameraExist", "") == "1")
                {
                    try
                    {
                        pictureBox_Camera.Image = Properties.Resources.ON;
                        _captureInProgress = false;
                        OnOffCamera();
                        button_VirtualRC.Enabled = true;
                        comboBox_CameraDevice.Enabled = false;
                        button_Camera.Enabled = true;
                        string[] cameraDevice = ini12.INIRead(MainSettingPath, "Camera", "CameraDevice", "").Split(',');
                        comboBox_CameraDevice.Items.Clear();
                        foreach (string cd in cameraDevice)
                        {
                            comboBox_CameraDevice.Items.Add(cd);
                        }
                        comboBox_CameraDevice.SelectedIndex = Int32.Parse(ini12.INIRead(MainSettingPath, "Camera", "VideoIndex", ""));
                    }
                    catch (ArgumentOutOfRangeException)
                    {

                    }
                }
                else
                {
                    pictureBox_Camera.Image = Properties.Resources.OFF;
                    button_Camera.Enabled = false;
                }

                //if (ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "1" || ini12.INIRead(MainSettingPath, "Device", "CA410Exist", "") == "1")
                //{
                //    if (CA210.Status() == false)
                //    {
                //        CA210.Connect();
                //        if (CA210.Status() == false)
                //        {
                //            MessageBox.Show("Minolta is not connected!\r\nPlease restart the OPTT to reload the device.", "Connection Error");
                //            pictureBox_Minolta.Image = Properties.Resources.OFF;
                //        }
                //    }
                //}
                //else
                //{
                //    pictureBox_Minolta.Image = Properties.Resources.OFF;
                //}

                /* Hidden serial port.
                button_SerialPort1.Visible = ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1" ? true : false;
                button_SerialPort2.Visible = ini12.INIRead(MainSettingPath, "Port B", "Checked", "") == "1" ? true : false;
                button_SerialPort3.Visible = ini12.INIRead(MainSettingPath, "Port C", "Checked", "") == "1" ? true : false;
                button_CanbusPort.Visible = ini12.INIRead(MainSettingPath, "Canbus", "Log", "") == "1" ? true : false;
                button_kline.Visible = ini12.INIRead(MainSettingPath, "Kline", "Checked", "") == "1" ? true : false;
                */
                List<string> SchExist = new List<string> { };
                for (int i = 2; i < 6; i++)
                {
                    SchExist.Add(ini12.INIRead(MainSettingPath, "Schedule" + i, "Exist", ""));
                }

                comboBox_savelog.Items.Clear();
                InitComboboxSaveLog();

                button_Schedule2.Visible = SchExist[0] == "0" ? false : true;
                button_Schedule3.Visible = SchExist[1] == "0" ? false : true;
                button_Schedule4.Visible = SchExist[2] == "0" ? false : true;
                button_Schedule5.Visible = SchExist[3] == "0" ? false : true;
            }

            FormTabControl.Dispose();
            button_Schedule1.Enabled = true;
            button_Schedule1.PerformClick();

            //setStyle();
        }

        //系統時間
        private void timer_Tick(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            TimeLabel.Text = string.Format("{0:R}", dt);            //拍照打印時間
            TimeLabel2.Text = string.Format("{0:yyyy-MM-dd  HH:mm:ss}", dt);

            #region -- schedule timer --
            if (ini12.INIRead(MainSettingPath, "Schedule1", "OnTimeStart", "") == "1")
                labelSch1Timer.Text = "Schedule 1 will start at" + "\r\n" + ini12.INIRead(MainSettingPath, "Schedule1", "Timer", "");
            else if (ini12.INIRead(MainSettingPath, "Schedule1", "OnTimeStart", "") == "0")
                labelSch1Timer.Text = "";

            if (ini12.INIRead(MainSettingPath, "Schedule2", "OnTimeStart", "") == "1")
                labelSch2Timer.Text = "Schedule 2 will start at" + "\r\n" + ini12.INIRead(MainSettingPath, "Schedule2", "Timer", "");
            else if (ini12.INIRead(MainSettingPath, "Schedule2", "OnTimeStart", "") == "0")
                labelSch2Timer.Text = "";

            if (ini12.INIRead(MainSettingPath, "Schedule3", "OnTimeStart", "") == "1")
                labelSch3Timer.Text = "Schedule 3 will start at" + "\r\n" + ini12.INIRead(MainSettingPath, "Schedule3", "Timer", "");
            else if (ini12.INIRead(MainSettingPath, "Schedule3", "OnTimeStart", "") == "0")
                labelSch3Timer.Text = "";

            if (ini12.INIRead(MainSettingPath, "Schedule4", "OnTimeStart", "") == "1")
                labelSch4Timer.Text = "Schedule 4 will start at" + "\r\n" + ini12.INIRead(MainSettingPath, "Schedule4", "Timer", "");
            else if (ini12.INIRead(MainSettingPath, "Schedule4", "OnTimeStart", "") == "0")
                labelSch4Timer.Text = "";

            if (ini12.INIRead(MainSettingPath, "Schedule5", "OnTimeStart", "") == "1")
                labelSch5Timer.Text = "Schedule 5 will start at" + "\r\n" + ini12.INIRead(MainSettingPath, "Schedule5", "Timer", "");
            else if (ini12.INIRead(MainSettingPath, "Schedule5", "OnTimeStart", "") == "0")
                labelSch5Timer.Text = "";

            if (ini12.INIRead(MainSettingPath, "Schedule1", "OnTimeStart", "") == "1" &&
                ini12.INIRead(MainSettingPath, "Schedule1", "Timer", "") == TimeLabel2.Text)
                button_Start.PerformClick();
            if (ini12.INIRead(MainSettingPath, "Schedule2", "OnTimeStart", "") == "1" &&
                ini12.INIRead(MainSettingPath, "Schedule2", "Timer", "") == TimeLabel2.Text &&
                timeCount != 0)
                GlobalData.Break_Out_Schedule = 1;
            if (ini12.INIRead(MainSettingPath, "Schedule3", "OnTimeStart", "") == "1" &&
                ini12.INIRead(MainSettingPath, "Schedule3", "Timer", "") == TimeLabel2.Text &&
                timeCount != 0)
                GlobalData.Break_Out_Schedule = 1;
            if (ini12.INIRead(MainSettingPath, "Schedule4", "OnTimeStart", "") == "1" &&
                ini12.INIRead(MainSettingPath, "Schedule4", "Timer", "") == TimeLabel2.Text &&
                timeCount != 0)
                GlobalData.Break_Out_Schedule = 1;
            if (ini12.INIRead(MainSettingPath, "Schedule5", "OnTimeStart", "") == "1" &&
                ini12.INIRead(MainSettingPath, "Schedule5", "Timer", "") == TimeLabel2.Text &&
                timeCount != 0)
                GlobalData.Break_Out_Schedule = 1;
            #endregion
        }

        //關閉Excel
        private void CloseExcel()
        {
            Process[] processes = Process.GetProcessesByName("EXCEL");

            foreach (Process p in processes)
            {
                p.Kill();
            }
        }

        //關閉DtPlay
        private void CloseDtplay()
        {
            Process[] processes = Process.GetProcessesByName("DtPlay");

            foreach (Process p in processes)
            {
                p.Kill();
            }
        }

        //關閉AutoKit
        private void CloseAutobox()
        {
            FormIsClosing = true;
            if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "1")
            {
                DisconnectAutoBox1();
            }

            if (ini12.INIRead(MainSettingPath, "Device", "AutoboxVerson", "") == "2")
            {
                DisconnectAutoBox2();
            }

            if (CA210.Status() == true)
                CA210.DisConnect();

            if (GlobalData.m_Arduino_Port.IsOpen() == true)
            {
                GlobalData.m_Arduino_Port.ClosePort();
            }

            if (ini12.INIRead(MainSettingPath, "Device", "FTDIExist", "") == "1")
            {
                DisconnectFtdi();
            }

            Application.ExitThread();
            Application.Exit();
            Environment.Exit(Environment.ExitCode);
        }

        private void Com1Btn_Click(object sender, EventArgs e)
        {
            OpenSerialPort("A");
            Controls.Add(textBox_serial);
            textBox_serial.BringToFront();
            GlobalData.TEXTBOX_FOCUS = 1;
        }

        private void Button_TabScheduler_Click(object sender, EventArgs e) => DataGridView_Schedule.BringToFront();
        private void Button_TabCamera_Click(object sender, EventArgs e)
        {
            if (!_captureInProgress)
            {
                _captureInProgress = true;
                OnOffCamera();
            }
            panelVideo.BringToFront();
            comboBox_CameraDevice.Enabled = true;
            comboBox_CameraDevice.BringToFront();
        }

        private void MyExportCamd()
        {
            string ab_num = label_LoopNumber_Value.Text,                                                        //自動編號
                        ab_p_id = ini12.INIRead(MailPath, "Data Info", "ProjectNumber", ""),                    //Project number
                        ab_c_id = ini12.INIRead(MailPath, "Data Info", "TestCaseNumber", ""),                   //Test case number
                        ab_result = ini12.INIRead(MailPath, "Data Info", "Result", ""),                         //Woodpecker 測試結果
                        ab_version = ini12.INIRead(MailPath, "Mail Info", "Version", ""),                       //軟體版號
                        ab_ng = ini12.INIRead(MailPath, "Data Info", "NGfrequency", ""),                        //NG frequency
                        ab_create = ini12.INIRead(MailPath, "Data Info", "CreateTime", ""),                     //測試開始時間
                        ab_close = ini12.INIRead(MailPath, "Data Info", "CloseTime", ""),                       //測試結束時間
                        ab_time = ini12.INIRead(MailPath, "Total Test Time", "value", ""),                      //測試執行花費時間
                        ab_loop = GlobalData.Schedule_Loop.ToString(),                                              //執行loop次數
                        ab_loop_time = ini12.INIRead(MailPath, "Total Test Time", "value", ""),                 //1個loop需要次數
                        ab_loop_step = (DataGridView_Schedule.Rows.Count - 1).ToString(),                       //1個loop的step數
                        ab_root = ini12.INIRead(MailPath, "Data Info", "Reboot", ""),                           //測試重啟次數
                        ab_user = ini12.INIRead(MailPath, "Mail Info", "Tester", ""),                           //測試人員
                        ab_mail = ini12.INIRead(MailPath, "Mail Info", "To", "");                               //Mail address 列表

            List<string> DataList = new List<string> { };
            DataList.Add(ab_num);
            DataList.Add(ab_p_id);
            DataList.Add(ab_c_id);
            DataList.Add(ab_result);
            DataList.Add(ab_version);
            DataList.Add(ab_ng);
            DataList.Add(ab_create);
            DataList.Add(ab_close);
            DataList.Add(ab_time);
            DataList.Add(ab_loop);
            DataList.Add(ab_loop_time);
            DataList.Add(ab_loop_step);
            DataList.Add(ab_root);
            DataList.Add(ab_user);
            DataList.Add(ab_mail);

            //Form_DGV_Autobox.DataInsert(DataList);
            //Form_DGV_Autobox.ToCsV(Form_DGV_Autobox.DGV_Autobox, "C:\\Woodpecker v2\\Report.xls");
        }

        #region -- 另存Schedule --
        private void WriteBtn_Click(object sender, EventArgs e)
        {
            string delimiter = ",";

            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "CSV files (*.csv)|*.csv";
            sfd.FileName = ini12.INIRead(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "Path", "");
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sfd.FileName, false))
                {
                    //output header data
                    string strHeader = "";
                    for (int i = 0; i < DataGridView_Schedule.Columns.Count; i++)
                    {
                        strHeader += DataGridView_Schedule.Columns[i].HeaderText + delimiter;
                    }
                    sw.WriteLine(strHeader.Replace("\r\n", "~"));

                    //output rows data
                    for (int j = 0; j < DataGridView_Schedule.Rows.Count - 1; j++)
                    {
                        string strRowValue = "";

                        for (int k = 0; k < DataGridView_Schedule.Columns.Count; k++)
                        {
                            string scheduleOutput = DataGridView_Schedule.Rows[j].Cells[k].Value + "";
                            if (scheduleOutput.Contains(","))
                            {
                                scheduleOutput = String.Format("\"{0}\"", scheduleOutput);
                            }
                            strRowValue += scheduleOutput + delimiter;
                        }
                        sw.WriteLine(strRowValue);
                    }
                    sw.Close();
                }

                if (sfd.FileName != ini12.INIRead(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "Path", ""))
                {
                    ini12.INIWrite(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "Path", sfd.FileName);
                }
            }
            ReadSch();
        }
        #endregion

        private void button_insert_a_row_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridView_Schedule.Rows.Insert(DataGridView_Schedule.CurrentCell.RowIndex, new DataGridViewRow());
            }
            catch (Exception)
            {
                MessageBox.Show("Please load or write a new schedule", "Schedule Error");
            }
        }

        #region -- Form1的Schedule 1~5按鈕功能 --
        private void SchBtn1_Click(object sender, EventArgs e)          ////////////Schedule1
        {
            portos_online = new SafeDataGridView();
            GlobalData.Schedule_Number = 1;
            string loop = ini12.INIRead(MainSettingPath, "Schedule1", "Loop", "");
            if (loop != "")
                GlobalData.Schedule_Loop = int.Parse(loop);
            labellabel_LoopTimes_Value.Text = GlobalData.Schedule_Loop.ToString();
            button_Schedule1.Enabled = false;
            button_Schedule2.Enabled = true;
            button_Schedule3.Enabled = true;
            button_Schedule4.Enabled = true;
            button_Schedule5.Enabled = true;
            ReadSch();
            ini12.INIWrite(MailPath, "Data Info", "TestCaseNumber", "0");
            ini12.INIWrite(MailPath, "Data Info", "Result", "N/A");
            ini12.INIWrite(MailPath, "Data Info", "NGfrequency", "0");
        }
        private void SchBtn2_Click(object sender, EventArgs e)          ////////////Schedule2
        {
            portos_online = new SafeDataGridView();
            GlobalData.Schedule_Number = 2;
            string loop = "";
            loop = ini12.INIRead(MainSettingPath, "Schedule2", "Loop", "");
            if (loop != "")
                GlobalData.Schedule_Loop = int.Parse(loop);
            labellabel_LoopTimes_Value.Text = GlobalData.Schedule_Loop.ToString();
            button_Schedule1.Enabled = true;
            button_Schedule2.Enabled = false;
            button_Schedule3.Enabled = true;
            button_Schedule4.Enabled = true;
            button_Schedule5.Enabled = true;
            LoadRCDB();
            ReadSch();
        }
        private void SchBtn3_Click(object sender, EventArgs e)          ////////////Schedule3
        {
            portos_online = new SafeDataGridView();
            GlobalData.Schedule_Number = 3;
            string loop = ini12.INIRead(MainSettingPath, "Schedule3", "Loop", "");
            if (loop != "")
                GlobalData.Schedule_Loop = int.Parse(loop);
            labellabel_LoopTimes_Value.Text = GlobalData.Schedule_Loop.ToString();
            button_Schedule1.Enabled = true;
            button_Schedule2.Enabled = true;
            button_Schedule3.Enabled = false;
            button_Schedule4.Enabled = true;
            button_Schedule5.Enabled = true;
            ReadSch();
        }
        private void SchBtn4_Click(object sender, EventArgs e)          ////////////Schedule4
        {
            portos_online = new SafeDataGridView();
            GlobalData.Schedule_Number = 4;
            string loop = ini12.INIRead(MainSettingPath, "Schedule4", "Loop", "");
            if (loop != "")
                GlobalData.Schedule_Loop = int.Parse(loop);
            labellabel_LoopTimes_Value.Text = GlobalData.Schedule_Loop.ToString();
            button_Schedule1.Enabled = true;
            button_Schedule2.Enabled = true;
            button_Schedule3.Enabled = true;
            button_Schedule4.Enabled = false;
            button_Schedule5.Enabled = true;
            ReadSch();
        }
        private void SchBtn5_Click(object sender, EventArgs e)          ////////////Schedule5
        {
            portos_online = new SafeDataGridView();
            GlobalData.Schedule_Number = 5;
            string loop = ini12.INIRead(MainSettingPath, "Schedule5", "Loop", "");
            if (loop != "")
                GlobalData.Schedule_Loop = int.Parse(loop);
            labellabel_LoopTimes_Value.Text = GlobalData.Schedule_Loop.ToString();
            button_Schedule1.Enabled = true;
            button_Schedule2.Enabled = true;
            button_Schedule3.Enabled = true;
            button_Schedule4.Enabled = true;
            button_Schedule5.Enabled = false;
            ReadSch();
        }

        List<byte[]> hexCommandList = new List<byte[]> { };
        int hexCount = 0;
        private void PreProcess(int z)
        {
            //Pre-processing hex
            string columns_command = DataGridView_Schedule.Rows[z].Cells[0].Value.ToString();
            string columns_serial = DataGridView_Schedule.Rows[z].Cells[6].Value.ToString();
            if (columns_command == "_HEX")
            {
                byte[] bytes = { };
                if (columns_serial != "_save" && columns_serial != "_clear" && columns_serial != "")
                {
                    string hexValues = columns_serial;
                    string[] hexValuesSplit = hexValues.Split(' ');
                    hexCount = hexValuesSplit.Count();
                    int index = 0;
                    byte number = 0;
                    bytes = new byte[hexCount];
                    foreach (string hex in hexValuesSplit)
                    {
                        // Convert the number expressed in base-16 to an integer.
                        number = Convert.ToByte(Convert.ToInt32(hex, 16));
                        // Get the character corresponding to the integral value.
                        bytes[index] = number;
                        index++;

                        if (index == hexCount)
                        {
                            hexCommandList.Add(bytes);
                        }
                    }
                }
            }
        }

        private void ReadSch()
        {
            // Console.WriteLine(GlobalData.Schedule_Num);
            // 戴入Schedule CSV 檔
            string SchedulePath = ini12.INIRead(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "Path", "");
            string ScheduleExist = ini12.INIRead(MainSettingPath, "Schedule" + GlobalData.Schedule_Number, "Exist", "");

            //string TextLine = "";
            //string[] SplitLine;
            int i = 0;
            if ((File.Exists(SchedulePath) == true) && ScheduleExist == "1" && IsFileLocked(SchedulePath) == false)
            {
                DataGridView_Schedule.Rows.Clear();
                /*StreamReader objReader = new StreamReader(SchedulePath);
                while ((objReader.Peek() != -1))
                {
                    TextLine = objReader.ReadLine();
                    if (i != 0)
                    {
                        SplitLine = TextLine.Split(',');
                        DataGridView_Schedule.Rows.Add(SplitLine);
                    }
                    i++;
                }
                objReader.Close();*/

                TextFieldParser parser = new TextFieldParser(SchedulePath);
                parser.Delimiters = new string[] { "," };
                string[] parts = new string[11];
                while (!parser.EndOfData)
                {
                    try
                    {
                        parts = parser.ReadFields();
                        if (parts == null)
                        {
                            break;
                        }

                        if (i != 0)
                        {
                            DataGridView_Schedule.Rows.Add(parts);
                        }
                        i++;
                    }
                    catch (MalformedLineException)
                    {
                        MessageBox.Show("Schedule cannot contain double quote ( \" \" ).", "Schedule foramt error");
                    }
                }
                parser.Close();

                int j = parts.Length;
                if ((j == 11 || j == 10))
                {
                    long TotalDelay = 0;        //計算各個schedule測試時間
                    long RepeatTime = 0;
                    button_Start.Enabled = true;
                    for (int z = 0; z < DataGridView_Schedule.Rows.Count - 1; z++)
                    {
                        PreProcess(z);
                        if (DataGridView_Schedule.Rows[z].Cells[8].Value.ToString() != "")
                        {
                            if (DataGridView_Schedule.Rows[z].Cells[1].Value.ToString() != "" && DataGridView_Schedule.Rows[z].Cells[2].Value.ToString() != "")
                                RepeatTime = (long.Parse(DataGridView_Schedule.Rows[z].Cells[1].Value.ToString())) * (long.Parse(DataGridView_Schedule.Rows[z].Cells[2].Value.ToString()));
                            else if (DataGridView_Schedule.Rows[z].Cells[1].Value.ToString() == "" && DataGridView_Schedule.Rows[z].Cells[2].Value.ToString() != "")
                                RepeatTime = (long.Parse("1")) * (long.Parse(DataGridView_Schedule.Rows[z].Cells[2].Value.ToString()));

                            if (DataGridView_Schedule.Rows[z].Cells[8].Value.ToString().Contains('m') == true)
                                TotalDelay += (Convert.ToInt64(DataGridView_Schedule.Rows[z].Cells[8].Value.ToString().Replace('m', ' ').Trim()) * 60000 + RepeatTime);
                            else
                                TotalDelay += (long.Parse(DataGridView_Schedule.Rows[z].Cells[8].Value.ToString()) + RepeatTime);

                            RepeatTime = 0;
                        }
                    }

                    if (ini12.INIRead(MainSettingPath, "Record", "EachVideo", "") == "1")
                    {
                        ConvertToRealTime(((TotalDelay * GlobalData.Schedule_Loop) + 63000));
                    }
                    else
                    {
                        ConvertToRealTime((TotalDelay * GlobalData.Schedule_Loop));
                    }

                    switch (GlobalData.Schedule_Number)
                    {
                        case 1:
                            GlobalData.Schedule_1_TestTime = (TotalDelay * GlobalData.Schedule_Loop);
                            timeCount = GlobalData.Schedule_1_TestTime;
                            break;
                        case 2:
                            GlobalData.Schedule_2_TestTime = (TotalDelay * GlobalData.Schedule_Loop);
                            timeCount = GlobalData.Schedule_2_TestTime;
                            break;
                        case 3:
                            GlobalData.Schedule_3_TestTime = (TotalDelay * GlobalData.Schedule_Loop);
                            timeCount = GlobalData.Schedule_3_TestTime;
                            break;
                        case 4:
                            GlobalData.Schedule_4_TestTime = (TotalDelay * GlobalData.Schedule_Loop);
                            timeCount = GlobalData.Schedule_4_TestTime;
                            break;
                        case 5:
                            GlobalData.Schedule_5_TestTime = (TotalDelay * GlobalData.Schedule_Loop);
                            timeCount = GlobalData.Schedule_5_TestTime;
                            break;
                    }
                }
                else
                {
                    button_Start.Enabled = false;
                    MessageBox.Show("Please check your .csv file format.", "Schedule format error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (IsFileLocked(SchedulePath))
            {
                MessageBox.Show("Please check your .csv file is closed, then press Settings to reload the schedule.", "File lock error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_Start.Enabled = false;
                button_Schedule1.PerformClick();
            }
            else
            {
                button_Start.Enabled = false;
                button_Schedule1.PerformClick();
            }

            //setStyle();
        }
        #endregion

        public static bool IsFileLocked(string file)
        {
            try
            {
                using (File.Open(file, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException exception)
            {
                var errorCode = Marshal.GetHRForException(exception) & 65535;
                return errorCode == 32 || errorCode == 33;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #region -- 測試時間 --
        private string ConvertToRealTime(long ms)
        {
            string sResult = "";
            try
            {
                TimeSpan finishTime = TimeSpan.FromMilliseconds(ms);
                label_ScheduleTime_Value.Invoke((MethodInvoker)(() => label_ScheduleTime_Value.Text = finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms"));
                ini12.INIWrite(MailPath, "Total Test Time", "value", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");

                // 寫入每個Schedule test time
                if (GlobalData.Schedule_Number == 1)
                    ini12.INIWrite(MailPath, "Total Test Time", "value1", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");

                if (StartButtonPressed == true)
                {
                    switch (GlobalData.Schedule_Number)
                    {
                        case 2:
                            ini12.INIWrite(MailPath, "Total Test Time", "value2", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");
                            break;
                        case 3:
                            ini12.INIWrite(MailPath, "Total Test Time", "value3", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");
                            break;
                        case 4:
                            ini12.INIWrite(MailPath, "Total Test Time", "value4", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");
                            break;
                        case 5:
                            ini12.INIWrite(MailPath, "Total Test Time", "value5", finishTime.Days.ToString("0") + "d " + finishTime.Hours.ToString("0") + "h " + finishTime.Minutes.ToString("0") + "m " + finishTime.Seconds.ToString("0") + "s " + finishTime.Milliseconds.ToString("0") + "ms");
                            break;
                    }
                }
            }
            catch
            {
                sResult = "Error!";
            }
            return sResult;
        }
        #endregion

        #region -- UI相關 --
        /*
        #region 陰影
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                if (!DesignMode)
                {
                    cp.ClassStyle |= CS_DROPSHADOW;
                }
                return cp;
            }
        }
        #endregion
        */
        #region -- 關閉、縮小按鈕 --
        private void ClosePicBox_Enter(object sender, EventArgs e)
        {
            ClosePicBox.Image = Properties.Resources.close2;
        }

        private void ClosePicBox_Leave(object sender, EventArgs e)
        {
            ClosePicBox.Image = Properties.Resources.close1;
        }

        private void ClosePicBox_Click(object sender, EventArgs e)
        {
            CloseDtplay();
            CloseAutobox();
        }

        private void MiniPicBox_Enter(object sender, EventArgs e)
        {
            MiniPicBox.Image = Properties.Resources.mini2;
        }

        private void MiniPicBox_Leave(object sender, EventArgs e)
        {
            MiniPicBox.Image = Properties.Resources.mini1;
        }

        private void MiniPicBox_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        #endregion

        #region -- 滑鼠拖曳視窗 --
        private void GPanelTitleBack_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0);        //調用移動無窗體控件函數
        }
        #endregion

        #endregion

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs anError)
        {
            DataGridView_Schedule.CausesValidation = false;
        }

        private void DataBtn_Click(object sender, EventArgs e)            //背景執行填入測試步驟然後匯出reprot>>>>>>>>>>>>>
        {
            //Form_DGV_Autobox.ShowDialog();
        }

        private void PauseButton_Click(object sender, EventArgs e)      //暫停SCHEDULE
        {
            Pause = !Pause;

            if (Pause == true)
            {
                button_Pause.Text = "RESUME";
                button_Start.Enabled = false;
                //setStyle();
                SchedulePause.Reset();

                log.Debug("Datagridview highlight.");
                GridUI(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview highlight//
                log.Debug("Datagridview scollbar.");
                Gridscroll(GlobalData.Scheduler_Row.ToString(), DataGridView_Schedule);//控制Datagridview scollbar//
            }
            else
            {
                button_Pause.Text = "PAUSE";
                button_Start.Enabled = true;
                //setStyle();
                SchedulePause.Set();
                timer_countdown.Start();
            }
        }

        #region -- Schedule Time --
        private void Schedule_Time()        //Estimated schedule time
        {
            if (timeCount > 0)
            {
                if (!String.IsNullOrEmpty(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[8].Value.ToString()))
                {
                    if (DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[1].Value.ToString() != "" && DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[2].Value.ToString() != "")
                        repeatTime = (long.Parse(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[1].Value.ToString())) * (long.Parse(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[2].Value.ToString()));
                    else if (DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[1].Value.ToString() == "" && DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[2].Value.ToString() != "")
                        repeatTime = (long.Parse("1")) * (long.Parse(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[2].Value.ToString()));

                    if (DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[8].Value.ToString().Contains("m") == true)
                        delayTime = (long.Parse(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[8].Value.ToString().Replace('m', ' ').Trim()) * 60000 + repeatTime);
                    else
                        delayTime = (long.Parse(DataGridView_Schedule.Rows[GlobalData.Scheduler_Row].Cells[8].Value.ToString()) + repeatTime);

                    if (GlobalData.Schedule_Step == 0 && GlobalData.Loop_Number == 1)
                    {
                        timeCountUpdated = timeCount - delayTime;
                        ConvertToRealTime(timeCountUpdated);
                    }
                    else
                    {
                        timeCountUpdated = timeCountUpdated - delayTime;
                        ConvertToRealTime(timeCountUpdated);
                    }
                    repeatTime = 0;
                }
            }
        }
        #endregion

        //倒數計時
        private void timer_countdown_Tick(object sender, EventArgs e)
        {
            timer_countdown.Interval = 500;
            TimeSpan timeElapsed = DateTime.Now - startTime;

            /*
            if (timeCount > 0)
            {
                if (Convert.ToInt64(timeElapsed.TotalMilliseconds) <= timeCount)
                {
                    ConvertToRealTime(timeCount - Convert.ToInt64(timeElapsed.TotalMilliseconds));
                }
                else
                {
                    ConvertToRealTime(0);
                }
            }*/

            label_TestTime_Value.Invoke((MethodInvoker)(() => label_TestTime_Value.Text = timeElapsed.Days.ToString("0") + "d " + timeElapsed.Hours.ToString("0") + "h " + timeElapsed.Minutes.ToString("0") + "m " + timeElapsed.Seconds.ToString("0") + "s " + timeElapsed.Milliseconds.ToString("0") + "ms"));
            ini12.INIWrite(MailPath, "Total Test Time", "How Long", timeElapsed.Days.ToString("0") + "d " + timeElapsed.Hours.ToString("0") + "h " + timeElapsed.Minutes.ToString("0") + "m " + timeElapsed.Seconds.ToString("0") + "s" + timeElapsed.Milliseconds.ToString("0") + "ms");
        }

        private void TimerPanelbutton_Click(object sender, EventArgs e)
        {
            TimerPanel = !TimerPanel;

            if (TimerPanel == true)
            {
                panel1.Show();
                panel1.BringToFront();
            }
            else
                panel1.Hide();
        }

        static List<USBDeviceInfo> GetUSBDevices()
        {
            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();

            ManagementObjectCollection collection;
            using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))
                collection = searcher.Get();

            foreach (var device in collection)
            {
                devices.Add(new USBDeviceInfo(
                (string)device.GetPropertyValue("DeviceID"),
                (string)device.GetPropertyValue("PNPDeviceID"),
                (string)device.GetPropertyValue("Description")
                ));
            }

            collection.Dispose();
            return devices;
        }

        class USBDeviceInfo
        {
            public USBDeviceInfo(string deviceID, string pnpDeviceID, string description)
            {
                DeviceID = deviceID;
                PnpDeviceID = pnpDeviceID;
                Description = description;
            }
            public string DeviceID { get; private set; }
            public string PnpDeviceID { get; private set; }
            public string Description { get; private set; }
        }

        //釋放記憶體//
        [System.Runtime.InteropServices.DllImportAttribute("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize", ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Ansi, SetLastError = true)]
        private static extern int SetProcessWorkingSetSize(IntPtr process, int minimumWorkingSetSize, int maximumWorkingSetSize);
        private void DisposeRam()
        {
            GC.Collect();
            GC.SuppressFinalize(this);
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            CloseAutobox();
        }

        private void button_Input_Click(object sender, EventArgs e)
        {
            //UInt32 gpio_input_value;
            //MyBlueRat.Get_GPIO_Input(out gpio_input_value);
            //byte GPIO_Read_Data = Convert.ToByte(gpio_input_value & 0xff);
            //labelGPIO_Input.Text = "GPIO_IN:" + GPIO_Read_Data.ToString();
            //Console.WriteLine("GPIO_IN:" + GPIO_Read_Data.ToString());

            UInt32 GPIO_input_value, retry_cnt;
            bool bRet = false;
            retry_cnt = 3;
            do
            {
                String modified0 = "";
                bRet = MyBlueRat.Get_GPIO_Input(out GPIO_input_value);

                if (GPIO_input_value == 31)
                {
                    modified0 = "0" + Convert.ToString(31, 2);
                }
                else
                {
                    modified0 = Convert.ToString(GPIO_input_value, 2);
                }

                string modified1 = modified0.Insert(1, ",");
                string modified2 = modified1.Insert(3, ",");
                string modified3 = modified2.Insert(5, ",");
                string modified4 = modified3.Insert(7, ",");
                string modified5 = modified4.Insert(9, ",");

                GlobalData.IO_INPUT = modified5;
                Console.WriteLine(GlobalData.IO_INPUT);
                Console.WriteLine(GlobalData.IO_INPUT.Substring(0, 1));
            }
            while ((bRet == false) && (--retry_cnt > 0));

            if (bRet)
            {
                labelGPIO_Input.Text = "GPIO_input: " + GPIO_input_value.ToString();
            }
            else
            {
                labelGPIO_Input.Text = "GPIO_input fail after retry";
            }
        }

        private void button_Output_Click(object sender, EventArgs e)
        {
            //string GPIO = "01010101";
            //byte GPIO_B = Convert.ToByte(GPIO, 2);
            //MyBlueRat.Set_GPIO_Output(GPIO_B);

            Graphics graphics = this.CreateGraphics();
            Console.WriteLine("dpiX = " + graphics.DpiX);
            Console.WriteLine("dpiY = " + graphics.DpiY);
            Console.WriteLine("-----------");
            Console.WriteLine("height = " + this.Size.Height);
            Console.WriteLine("width = " + this.Size.Width);
        }

        #region -- GPIO --
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool DeviceIoControl(SafeFileHandle hDevice,
                                                   uint dwIoControlCode,
                                                   ref uint InBuffer,
                                                   int nInBufferSize,
                                                   byte[] OutBuffer,
                                                   UInt32 nOutBufferSize,
                                                   ref UInt32 out_count,
                                                   IntPtr lpOverlapped);
        public SafeFileHandle hCOM;

        public const uint FILE_DEVICE_UNKNOWN = 0x00000022;
        public const uint USB2SER_IOCTL_INDEX = 0x0800;
        public const uint METHOD_BUFFERED = 0;
        public const uint FILE_ANY_ACCESS = 0;

        public bool PowerState;
        public bool USBState;

        public static uint GP0_SET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 22, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP1_SET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 23, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP2_SET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 47, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP3_SET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 48, METHOD_BUFFERED, FILE_ANY_ACCESS);

        public static uint GP0_GET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 24, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP1_GET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 25, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP2_GET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 49, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP3_GET_VALUE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 50, METHOD_BUFFERED, FILE_ANY_ACCESS);

        public static uint GP0_OUTPUT_ENABLE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 20, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP1_OUTPUT_ENABLE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 21, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP2_OUTPUT_ENABLE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 45, METHOD_BUFFERED, FILE_ANY_ACCESS);
        public static uint GP3_OUTPUT_ENABLE = CTL_CODE(FILE_DEVICE_UNKNOWN, USB2SER_IOCTL_INDEX + 46, METHOD_BUFFERED, FILE_ANY_ACCESS);

        static uint CTL_CODE(uint DeviceType, uint Function, uint Method, uint Access)
        {
            return ((DeviceType << 16) | (Access << 14) | (Function << 2) | Method);
        }

        #region -- GP0 --
        public bool PL2303_GP0_Enable(SafeFileHandle hDrv, uint enable)
        {
            UInt32 nBytes = 0;
            bool bSuccess = DeviceIoControl(hDrv, GP0_OUTPUT_ENABLE,
            ref enable, sizeof(byte), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        public bool PL2303_GP0_SetValue(SafeFileHandle hDrv, uint val)
        {
            UInt32 nBytes = 0;
            byte[] addr = new byte[6];
            bool bSuccess = DeviceIoControl(hDrv, GP0_SET_VALUE, ref val, sizeof(uint), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        #endregion

        #region -- GP1 --
        public bool PL2303_GP1_Enable(SafeFileHandle hDrv, uint enable)
        {
            UInt32 nBytes = 0;
            bool bSuccess = DeviceIoControl(hDrv, GP1_OUTPUT_ENABLE,
            ref enable, sizeof(byte), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        public bool PL2303_GP1_SetValue(SafeFileHandle hDrv, uint val)
        {
            UInt32 nBytes = 0;
            byte[] addr = new byte[6];
            bool bSuccess = DeviceIoControl(hDrv, GP1_SET_VALUE, ref val, sizeof(uint), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        #endregion

        #region -- GP2 --
        public bool PL2303_GP2_Enable(SafeFileHandle hDrv, uint enable)
        {
            UInt32 nBytes = 0;
            bool bSuccess = DeviceIoControl(hDrv, GP2_OUTPUT_ENABLE,
            ref enable, sizeof(byte), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        public bool PL2303_GP2_SetValue(SafeFileHandle hDrv, uint val)
        {
            UInt32 nBytes = 0;
            byte[] addr = new byte[6];
            bool bSuccess = DeviceIoControl(hDrv, GP2_SET_VALUE, ref val, sizeof(uint), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        #endregion

        #region -- GP3 --
        public bool PL2303_GP3_Enable(SafeFileHandle hDrv, uint enable)
        {
            UInt32 nBytes = 0;
            bool bSuccess = DeviceIoControl(hDrv, GP3_OUTPUT_ENABLE,
            ref enable, sizeof(byte), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        public bool PL2303_GP3_SetValue(SafeFileHandle hDrv, uint val)
        {
            UInt32 nBytes = 0;
            byte[] addr = new byte[6];
            bool bSuccess = DeviceIoControl(hDrv, GP3_SET_VALUE, ref val, sizeof(uint), null, 0, ref nBytes, IntPtr.Zero);
            return bSuccess;
        }
        #endregion

        private void GP0_GP1_AC_ON()
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;
            uint val = (uint)int.Parse("1");
            try
            {
                bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);

                bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);
            }
            catch (Exception)
            {
                MessageBox.Show("Woodpecker is already running.", "GP0_GP1_AC_ON Error");
            }
            PowerState = true;
            pictureBox_AcPower.Image = Properties.Resources.ON;
        }

        private void GP0_GP1_AC_OFF_ON()
        {
            if (StartButtonPressed == true)
            {
                // 電源開或關
                byte[] val1;
                val1 = new byte[2];
                val1[0] = 0;

                bool Success_GP0_Enable = PL2303_GP0_Enable(hCOM, 1);
                bool Success_GP1_Enable = PL2303_GP1_Enable(hCOM, 1);
                if (Success_GP0_Enable && Success_GP1_Enable && PowerState == false)
                {
                    uint val;
                    val = (uint)int.Parse("1");
                    bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);
                    bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);
                    if (Success_GP0_SetValue && Success_GP1_SetValue)
                    {
                        {
                            PowerState = true;
                            pictureBox_AcPower.Image = Properties.Resources.ON;
                        }
                    }
                }
                else if (Success_GP0_Enable && Success_GP1_Enable && PowerState == true)
                {
                    uint val;
                    val = (uint)int.Parse("0");
                    bool Success_GP0_SetValue = PL2303_GP0_SetValue(hCOM, val);
                    bool Success_GP1_SetValue = PL2303_GP1_SetValue(hCOM, val);
                    if (Success_GP0_SetValue && Success_GP1_SetValue)
                    {
                        {
                            PowerState = false;
                            pictureBox_AcPower.Image = Properties.Resources.OFF;
                        }
                    }
                }
            }
        }

        private void GP2_GP3_USB_PC()
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;
            uint val = (uint)int.Parse("0");

            try
            {
                bool Success_GP2_Enable = PL2303_GP2_Enable(hCOM, 1);
                bool Success_GP2_SetValue = PL2303_GP2_SetValue(hCOM, val);

                bool Success_GP3_Enable = PL2303_GP3_Enable(hCOM, 1);
                bool Success_GP3_SetValue = PL2303_GP3_SetValue(hCOM, val);
            }
            catch (Exception)
            {
                MessageBox.Show("Woodpecker is already running.", "GP2_GP3_USB_PC Error");
            }
            USBState = true;
        }

        private void IO_INPUT()
        {
            UInt32 GPIO_input_value, retry_cnt;
            bool bRet = false;
            retry_cnt = 3;
            do
            {
                String modified0 = "";
                bRet = MyBlueRat.Get_GPIO_Input(out GPIO_input_value);
                if (Convert.ToString(GPIO_input_value, 2).Length == 5)
                {
                    modified0 = "0" + Convert.ToString(GPIO_input_value, 2);
                }
                else if (Convert.ToString(GPIO_input_value, 2).Length == 4)
                {
                    modified0 = "0" + "0" + Convert.ToString(GPIO_input_value, 2);
                }
                else if (Convert.ToString(GPIO_input_value, 2).Length == 3)
                {
                    modified0 = "0" + "0" + "0" + Convert.ToString(GPIO_input_value, 2);
                }
                else if (Convert.ToString(GPIO_input_value, 2).Length == 2)
                {
                    modified0 = "0" + "0" + "0" + "0" + Convert.ToString(GPIO_input_value, 2);
                }
                else if (Convert.ToString(GPIO_input_value, 2).Length == 1)
                {
                    modified0 = "0" + "0" + "0" + "0" + "0" + Convert.ToString(GPIO_input_value, 2);
                }
                else
                {
                    modified0 = Convert.ToString(GPIO_input_value, 2);
                }

                string modified1 = modified0.Insert(1, ",");
                string modified2 = modified1.Insert(3, ",");
                string modified3 = modified2.Insert(5, ",");
                string modified4 = modified3.Insert(7, ",");
                string modified5 = modified4.Insert(9, ",");

                GlobalData.IO_INPUT = modified5;
            }
            while ((bRet == false) && (--retry_cnt > 0));

            if (bRet)
            {
                labelGPIO_Input.Text = "GPIO_input: " + GPIO_input_value.ToString();
            }
            else
            {
                labelGPIO_Input.Text = "GPIO_input fail after retry";
            }
        }

        private void Arduino_IO_INPUT(int delay_time = 1000)
        {
            UInt32 GPIO_input_value;
            bool aGpio = false;
            String low_binary_value = "";
            String high_binary_value = "";
            String binary_dot_value = "";
            aGpio = Arduino_Get_GPIO_Input(out GPIO_input_value, delay_time);
            if (aGpio)
            {
                if (GPIO_input_value < 0x100)
                {
                    low_binary_value = Convert.ToString(GPIO_input_value, 2).PadLeft(8, '0');
                    for (int i = 0; i < low_binary_value.Length; i++)
                        binary_dot_value = binary_dot_value + low_binary_value.Substring(i, 1) + ",";
                }
                else
                {
                    high_binary_value = Convert.ToString(GPIO_input_value & 0xFF00, 2).PadLeft(16, '0');
                    for (int i = 0; i <= high_binary_value.Length - 8; i++)
                    {
                        if (high_binary_value.Substring(i, 1) == "1")
                        {
                            string DebugValue = "adc" + "\r\n";
                            GlobalData.m_Arduino_Port.WriteDataOut(DebugValue, DebugValue.Length);
                            MessageBox.Show("Please check the Arduino-P0" + (9 - i) + " :status. Maybe voltage have issue.", "Arduino-PIN Error!");
                        }
                    }
                    binary_dot_value = "Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,";
                }
                GlobalData.Arduino_IO_INPUT = binary_dot_value;
                GlobalData.Arduino_IO_INPUT_value = GPIO_input_value;
            }
            else
            {
                GlobalData.Arduino_IO_INPUT = "Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,Undefine,";
                GlobalData.Arduino_IO_INPUT_value = 0x10000;
            }

            if (aGpio)
            {
                labelGPIO_Input.Text = "Arduino_GPIO_input: " + GPIO_input_value.ToString();
            }
            else
            {
                labelGPIO_Input.Text = "Arduino_GPIO_input fail after retry";
            }

            string dataValue = "Arduino_GPIO_INPUT=" + GlobalData.Arduino_IO_INPUT_value;
            if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
            {
                DateTime dt = DateTime.Now;
                dataValue = "[Receive_Port_Arduino_IO_INPUT] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
            }
            logDumpping.LogCat(ref arduino_text, dataValue);
            logDumpping.LogCat(ref logAll_text, dataValue);
        }

        public bool Arduino_Get_GPIO_Input(out UInt32 GPIO_Read_Data, int delay_time)
        {
            bool aGpio = false;
            uint retry_cnt = 5;
            GPIO_Read_Data = 0x10000;

            if (GlobalData.m_Arduino_Port.IsOpen())
            {
                try
                {
                    string dataValue = "io i" + "\r\n";
                    GlobalData.Arduino_recFlag = false;
                    do
                    {
                        GlobalData.m_Arduino_Port.WriteDataOut(dataValue, dataValue.Length);
                        retry_cnt--;
                        Thread.Sleep(300);
                        if (GlobalData.Arduino_recFlag == false && retry_cnt == 0)
                        {
                            MessageBox.Show("Arduino response input timeout and please replug the Arduino board.", "Connection Error");
                            aGpio = false;
                        }
                        else if (GlobalData.Arduino_recFlag && GlobalData.Arduino_Read_String != "")
                        {
                            string l_strResult = GlobalData.Arduino_Read_String.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("ioi", "");
                            //GPIO_Read_Data = Convert.ToUInt32(l_strResult);
                            GPIO_Read_Data = Convert.ToUInt16(l_strResult, 16);
                            aGpio = true;
                        }
                    }
                    while (GlobalData.Arduino_recFlag == false && retry_cnt > 0);
                }
                catch (System.FormatException)
                {

                }
            }
            else
            {
                if (StartButtonPressed == true)
                {
                    MessageBox.Show("Arduino didn't connected!\r\nPlease replug the Arduino board and restart the Woodpecker.", "Connection Error");
                    button_Start.PerformClick();
                }
            }

            return aGpio;
        }

        public bool Arduino_Set_GPIO_Output(byte output_value, int delay_time)
        {
            bool aGpio = false;
            uint retry_cnt = 5;

            if (GlobalData.m_Arduino_Port.IsOpen())
            {
                try
                {
                    string dataValue = "io x " + output_value + "\r\n";
                    GlobalData.Arduino_recFlag = false;
                    do
                    {
                        GlobalData.m_Arduino_Port.WriteDataOut(dataValue, dataValue.Length);
                        retry_cnt--;
                        Thread.Sleep(300);
                        if (GlobalData.Arduino_recFlag == false && retry_cnt == 0)
                        {
                            MessageBox.Show("Arduino response output timeout and please replug the Arduino board.", "Connection Error");
                            aGpio = false;
                        }
                        else if (GlobalData.Arduino_recFlag)
                        {
                            if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                            {
                                DateTime dt = DateTime.Now;
                                dataValue = "[Send_Port_Arduino_IO_OUTPUT] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                            }
                            logDumpping.LogCat(ref arduino_text, dataValue);
                            logDumpping.LogCat(ref logAll_text, dataValue);
                            aGpio = true;
                        }
                    }
                    while (GlobalData.Arduino_recFlag == false && retry_cnt > 0);
                }
                catch (System.FormatException)
                {

                }
            }
            else
            {
                if (StartButtonPressed == true)
                {
                    MessageBox.Show("Arduino didn't connected!\r\nPlease replug the Arduino board and restart the Woodpecker.", "Connection Error");
                    button_Start.PerformClick();
                }
            }
            
            return aGpio;
        }

        public bool Arduino_Set_Value(string input_value, int delay_time)
        {
            bool aGpio = false;
            uint retry_cnt = 5;

            if (GlobalData.m_Arduino_Port.IsOpen())
            {
                try
                {
                    string dataValue = input_value + "\r\n";
                    GlobalData.Arduino_recFlag = false;
                    do
                    {
                        GlobalData.m_Arduino_Port.WriteDataOut(dataValue, dataValue.Length);
                        retry_cnt--;
                        Thread.Sleep(300);
                        if (GlobalData.Arduino_recFlag == false && retry_cnt == 0)
                        {
                            MessageBox.Show("Arduino response output timeout and please replug the Arduino board.", "Connection Error");
                            aGpio = false;
                        }
                        else if (GlobalData.Arduino_recFlag)
                        {
                            if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                            {
                                DateTime dt = DateTime.Now;
                                dataValue = "[Send_Port_Arduino_Command] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + dataValue + "\r\n"; //OK
                            }
                            logDumpping.LogCat(ref arduino_text, dataValue);
                            logDumpping.LogCat(ref logAll_text, dataValue);
                            aGpio = true;
                        }
                    }
                    while (GlobalData.Arduino_recFlag == false && retry_cnt > 0);
                }
                catch (System.FormatException)
                {

                }
            }
            else
            {
                if (StartButtonPressed == true)
                {
                    MessageBox.Show("Arduino didn't connected!\r\nPlease replug the Arduino board and restart the Woodpecker.", "Connection Error");
                    button_Start.PerformClick();
                }
            }

            return aGpio;
        }
        #endregion

        private void button_VirtualRC_Click(object sender, EventArgs e)
        {
            /*
            VirtualRcPanel = !VirtualRcPanel;
            if (VirtualRcPanel == true)
            {
                LoadVirtualRC();
                panel_VirtualRC.Show();
                panel_VirtualRC.BringToFront();
            }
            else
            {
                panel_VirtualRC.Controls.Clear();
                panel_VirtualRC.Hide();
            }
            */
            FormRC formRC = new FormRC();
            formRC.Owner = this;
            if (GlobalData.FormRC == false)
            {
                formRC.Show();
            }
        }

        private void button_VirtualGPIO_Click(object sender, EventArgs e)
        {
            FormGPIO formGPIO = new FormGPIO();
            formGPIO.Owner = this;
            if (GlobalData.FormGPIO == false)
            {
                formGPIO.Show();
            }
        }

        private void button_AcUsb_Click(object sender, EventArgs e)
        {
            AcUsbPanel = !AcUsbPanel;

            if (AcUsbPanel == true)
            {
                panel_AcUsb.Show();
                panel_AcUsb.BringToFront();
            }
            else
            {
                panel_AcUsb.Hide();
            }
        }

        private void pictureBox_Ac1_Click(object sender, EventArgs e)
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;

            bool jSuccess = PL2303_GP0_Enable(hCOM, 1);
            if (PowerState == false) //Set GPIO Value as 1
            {
                uint val;
                val = (uint)int.Parse("1");
                bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        PowerState = true;
                        pictureBox_Ac1.Image = Properties.Resources.Switch_On_AC;
                    }
                }
            }
            else if (PowerState == true) //Set GPIO Value as 0
            {
                uint val;
                val = (uint)int.Parse("0");
                bool bSuccess = PL2303_GP0_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        PowerState = false;
                        pictureBox_Ac1.Image = Properties.Resources.Switch_Off_AC;
                    }
                }
            }
        }

        private void pictureBox_Ac2_Click(object sender, EventArgs e)
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;

            bool jSuccess = PL2303_GP1_Enable(hCOM, 1);
            if (PowerState == false) //Set GPIO Value as 1
            {
                uint val;
                val = (uint)int.Parse("1");
                bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        PowerState = true;
                        pictureBox_Ac2.Image = Properties.Resources.Switch_On_AC;
                    }
                }
            }
            else if (PowerState == true) //Set GPIO Value as 0
            {
                uint val;
                val = (uint)int.Parse("0");
                bool bSuccess = PL2303_GP1_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        PowerState = false;
                        pictureBox_Ac2.Image = Properties.Resources.Switch_Off_AC;
                    }
                }
            }
        }

        private void pictureBox_Usb1_Click(object sender, EventArgs e)
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;

            bool jSuccess = PL2303_GP2_Enable(hCOM, 1);
            if (USBState == true) //Set GPIO Value as 1
            {
                uint val;
                val = (uint)int.Parse("1");
                bool bSuccess = PL2303_GP2_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        USBState = false;
                        pictureBox_Usb1.Image = Properties.Resources.Switch_to_TV;
                    }
                }
            }
            else if (USBState == false) //Set GPIO Value as 0
            {
                uint val;
                val = (uint)int.Parse("0");
                bool bSuccess = PL2303_GP2_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        USBState = true;
                        pictureBox_Usb1.Image = Properties.Resources.Switch_to_PC;
                    }
                }
            }
        }

        private void pictureBox_Usb2_Click(object sender, EventArgs e)
        {
            byte[] val1 = new byte[2];
            val1[0] = 0;

            bool jSuccess = PL2303_GP3_Enable(hCOM, 1);
            if (USBState == true) //Set GPIO Value as 1
            {
                uint val;
                val = (uint)int.Parse("1");
                bool bSuccess = PL2303_GP3_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        USBState = false;
                        pictureBox_Usb2.Image = Properties.Resources.Switch_to_TV;
                    }
                }
            }
            else if (USBState == false) //Set GPIO Value as 0
            {
                uint val;
                val = (uint)int.Parse("0");
                bool bSuccess = PL2303_GP3_SetValue(hCOM, val);
                if (bSuccess == true)
                {
                    {
                        USBState = true;
                        pictureBox_Usb2.Image = Properties.Resources.Switch_to_PC;
                    }
                }
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (ini12.INIRead(MainSettingPath, "Device", "RunAfterStartUp", "") == "1")
            {
                button_Start.PerformClick();
            }
        }

        private string strValue;
        public string StrValue
        {
            set
            {
                strValue = value;
            }
        }

        private void comboBox_CameraDevice_SelectedIndexChanged(object sender, EventArgs e)
        {
            ini12.INIWrite(MainSettingPath, "Camera", "VideoIndex", comboBox_CameraDevice.SelectedIndex.ToString());
            if (_captureInProgress == true)
            {
                capture.Stop();
                capture.Dispose();
                Camstart();
            }
        }

        private void DataGridView_Schedule_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Nowpoint = DataGridView_Schedule.Rows[e.RowIndex].Index;

            if (Breakfunction == true && Nowpoint != Breakpoint)
            {
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 51);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionBackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionForeColor = Color.White;
                DataGridView_Schedule.Rows[Nowpoint].DefaultCellStyle.BackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Nowpoint].DefaultCellStyle.SelectionBackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Nowpoint].DefaultCellStyle.SelectionForeColor = Color.White;
                Breakpoint = Nowpoint;
                //Console.WriteLine("Change the Nowpoint");
            }
            else if (Breakfunction == true && Nowpoint == Breakpoint)
            {
                Breakfunction = false;
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 51);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionBackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionForeColor = Color.White;
                DataGridView_Schedule.Rows[Nowpoint].DefaultCellStyle.SelectionBackColor = Color.FromArgb(51, 51, 51);
                DataGridView_Schedule.Rows[Nowpoint].DefaultCellStyle.SelectionForeColor = Color.White;
                Breakpoint = -1;
                //Console.WriteLine("Disable the Breakfunction");
            }
            else
            {
                Breakfunction = true;
                Breakpoint = Nowpoint;
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.BackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionBackColor = Color.FromArgb(3, 218, 198);
                DataGridView_Schedule.Rows[Breakpoint].DefaultCellStyle.SelectionForeColor = Color.White;
                //Console.WriteLine("Enable the Breakfunction");
            }
        }

        private void DataGridView_Schedule_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                FormScriptHelper formScriptHelper = new FormScriptHelper();
                formScriptHelper.Owner = this;

                try
                {
                    if (DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_cmd" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == "Picture" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_cmd" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == "AC/USB Switch" ||

                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_ascii" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">COM  >Pin" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_HEX" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">COM  >Pin" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_HEX" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == "Function" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_ascii" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == "AC/USB Switch" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_ascii" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_HEX" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd" ||

                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Pin" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">COM  >Pin" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Pin" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd" ||

                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Arduino_Pin" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">COM  >Pin" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Arduino_Pin" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Condition_OR" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == "Function" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Condition_OR" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd" ||

                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_Execute" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd")
                    {
                        formScriptHelper.RCKeyForm1 = DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString();
                        formScriptHelper.SetValue(DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText);
                        formScriptHelper.ShowDialog();

                        DataGridView_Schedule[DataGridView_Schedule.CurrentCell.ColumnIndex,
                                              DataGridView_Schedule.CurrentCell.RowIndex].Value = strValue;
                        DataGridView_Schedule.RefreshEdit();
                    }

                    if (DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString().Length >= 8)
                    {
                        if (DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_keyword" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">Times >Keyword#" ||
                        DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString() == "_keyword" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">SerialPort                   >I/O cmd")
                        {
                            formScriptHelper.RCKeyForm1 = DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString();
                            formScriptHelper.SetValue(DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText);
                            formScriptHelper.ShowDialog();

                            DataGridView_Schedule[DataGridView_Schedule.CurrentCell.ColumnIndex,
                                                  DataGridView_Schedule.CurrentCell.RowIndex].Value = strValue;
                            DataGridView_Schedule.RefreshEdit();
                        }

                        if (DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString().Substring(0, 10) == "_IO_Output" &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">Times >Keyword#")
                        {
                            DataGridViewTextBoxColumn targetColumn = (DataGridViewTextBoxColumn)DataGridView_Schedule.Columns[e.ColumnIndex];
                            targetColumn.MaxInputLength = 8;
                        }

                        if ((DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString().Substring(0, 10) == "_WaterTemp" || DataGridView_Schedule.Rows[e.RowIndex].Cells[0].Value.ToString().Substring(0, 12) == "_FuelDisplay") &&
                        DataGridView_Schedule.Columns[e.ColumnIndex].HeaderText == ">Times >Keyword#")
                        {
                            DataGridViewTextBoxColumn targetColumn = (DataGridViewTextBoxColumn)DataGridView_Schedule.Columns[e.ColumnIndex];
                            targetColumn.MaxInputLength = 9;
                        }
                    }
                    strValue = "";
                }
                catch (Exception error)
                {
                    Console.WriteLine(error);
                }
            }
        }

        private void button_Network_Click(object sender, EventArgs e)
        {
            string ip = ini12.INIRead(MainSettingPath, "Network", "IP", "");
            int port = int.Parse(ini12.INIRead(MainSettingPath, "Network", "Port", ""));

            Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            socket.Connect(ip, port); // 1.設定 IP:Port 2.連線至伺服器
            NetworkStream stream = new NetworkStream(socket);
            StreamReader sr = new StreamReader(stream);
            StreamWriter sw = new StreamWriter(stream);

            sw.WriteLine("你好伺服器，我是客戶端。"); // 將資料寫入緩衝
            sw.Flush(); // 刷新緩衝並將資料上傳到伺服器

            Console.WriteLine("從伺服器接收的資料： " + sr.ReadLine());

            Console.ReadLine();

            Process p = new Process();
            string cmd = ini12.INIRead(MainSettingPath, "Python", "Parameter", "");
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = @"python.exe";
            p.StartInfo.Arguments = cmd;
            p.StartInfo.RedirectStandardInput = true;
            p.Start();
            StreamWriter myStreamWriter = p.StandardInput;
            myStreamWriter.WriteLine(cmd);
            string output = "";
            output = p.StandardOutput.ReadLine();
            Console.WriteLine(output);
            p.WaitForExit();
            p.Close();
        }

        private void button_savelog_Click(object sender, EventArgs e)
        {
            string save_option = comboBox_savelog.Text;
            switch (save_option)
            {
                case "Port A":
                    //Serialportsave("A");
                    logDumpping.LogDumpToFile(serialPortConfig_A, GlobalData.portConfigGroup_A.portName, ref logA_text);
                    //Console.WriteLine("[YFC]logA_text: " + logA_text);
                    MessageBox.Show("Port A is saved.", "Reminder");
                    break;
                case "Port B":
                    //Serialportsave("B");
                    logDumpping.LogDumpToFile(serialPortConfig_B, GlobalData.portConfigGroup_B.portName, ref logB_text);
                    //Console.WriteLine("[YFC]logB_text: " + logB_text);
                    MessageBox.Show("Port B is saved.", "Reminder");
                    break;
                case "Port C":
                    logDumpping.LogDumpToFile(serialPortConfig_C, GlobalData.portConfigGroup_C.portName, ref logC_text);
                    MessageBox.Show("Port C is saved.", "Reminder");
                    break;
                case "Port D":
                    logDumpping.LogDumpToFile(serialPortConfig_D, GlobalData.portConfigGroup_D.portName, ref logD_text);
                    MessageBox.Show("Port D is saved.", "Reminder");
                    break;
                case "Port E":
                    logDumpping.LogDumpToFile(serialPortConfig_E, GlobalData.portConfigGroup_E.portName, ref logE_text);
                    MessageBox.Show("Port E is saved.", "Reminder");
                    break;
                case "Minolta":
                    logDumpping.LogDumpToFile("Minolta", "USB", ref minolta_text);
                    MessageBox.Show("Minolta is saved.", "Reminder");
                    break;
				case "Arduino":
                    logDumpping.LogDumpToFile("Arduino", "USB", ref arduino_text);
                    MessageBox.Show("Arduino is saved.", "Reminder");
                    break;
                case "Canbus":
                    logDumpping.LogDumpToFile("Canbus", "USB", ref canbus_text);
                    MessageBox.Show("Canbus is saved.", "Reminder");
                    break;
                case "Kline":
                    logDumpping.LogDumpToFile("Kline", "USB", ref kline_text);
                    MessageBox.Show("Kline Port is saved.", "Reminder");
                    break;
                case "Port All":
                    logDumpping.LogDumpToFile("All", "All", ref logAll_text);
                    MessageBox.Show("All Port is saved.", "Reminder");
                    break;
                default:
                    break;
            }
        }

        private void vectorcanloop()
        {
            UInt64 CAN_Count = 0;
            while (set_timer_rate)
            {
                Console.WriteLine("Vector_Can_Send (Repeat): " + CAN_Count + " times.");
                uint columns_times = can_id;
                byte[] columns_serial = can_data[columns_times];
                int columns_interval = (int)can_rate[columns_times];
                Can_1630A.LoopCANTransmit(columns_times, (uint)columns_interval, columns_serial);

                string Outputstring = "ID: 0x";
                //Outputstring += columns_times + " Data: " + columns_serial;
                DateTime dt = DateTime.Now;
                string canbus_log_text = "[Send_Canbus] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + Outputstring + "\r\n";
                logDumpping.LogCat(canbus_text, canbus_log_text);
                logDumpping.LogCat(logAll_text, canbus_log_text);
                CAN_Count++;
            }

        }

        unsafe private void timer_canbus_Tick(object sender, EventArgs e)
        {
            UInt32 res = new UInt32();
            res = Can_Usb2C.ReceiveData();
            USB_CAN_Process usb_can_2c = new USB_CAN_Process();

            if (can_send == 1)
            {
                if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "UsbCAN")
                {
                    foreach (var can in can_data_list)
                    {
                        usb_can_2c.CAN_Write_Queue_Add(can);
                    }
                    usb_can_2c.CAN_Write_Queue_SendData();
                    can_data_list.Clear();
                }
                else if (ini12.INIRead(MainSettingPath, "Device", "CAN1630AExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "Vector")
                {
                    foreach (var can in can_data_list)
                    {
                        Can_1630A.CAN_Write_Queue_Add(can);
                    }
                    Can_1630A.CAN_Write_Queue_SendData();
                    can_data_list.Clear();
                }
            }
            else
            {
                usb_can_2c.CAN_Write_Queue_Clear();
                Can_1630A.CAN_Write_Queue_Clear();
            }

            if (ini12.INIRead(MainSettingPath, "Device", "UsbCANExist", "") == "1" && ini12.INIRead(MainSettingPath, "Canbus", "Device", "") == "UsbCAN")
            {
                if (res == 0)
                {
                    if (res >= CAN_USB2C.MAX_CAN_OBJ_ARRAY_LEN)     // Must be something wrong
                    {
                        timer_canbus.Enabled = false;
                        Can_Usb2C.StopCAN();
                        Can_Usb2C.Disconnect();

                        pictureBox_can.Image = Properties.Resources.OFF;

                        ini12.INIWrite(MainSettingPath, "Device", "UsbCANExist", "0");

                        return;
                    }
                    return;
                }
                else
                {
                    uint ID = 0, DLC = 0;
                    const int DATA_LEN = 8;
                    byte[] DATA = new byte[DATA_LEN];

                    String str = "";
                    for (UInt32 i = 0; i < res; i++)
                    {
                        DateTime.Now.ToShortTimeString();
                        DateTime dt = DateTime.Now;
                        Can_Usb2C.GetOneCommand(i, out str, out ID, out DLC, out DATA);
                        string canbus_log_text = "[Receive_Canbus] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + str + "\r\n";
                        logDumpping.LogCat(canbus_text, canbus_log_text);
                        logDumpping.LogCat(logAll_text, canbus_log_text);
                        if (Can_Usb2C.ReceiveData() >= CAN_USB2C.MAX_CAN_OBJ_ARRAY_LEN)
                        {
                            timer_canbus.Enabled = false;
                            Can_Usb2C.StopCAN();
                            Can_Usb2C.Disconnect();
                            pictureBox_can.Image = Properties.Resources.OFF;
                            ini12.INIWrite(MainSettingPath, "Device", "UsbCANExist", "0");
                            return;
                        }
                    }
                }
            }
        }

/*
        private void timer_ca310_Tick(object sender, EventArgs e)
        {
            if (ini12.INIRead(GlobalData.MainSettingPath, "Device", "CA310Exist", "") == "1" && isMsr == true)
            {
                try
                {
                    objCa.Measure();
                    string str = " Lv:" + objProbe.Lv.ToString("##0.0000") +
                                 " Sx:" + objProbe.sx.ToString("0.000000") +
                                 " Sy:" + objProbe.sy.ToString("0.000000") +
                                 " T:" + objProbe.T.ToString("####") +
                                 " Duv:" + objProbe.duv.ToString("0.000000") +
                                 " R:" + objProbe.R.ToString("##0.00") +
                                 " G:" + objProbe.G.ToString("##0.00") +
                                 " B:" + objProbe.B.ToString("##0.00");
                    DateTime.Now.ToShortTimeString();
                    DateTime dt = DateTime.Now;
                    string ca310_log_text = "[Receive_CA310] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + str + "\r\n";
                    logDumpping.LogCat(ref ca310_text, ca310_log_text);
                    logDumpping.LogCat(ref logAll_text, ca310_log_text);
                }
                catch (Exception)
                {
                    isMsr = false;
                    timer_ca310.Enabled = false;
                    pictureBox_ca310.Image = Properties.Resources.OFF;
                    MessageBox.Show("CA310 already disconnected, please restart the Woodpecker.", "CA310 Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
*/

        string chamberCommandLog = string.Empty;
        bool chamberTimer_IsTick = false;

        private void timer_Chamber_Tick(object sender, EventArgs e)
        {
            if (ini12.INIRead(MainSettingPath, "Port A", "Checked", "") == "1")
            {
                string chamberCommand = "01 03 00 00 00 02 C4 0B"; //Read chamber temperature
                byte[] Outputbytes = new byte[chamberCommand.Split(' ').Count()];
                Outputbytes = HexConverter.StrToByte(chamberCommand);
                PortA.Write(Outputbytes, 0, Outputbytes.Length); //Send command
                chamberTimer_IsTick = true;

                DateTime dt = DateTime.Now;
                chamberCommandLog = "[Send_Port_A] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + chamberCommand + "\r\n";
                logA_text = string.Concat(logA_text, chamberCommandLog); // Save log for sending chamber command
            }
        }

        //Select & copy log from textbox
        private void Button_Copy_Click(object sender, EventArgs e)
        {
            /*string fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
            System.Diagnostics.Process CANLog = new System.Diagnostics.Process();
            System.Diagnostics.Process.Start(Application.StartupPath + @"\Canlog\CANLog.exe", fName);

            uint canBusStatus;
            canBusStatus = Can_Usb2C.Connect();

            if (GlobalData.TEXTBOX_FOCUS == 1)
            {
                if (textBox_serial.SelectionLength == 0) //Determine if any text is selected in the TextBox control.
                {
                    CopyLog(textBox_serial);
                }
            }
            else if (GlobalData.TEXTBOX_FOCUS == 2)
            {
                if (textBox2.SelectionLength == 0)
                {
                    CopyLog(textBox2);
                }
            }
            else if (GlobalData.TEXTBOX_FOCUS == 3)
            {
                if (textBox_canbus.SelectionLength == 0)
                {
                    CopyLog(textBox3);
                }
            }
            else if (GlobalData.TEXTBOX_FOCUS == 4)
            {
                CopyLog(textBox_kline);
            }

            //copy schedule log (might be removed in near future)
            else if (GlobalData.TEXTBOX_FOCUS == 5)
            {
                CopyLog(textBox_canbus);
            }

            else if (GlobalData.TEXTBOX_FOCUS == 6)
            {
                CopyLog(textBox_TestLog);
                string fName = "";

                // 讀取ini中的路徑
                fName = ini12.INIRead(MainSettingPath, "Record", "LogPath", "");
                string t = fName + "\\CanbusLog_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";

                StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                MYFILE.Write(textBox_TestLog.Text);
                MYFILE.Close();

                System.Diagnostics.Process CANLog = new System.Diagnostics.Process();
                System.Diagnostics.Process.Start(Application.StartupPath + @"\Canlog\CANLog.exe", fName);
            }*/
        }

        public void CopyLog(TextBox tb)
        {
            //Select all text in the text box.
            tb.Focus();
            tb.SelectAll();
            // Copy the contents of the control to the Clipboard.
            tb.Copy();
        }

        private void Timer_kline_Tick(object sender, EventArgs e)
        {
            // Regularly polling request message
            while (MySerialPort.KLineBlockMessageList.Count() > 0)
            {
                // Pop 1st KLine Block Message
                BlockMessage in_message = MySerialPort.KLineBlockMessageList[0];
                MySerialPort.KLineBlockMessageList.RemoveAt(0);

                // Display debug message on RichTextBox
                String raw_data_in_string = MySerialPort.KLineRawDataInStringList[0];
                MySerialPort.KLineRawDataInStringList.RemoveAt(0);
                DisplayKLineBlockMessage(textBox_serial, "raw_input: " + raw_data_in_string + "\n\r");

                logDumpping.LogCat(kline_text, textBox_serial.Text);
                logDumpping.LogCat(logAll_text, textBox_serial.Text);
                DisplayKLineBlockMessage(textBox_serial, "In - " + in_message.GenerateDebugString() + "\n\r");
                logDumpping.LogCat(kline_text, textBox_serial.Text);
                logDumpping.LogCat(logAll_text, textBox_serial.Text);

                // Process input Kline message and generate output KLine message
                KWP_2000_Process kwp_2000_process = new KWP_2000_Process();
                BlockMessage out_message = new BlockMessage();

                //Use_Random_DTC(kwp_2000_process);  // Random Test
                //Use_Fixed_DTC_from_HQ(kwp_2000_process);  // Simulate response from a ECU device
                //Scan_DTC_from_UI(kwp_2000_process);  // Scan Checkbox status and add DTC into queue
                if (kline_send == 1)
                {
                    foreach (var dtc in ABS_error_list)
                    {
                        kwp_2000_process.ABS_DTC_Queue_Add(dtc);
                    }
                    foreach (var dtc in OBD_error_list)
                    {
                        kwp_2000_process.OBD_DTC_Queue_Add(dtc);
                    }
                }
                else
                {
                    kwp_2000_process.ABS_DTC_Queue_Clear();
                    kwp_2000_process.OBD_DTC_Queue_Clear();
                }


                // Generate output block message according to input message and DTC codes
                kwp_2000_process.ProcessMessage(in_message, ref out_message);

                // Convert output block message to List<byte> so that it can be sent via UART
                List<byte> output_data;
                out_message.GenerateSerialOutput(out output_data);

                // NOTE: because we will also receive all data sent by us, we need to tell UART to skip all data to be sent by SendToSerial
                MySerialPort.Add_ECU_Filtering_Data(output_data);
                MySerialPort.Enable_ECU_Filtering(true);
                // Send output KLine message via UART (after some delay)
                Thread.Sleep((KWP_2000_Process.min_delay_before_response - 1));
                MySerialPort.SendToSerial(output_data.ToArray());

                // Show output KLine message for debug purpose
                DisplayKLineBlockMessage(textBox_serial, "Out - " + out_message.GenerateDebugString() + "\n\r");
                logDumpping.LogCat(kline_text, textBox_serial.Text);
                logDumpping.LogCat(logAll_text, textBox_serial.Text);
            }
        }

        private void DisplayKLineBlockMessage(TextBox rtb, String msg)
        {
            String current_time_str = DateTime.Now.ToString("[HH:mm:ss.fff] ");
            rtb.AppendText(current_time_str + msg + "\n");
            rtb.ScrollToCaret();
        }

        private void button_Start_EnabledChanged(object sender, EventArgs e)
        {
            button_Start.FlatAppearance.BorderColor = Color.FromArgb(242, 242, 242);
            button_Start.FlatAppearance.BorderSize = 3;
            button_Start.BackColor = System.Drawing.Color.FromArgb(242, 242, 242);
        }

        private void button_Pause_EnabledChanged(object sender, EventArgs e)
        {
            button_Pause.FlatAppearance.BorderColor = Color.FromArgb(242, 242, 242);
            button_Pause.FlatAppearance.BorderSize = 3;
            button_Pause.BackColor = System.Drawing.Color.FromArgb(242, 242, 242);
        }

        private void button_Camera_EnabledChanged(object sender, EventArgs e)
        {
            button_Camera.FlatAppearance.BorderColor = Color.FromArgb(242, 242, 242);
            button_Camera.FlatAppearance.BorderSize = 3;
            button_Camera.BackColor = System.Drawing.Color.FromArgb(242, 242, 242);
        }
    }

    public class SafeDataGridView : DataGridView
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            try
            {
                base.OnPaint(e);
            }
            catch
            {
                Invalidate();
            }
        }
    }

    //public class GlobalData{} has been moved to Program.cs//

    /// <summary>
    /// 日期类型转换工具
    /// </summary>
    public class TimestampHelper
    {

        /// <summary>
        /// Unix时间戳转为C#格式时间
        /// </summary>
        /// <param name="timeStamp">Unix时间戳格式,例如:1482115779, 或long类型</param>
        /// <returns>C#格式时间</returns>
        public static DateTime GetDateTime(string timeStamp)
        {
            DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            long lTime = long.Parse(timeStamp + "0000000");
            TimeSpan toNow = new TimeSpan(lTime);
            return dtStart.Add(toNow);
        }

        /// <summary>
        /// 时间戳转为C#格式时间
        /// </summary>
        /// <param name="timeStamp">Unix时间戳格式</param>
        /// <returns>C#格式时间</returns>
        public static DateTime GetDateTime(long timeStamp)
        {
            DateTime time = new DateTime();
            try
            {
                DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
                long lTime = long.Parse(timeStamp + "0000000");
                TimeSpan toNow = new TimeSpan(lTime);
                time = dtStart.Add(toNow);
            }
            catch
            {
                time = DateTime.Now.AddDays(-30);
            }
            return time;
        }

        /// <summary>
        /// DateTime时间格式转换为Unix时间戳格式
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static long ToLong(System.DateTime time)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            return (long)(time - startTime).TotalSeconds;
        }

        /// <summary>
        /// 获取时间戳
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStamp()
        {
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return Convert.ToInt64(ts.TotalSeconds).ToString();
        }

        public static byte[] ConvertToNtp(DateTime datetime)
        {
            ulong milliseconds = (ulong)((datetime - new DateTime(1900, 1, 1)).TotalMilliseconds);

            ulong intpart = 0, fractpart = 0;
            var ntpData = new byte[8];

            intpart = milliseconds / 1000;
            fractpart = ((milliseconds % 1000) * 0x100000000L) / 1000;

            //Debug.WriteLine("intpart:      " + intpart);
            //Debug.WriteLine("fractpart:    " + fractpart);
            //Debug.WriteLine("milliseconds: " + milliseconds);

            var temp = intpart;
            for (var i = 3; i >= 0; i--)
            {
                ntpData[i] = (byte)(temp % 256);
                temp = temp / 256;
            }

            temp = fractpart;
            for (var i = 7; i >= 4; i--)
            {
                ntpData[i] = (byte)(temp % 256);
                temp = temp / 256;
            }
            return ntpData;
        }
    }

    class TimerCustom : System.Timers.Timer
    {
        public Queue<int> queue = new Queue<int>();

        public object lockMe = new object();

        /// <summary>
        /// 为保持连贯性，默认锁住两个
        /// </summary>
        public long lockNum = 0;

        public TimerCustom()
        {
            for (int i = 0; i < short.MaxValue; i++)
            {
                queue.Enqueue(i);
            }
        }
    }

    class Temperature_Data
    {
        public Temperature_Data(double list, bool shot, bool pause, string port, string log, string line)
        {
            temperatureList = list;
            temperatureShot = shot;
            temperaturePause = pause;
            temperaturePort = port;
            temperatureLog = log;
            temperatureNewline = line;
        }

        public static byte temperatureChannel
        {
            get; set;
        }

        public static float addTemperature
        {
            get; set;
        }

        public static float initialTemperature
        {
            get; set;
        }

        public static float finalTemperature
        {
            get; set;
        }

        public static int temperatureDuringtime
        {
            get; set;
        }

        public double temperatureList
        {
            get; set;
        }

        public bool temperatureShot
        {
            get; set;
        }

        public bool temperaturePause
        {
            get; set;
        }

        public string temperaturePort
        {
            get; set;
        }

        public string temperatureLog
        {
            get; set;
        }
        public string temperatureNewline
        {
            get; set;
        }
    }

    /*
        private void setStyle()
        {
            try
            {
                // Form design
                this.MinimumSize = new Size(1097, 659);
                this.BackColor = Color.FromArgb(18, 18, 18);

                //Init material skin
                var skinManager = MaterialSkinManager.Instance;
                skinManager.AddFormToManage(this);
                skinManager.Theme = MaterialSkinManager.Themes.DARK;
                skinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);

                // Button design
                List<Button> buttonsList = new List<Button> { button_Start, button_Setting, button_Pause, button_Schedule, button_Camera, button_SerialPort, button_AcUsb, button_Analysis,
                                                            button_VirtualRC, button_InsertRow, button_SaveSchedule, button_Schedule1, button_Schedule2, button_Schedule3,
                                                            button_Schedule4, button_Schedule5, button_savelog};
                foreach (Button buttonsAll in buttonsList)
                {
                    if (buttonsAll.Enabled == true)
                    {
                        buttonsAll.FlatAppearance.BorderColor = Color.FromArgb(45, 103, 179);
                        buttonsAll.FlatAppearance.BorderSize = 1;
                        buttonsAll.BackColor = System.Drawing.Color.FromArgb(45, 103, 179);
                    }
                    else
                    {
                        buttonsAll.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220);
                        buttonsAll.FlatAppearance.BorderSize = 1;
                        buttonsAll.BackColor = System.Drawing.Color.FromArgb(220, 220, 220);
                    }
                }
            }
            catch (InvalidOperationException)
            {
                //MessageBox.Show(Ex.Message.ToString(), "setStyle Error");
            }
        }
		*/
}
