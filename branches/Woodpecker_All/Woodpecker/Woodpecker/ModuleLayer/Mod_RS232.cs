using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using Microsoft.Win32.SafeHandles;           //support SafeFileHandle
using System.Runtime.InteropServices;      //support DIIImport
using System.Reflection;                                //support BindingFlags
using jini;
using Woodpecker;
using System.Collections;

namespace ModuleLayer
{
    public class Mod_RS232
    {
        //private int iPortNmuber = 0;
        private SerialPort _serialPort = new SerialPort();
        private Stream _internalSerialStream;
        public Queue<byte> ReceiveQueue = new Queue<byte>();
        public List<byte> ReceiveList = new List<byte>();
        public Queue<List<byte>> ReceiveQueueList = new Queue<List<byte>>();
        /*
        private SerialPortConfig portConfig;

        public class SerialPortConfig
        {   // pastebin.com/KmKEVzR8 //
            public string Name { get; private set; }
            public int BaudRate { get; private set; }
            public int DataBits { get; private set; }
            public StopBits StopBits { get; private set; }
            public Parity Parity { get; private set; }
            public bool DtrEnable { get; private set; }
            public bool RtsEnable { get; private set; }

            public SerialPortConfig(
                string name,
                int baudRate,
                int dataBits,
                StopBits stopBits,
                Parity parity,
                bool dtrEnable,
                bool rtsEnable)
            {
                if (String.IsNullOrWhiteSpace(name)) throw new ArgumentNullException("name");
                this.Name = name;
                this.BaudRate = baudRate;
                this.DataBits = dataBits;
                this.StopBits = stopBits;
                this.Parity = parity;
                this.DtrEnable = dtrEnable;
                this.RtsEnable = rtsEnable;
            }
        }
        */
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes,
                                                            uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private SafeFileHandle handle_Com = null;
        //private static object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_ComPort);
        //private static SafeFileHandle handle_Com = (SafeFileHandle)stream.GetType().GetField("_handle", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(stream);

        //Read data Into buffer
        public int ReadTerm(Byte[] ResultDataBuf, ref int Count, char TermByte)
        {
            int DataLen = _serialPort.BytesToRead;
            Count = 0;

            if (DataLen >= 1)
            {
                for (int i = 0; i <= (DataLen - 1); i++)
                {
                    
                    ResultDataBuf[i] = (Byte)_serialPort.ReadByte();
                    Count++;

                    if (ResultDataBuf[i] == TermByte)
                    {
                        Array.Resize(ref ResultDataBuf, (i+1));
                        return 1;
                    }
                }

                return -1;
            }
            return -1;

        }

        public int ReadDataIn(Byte[] InBuf, int Length)
        {
            try
            {
                _serialPort.Read(InBuf, 0, Length);
                for (int i = 0; i < Length; i++)
                    ReceiveQueue.Enqueue(InBuf[i]);
            }
            catch (System.TimeoutException)
            {//Time out
                return -1;
            }
            catch (System.ArgumentException)
            {
                return -1;
            }
            catch (System.IO.IOException)
            {//Port number error
                return -1;
            }

            return 1;
        }

        public int GetRxBytes()
        {
            return (_serialPort.BytesToRead);
        }

        //Write data out
        public int WriteDataOut(string InBuf, int DataLength)
        {
            if (_serialPort.IsOpen == true)
            {
                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(InBuf);

                }
                catch (System.ArgumentException)
                {
                    return -1;
                }
            }
            else
            {
                return -1;
            }
            return -1;
        }

        public int WriteDataOut(char[] InBuf, int DataLength)
        {

            if (_serialPort.IsOpen == true)
            {
                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(InBuf, 0, DataLength);

                }
                catch (System.ArgumentException)
                {
                    return -1;
                }
            }
            else
            {
                return -1;
            }
            return -1;
        }

        public int WriteDataOut(byte[] InBuf, int DataLength)
        {
            if (_serialPort.IsOpen == true)
            {
                try
                {
                    _serialPort.DiscardInBuffer();
                    _serialPort.DiscardOutBuffer();
                    _serialPort.Write(InBuf, 0, DataLength);
                }
                catch (System.ArgumentException)
                {
                    return -1;
                }
            }
            else
                return -1;
            
            return -1;
        }

        public int GetDataFromQueue(int Len, ref byte[] retBuf)
        {
            int i, j;
            if ((ReceiveQueue.Count <= 0) || (retBuf.Length < Len))
            {
                return -1;
            }
            else
            {
                try
                {
                    if (Len > ReceiveQueue.Count)
                    {
                        j = ReceiveQueue.Count;
                    }
                    else
                    {
                        j = Len;
                    }
                    Console.Write("\nInBuf:");

                    for (i = 0; i <= (j - 1); i++)
                    {
                        retBuf[i] = (byte)ReceiveQueue.Dequeue();
                        Console.Write("{0,2:X},", retBuf[i]);
                    }
                    Console.Write("\n");
                }
                catch (Exception)
                {
                    return 1;
                }

            }
            return 1;
        }

        public int ReceivedBufferLength()
        {
            return (ReceiveQueue.Count);
        }

        public void DataReceivedByEvent(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;   //this is used with Forms-Serialport component inserted
            try
            {
                if (sp.IsOpen)
                {
                    byte[] byteRead = new byte[sp.BytesToRead];
                    //it must use ReadLine() to do carriage-return and line-feed instead of using /r/n.
                    string dataValue = sp.ReadLine();
                    if (dataValue.Contains("io i"))
                        GlobalData.Arduino_Read_String = dataValue;
                    GlobalData.log.Debug("[" + sp + "] DataReceived: " + dataValue);

                    int recByteCount = byteRead.Count();
                    if (recByteCount > 0)
                        GlobalData.Arduino_recFlag = true;
                    else
                        GlobalData.Arduino_recFlag = false;

                    sp.DiscardInBuffer();                      //Clear buffer of SerialPort component
                }
                else
                {
                    Console.WriteLine("Arduino serialport is not opened!");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message, "DataReceivedEvent error!");
            }
        }

        //public int OpenSerialPort(int PortNumber, int BaudRate, int ParityBit, int DataLength, int StopBit, int Handshake)
        public int OpenSerialPort(string port_Name, string port_BR, bool receive_Data = false)
        {
            try
            {
                if (_serialPort.IsOpen == false)
                {
                    _serialPort.PortName = port_Name;
                    _serialPort.BaudRate = int.Parse(port_BR);
                    _serialPort.DataBits = 8;
                    _serialPort.StopBits = StopBits.One;
                    /*
                    string stopbits = ini12.INIRead(GlobalData.MainSettingPath, "Port A", "StopBits", "");
                    switch (stopbits)
                    {
                        case "One":
                            _ComPort.StopBits = StopBits.One;
                            break;
                        case "Two":
                            _ComPort.StopBits = StopBits.Two;
                            break;
                    }
                    */
                    _serialPort.Handshake = Handshake.None;
                    _serialPort.Parity = Parity.None;
                    _serialPort.ReadTimeout = 2000;
                    _serialPort.WriteTimeout = 100;
                    if (receive_Data)
                        _serialPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedByEvent);
                    _serialPort.Open();

                    Console.WriteLine("[DrvRS232] " + _serialPort.PortName + " is successfully opened.");

                    _internalSerialStream = _serialPort.BaseStream;
                }
                else
                {
                    Console.WriteLine("[Mod_RS232] " + _serialPort.PortName + " is not yet opened.");
                    return -3;
                }
            }
            catch (System.IO.IOException)
            {   //Port number error
                return -1;
            }
            catch (System.UnauthorizedAccessException)
            {   //Port is used by another application
                return -2;
            }
            catch (Exception e)
            {
                Console.WriteLine("[Drv_RS232]" + e.Message);
            }

            return 1;
        }

        public int ClosePort()
        {
            /*
            if (_internalSerialStream == null)
            {
                object internalStream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_serialPort);
                //handle_Com = (SafeFileHandle)internalStream.GetType().GetField("_handle", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(internalStream);
            }

            _internalSerialStream = _serialPort.BaseStream;
            _internalSerialStream.Close();
            GC.SuppressFinalize(_serialPort);
            GC.SuppressFinalize(_internalSerialStream);
            */
            _serialPort.Dispose();
            _serialPort.Close();
            /* In case there happened an error by unplugging USB wire during serial data read/write,
             *   one can use below way to release a handle which is used to handle the serial port.
             */
            //object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_ComPort);
            //handle_Com = (SafeFileHandle)stream.GetType().GetField("_handle", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(stream);


            return 1;
        }

        public bool IsOpen()
        {
            if (_serialPort.IsOpen)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public SafeFileHandle Handle
        {
            get
            {
                object stream = typeof(SerialPort).GetField("internalSerialStream", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_serialPort);
                // If the handle is valid, return it.
                if (!handle_Com.IsInvalid)
                {
                   
                    handle_Com = (SafeFileHandle)stream.GetType().GetField("_handle", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(stream);
                    return handle_Com;
                }
                else
                    return null;
            }
        }

    }
}
