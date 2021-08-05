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
using System.ComponentModel;
using Microsoft.VisualBasic.FileIO;
using USB_VN1630A;
using ModuleLayer;

namespace Woodpecker
{
    public class LogDumpping
    {
        string MainSettingPath = GlobalData.MainSettingPath;
        //static string strPortA = "Port A", strPortB = "Port B", strPortC = "Port C", strPortD = "Port D", strPortE = "Port E", strPort = "Port C", strPortAll = "All", logText = "";

        const int byteMessage_max_Hex = 16;
        const int byteMessage_max_Ascii = 256;
        //private DrvRS232 serialPortA;
        //private Queue<byte> LogQueue_A = new Queue<byte>();

        const int byteTemperatureMax = 64;
        int byteTemperatureLength = 0;
        byte[] byteTemperature = new byte[byteTemperatureMax];
        double previousTemperature = -300;
        List<Temperature_Data> temperatureList = new List<Temperature_Data> { };
        Queue<double> temperatureDouble = new Queue<double> { };

        public void LogDataReceiving(Mod_RS232 serialPort, string portConfig, byte portLF, ref string logText)
        {
            GlobalData.log.Debug("LogDataReceiving start:" + portConfig);
            while (serialPort.IsOpen())
            {
                int data_to_read = serialPort.GetRxBytes();
                if (data_to_read > 0)
                {
                    byte[] dataset = new byte[data_to_read];
                    //GlobalData.LogQueue_A.Enqueue(dataset);

                    serialPort.ReadDataIn(dataset, data_to_read);

                    for (int index = 0; index < data_to_read; index++)
                    {
                        byte input_ch = dataset[index];
                        LogRecording(portConfig, ref logText, input_ch, portLF, true);
                        /*
                        if (TemperatureIsFound == true)
                        {
                            log_temperature(input_ch);
                        }
						*/
                    }
                    //else
                    //    logA_recorder(0x00,true); // tell log_recorder no more data for now.
                }
                //else
                //    logA_recorder(0x00,true); // tell log_recorder no more data for now.
            }
            GlobalData.log.Debug("LogDataReceiving end:" + portConfig);
        }

        //byte[] byteMessage_A = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        //int byteMessage_length_A = 0;
        byte[] byteMessage = new byte[Math.Max(byteMessage_max_Ascii, byteMessage_max_Hex)];
        int byteMessage_length = 0;

        //private void LogRecording(string strPort, string strPortAll, byte ch, bool SaveToLog = false)
        private void LogRecording(string portConfig, ref string logText, byte ch, byte portLF, bool SaveToLog = false)
        {
            if (ini12.INIRead(MainSettingPath, "Record", "Displayhex", "") == "1")
            {
                // if (SaveToLog == false)
                {
                    byteMessage[byteMessage_length] = ch;
                    byteMessage_length++;
                }
                if ((ch == portLF) || (byteMessage_length >= byteMessage_max_Hex))
                {
                    string strData = BitConverter.ToString(byteMessage).Replace("-", "").Substring(0, byteMessage_length * 2);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        strData = "[Receive_" + portConfig + "] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strData + "\r\n"; //OK
                    }

                    LogCat(ref logText, strData);
                    LogCat(ref GlobalData.logAllText, strData);
                    byteMessage_length = 0;
                }
            }
            else
            {
                if ((ch == portLF) || (byteMessage_length >= byteMessage_max_Ascii))
                {
                    string strData = Encoding.ASCII.GetString(byteMessage).Substring(0, byteMessage_length);
                    if (ini12.INIRead(MainSettingPath, "Record", "Timestamp", "") == "1")
                    {
                        DateTime dt = DateTime.Now;
                        strData = "[Receive_" + portConfig + "] [" + dt.ToString("yyyy/MM/dd HH:mm:ss.fff") + "]  " + strData + "\r\n"; //OK
                    }

                    LogCat(ref logText, strData);
                    LogCat(ref GlobalData.logAllText, strData);
                    byteMessage_length = 0;
                }
                else
                {
                    byteMessage[byteMessage_length] = ch;
                    byteMessage_length++;
                }
            }
        }
        
        private void log_temperature(Mod_RS232 serialPort, byte byteTemperatureMax, int byteTemperatureLength, byte ch)
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
            if (byteTemperatureLength >= byteTemperatureMax)
            {
                int destinationIndex = 0;
                for (int i = (byteTemperatureMax - packet_len); i < byteTemperatureMax; i++)
                {
                    byteTemperature[destinationIndex++] = byteTemperature[i];
                }
                byteTemperatureLength = destinationIndex;
            }

            byteTemperature[byteTemperatureLength] = ch;
            byteTemperatureLength++;

            if (ch == 0x0D)
            {
                if (((byteTemperatureLength + header_offset_1) >= 0) &&
                     (byteTemperature[byteTemperatureLength + header_offset_1] == 0x02) &&
                     (byteTemperature[byteTemperatureLength + header_offset_2] == '4'))
                {
                    // Packet is valid here
                    if (byteTemperature[byteTemperatureLength + temp_ch_offset] == Temperature_Data.temperatureChannel)
                    {
                        // Channel number is checked and ok here
                        if ((byteTemperature[byteTemperatureLength + temp_unit_02] == '0'))
                        {
                            if ((byteTemperature[byteTemperatureLength + temp_unit_01] == '1')
                                || (byteTemperature[byteTemperatureLength + temp_unit_01] == '2'))
                            {
                                if ((byteTemperature[byteTemperatureLength + temp_data1_offset] != 0x18))
                                {
                                    // data is valid
                                    int DP_convert = '0';
                                    int byteArray_position = 0;
                                    byte[] byteArray = new byte[8];
                                    for (int pos = byteTemperatureLength + temp_data8_offset;
                                                pos <= (byteTemperatureLength + temp_data1_offset);
                                                pos++)
                                    {
                                        byteArray[byteArray_position] = byteTemperature[pos];
                                        byteArray_position++;
                                    }

                                    string tempSubstring = System.Text.Encoding.Default.GetString(byteArray);
                                    double digit = Math.Pow(10, Convert.ToInt64(byteTemperature[byteTemperatureLength + temp_dp_offset] - DP_convert));
                                    double currentTemperature = Convert.ToDouble(Convert.ToInt32(tempSubstring) / digit);

                                    // is value negative?
                                    if (byteTemperature[byteTemperatureLength + temp_polarity_offset] == '1')
                                    {
                                        currentTemperature = -currentTemperature;
                                    }

                                    // is value Fahrenheit?
                                    if (byteTemperature[byteTemperatureLength + temp_unit_01] == '2')
                                    {
                                        currentTemperature = (currentTemperature - 32) / 1.8;
                                        currentTemperature = Math.Round((currentTemperature), 2, MidpointRounding.AwayFromZero);
                                    }

                                    // check whether 2 temperatures are close enough
                                    if (Math.Abs(previousTemperature - currentTemperature) >= temp_abs_value)
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
                                                //label_Command.Text = "Condition: " + item.temperatureList + ", PAUSE: " + currentTemperature;
                                                //button_Pause.PerformClick();
                                                Console.WriteLine("Temperature: " + currentTemperature + "~~~~~~~~~Temperature matched. Pause the schedule.~~~~~~~~~");
                                            }

                                            if (item.temperatureList == currentTemperature &&
                                                     item.temperaturePort != "" &&
                                                     item.temperatureLog != "" &&
                                                     item.temperatureNewline != "")
                                            {
                                                //label_Command.Text = "Condition: " + item.temperatureList + ", Log: " + currentTemperature;
                                                if (item.temperatureLog.Contains('|'))
                                                {
                                                    string[] logArray = item.temperatureLog.Split('|');
                                                    for (int i = 0; i < logArray.Length; i++)
                                                        ReplaceNewLine(serialPort, logArray[i], item.temperatureNewline);
                                                }
                                                else
                                                {
                                                    ReplaceNewLine(serialPort, item.temperatureLog, item.temperatureNewline);
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

                byteTemperatureLength = 0;
            }
        }
        

        // Log record function
        public void LogCat(ref string logComponent, string log)
        {
            //PortA, PortB, PortC, PortD, PortE, CA310, Canbus, KlinePort, All
            logComponent = string.Concat(logComponent, log);
        }

        public void LogCat(string logComponent, string log)
        {
            try
            {
                //PortA, PortB, PortC, PortD, PortE, CA310, Canbus, KlinePort, All
                logComponent = string.Concat(logComponent, log);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message.ToString());
                //Serialportsave("All");
            }
        }

        public string ReturnLogCat(string logComponent, string log)
        {
            string allLog;
            allLog = string.Concat(logComponent, log);
            return allLog;
        }

        #region -- 換行符號置換 --
        //private void ReplaceNewLine(SerialPort port, string columns_serial, string columns_switch)
        public void ReplaceNewLine(Mod_RS232 serialPort, string columns_serial, string columns_switch)
        {
            List<string> originLineList = new List<string> { "\\r\\n", "\\n\\r", "\\r", "\\n" };
            List<string> newLineList = new List<string> { "\r\n", "\n\r", "\r", "\n" };
            int dataLength = columns_serial.Length + columns_switch.Length;
            var originAndNewLine = originLineList.Zip(newLineList, (o, n) => new { origin = o, newLine = n });
            foreach (var line in originAndNewLine)
            {
                if (columns_switch.Contains(line.origin))
                {
                    serialPort.WriteDataOut(columns_serial + columns_switch.Replace(line.origin, line.newLine), dataLength);    //Send RS232 data
                    return;
                }
            }
        }
        #endregion

        public void LogDumpToFile(string portConfig, string portName, ref string logText)
        {
            GlobalData.log.Debug("LogDumpToFile start:" + portName);
            string fName = "";
            // 讀取ini中的路徑
            fName = ini12.INIRead(GlobalData.MainSettingPath, "Record", "LogPath", "");
            Console.WriteLine("[YFC]" + fName);
            string t = fName + "\\_" + portConfig + "[" + portName + "]_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + GlobalData.Loop_Number.ToString() + ".txt";
            StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
            MYFILE.Write(logText);
            MYFILE.Close();
            logText = String.Empty;
            /*
            switch (portConfig)
            {
                case "PortA":
                    //string t = fName + "\\_PortA_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    //string t = fName + "\\_" + logConfig_A.portName + "[" + logConfig_A.portNumber + "]_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + ".txt";
                    StreamWriter MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logText);
                    //MYFILE.Write(logA_text);
                    MYFILE.Close();
                    //Txtbox1("", textBox_serial);
                    logText = String.Empty;
                    break;
                case "PortB":
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logText);
                    MYFILE.Close();
                    logText = String.Empty;
                    break;
                case "CA310":
                    t = fName + "\\_CA310_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(ca310_text);
                    MYFILE.Close();
                    ca310_text = String.Empty;
                    break;
                case "Canbus":
                    t = fName + "\\_Canbus_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    //MYFILE.Write(canbus_text);
                    MYFILE.Close();
                    //canbus_text = String.Empty;
                    break;
                case "KlinePort":
                    //t = fName + "\\_Kline_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    t = fName + "\\_Kline_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(kline_text);
                    MYFILE.Close();
                    kline_text = String.Empty;
                    break;
                case "All":
                    //t = fName + "\\_AllPort_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    t = fName + "\\_AllPort_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(logAll_text);
                    MYFILE.Close();
                    logAll_text = String.Empty;
                    break;
                case "Debug":
                    //t = fName + "\\_Debug_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + label_LoopNumber_Value.Text + ".txt";
                    t = fName + "\\_Debug_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + ".txt";
                    MYFILE = new StreamWriter(t, false, Encoding.ASCII);
                    MYFILE.Write(debug_text);
                    MYFILE.Close();
                    debug_text = String.Empty;
                    break;
            }*/
            GlobalData.log.Debug("LogDumpToFile end:" + portName);
        }
    }

    public class RK2797
    {
        // copy to mei
        enum command_index
        {
            // Calibration Command
            SET_GAMMA_INDEX = 0,
            GET_GAMMA_INDEX,
            //SET_OUTPUT_GAMMA_TABLE_INDEX,
            GET_OUTPUT_GAMMA_TABLE_INDEX,
            SET_COLOR_GAMUT_INDEX,
            GET_COLOR_GAMUT_INDEX,
            //SET_INPUT_GAMMA_TABLE_INDEX,
            GET_INPUT_GAMMA_TABLE_INDEX,
            //SET_PCM_MARTIX_TABLE_INDEX,
            GET_PCM_MARTIX_TABLE_INDEX,
            SET_COLOR_TEMP_INDEX,
            GET_COLOR_TEMP_INDEX,
            SET_RGB_GAIN_INDEX,
            GET_RGB_GAIN_INDEX,

            // Control Command
            SET_BACKLIGHT_INDEX,
            GET_BACKLIGHT_INDEX,
            SET_PQ_ONOFF_INDEX,
            SET_INTERNAL_PATTERN_INDEX,
            SET_PATTERN_RGB_INDEX,
            SET_SHARPNESS_INDEX,
            GET_SHARPNESS_INDEX,
            GET_BACKLIGHT_SENSOR_INDEX,
            GET_THERMAL_SENSOR_INDEX,
            SET_SPI_PORT_INDEX,
            GET_SPI_PORT_INDEX,
            SET_UART_PORT_INDEX,
            GET_UART_PORT_INDEX,
            SET_BRIGHTNESS_INDEX,
            GET_BRIGHTNESS_INDEX,
            SET_CONTRAST_INDEX,
            GET_CONTRAST_INDEX,
            SET_MAIN_INPUT_INDEX,
            GET_MAIN_INPUT_INDEX,
            SET_SUB_INPUT_INDEX,
            GET_SUB_INPUT_INDEX,
            SET_PIP_MODE_INDEX,
            GET_PIP_MODE_INDEX,

            // Write Data Command
            GET_SCALER_TYPE_INDEX,
            //SET_MODEL_NAME_INDEX,
            GET_MODEL_NAME_INDEX,
            //SET_EDID_INDEX,
            GET_EDID_INDEX,
            //SET_HDCP14_INDEX,
            GET_HDCP14_INDEX,
            //SET_HDCP2x_INDEX,
            GET_HDCP2x_INDEX,
            //SET_SERIAL_NUMBER_INDEX,
            GET_SERIAL_NUMBER_INDEX,
            GET_FW_VERSION_INDEX,
            GET_FAC_EEPROM_DATA_INDEX,

            // BenQ Command
            //SET_BENQ_MODEL_NAME_INDEX,
            //SET_BENQ_SERIAL_NAME_INDEX,
            //SET_BENQ_FW_VERSION_INDEX,
            //SET_BENQ_MONITOR_ID_INDEX,
            //SET_BENQ_DNA_VERSION_INDEX,
            //SET_BENQ_MANUFACTURE_YEARANDDATE_INDEX,
            //SET_BENQ_EEPROM_INIT_INDEX,
            //GET_BENQ_EEPROM_INDEX,
        }

        byte[][] Command_Packet =
        {
			// Calibration Command
            new byte[] { 0x06, 0x00, 0xE0, 0x00, 0xff, 0xff },              ///SET_GAMMA_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x01, 0xff },                    ///GET_GAMMA_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x02, 0xff, 0xff, 0xff },      ///SET_OUTPUT_GAMMA_TABLE_INDEX,
            new byte[] { 0x07, 0x00, 0xE0, 0x03, 0xff, 0xff, 0xff },        ///GET_OUTPUT_GAMMA_TABLE_INDEX,
            new byte[] { 0x06, 0x00, 0xE0, 0x04, 0xff, 0xff },              ///SET_COLOR_GAMUT_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x05, 0xff },                    ///GET_COLOR_GAMUT_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x06, 0xff, 0xff, 0xff, 0xff, 0xff },						///SET_INPUT_GAMMA_TABLE_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x07, 0xff },              		///GET_INPUT_GAMMA_TABLE_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x08, 0xff, 0xff, 0xff },      ///SET_PCM_MARTIX_TABLE_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x09, 0xff },                    ///GET_PCM_MARTIX_TABLE_INDEX,
            new byte[] { 0x06, 0x00, 0xE0, 0x0A, 0xff, 0xff },              ///SET_COLOR_TEMP_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x0B, 0xff },              		///GET_COLOR_TEMP_INDEX,
            new byte[] { 0x07, 0x00, 0xE0, 0x0C, 0xff, 0xff, 0xff },		///SET_RGB_GAIN_INDEX,
            new byte[] { 0x05, 0x00, 0xE0, 0x0D, 0xff },              		///GET_RGB_GAIN_INDEX,		
            
            // Control Command
            new byte[] { 0x06, 0x01, 0xE0, 0x00, 0xff, 0xff },		    	///SET_BACKLIGHT_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x01, 0xff },           			///GET_BACKLIGHT_INDEX,			
            new byte[] { 0x07, 0x01, 0xE0, 0x02, 0xff, 0xff, 0xff },        ///SET_PQ_ONOFF_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x03, 0xff, 0xff },              ///SET_INTERNAL_PATTERN_INDEX,
            new byte[] { 0x0B, 0x01, 0xE0, 0x04, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff },            ///SET_PATTERN_RGB_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x05, 0xff, 0xff },              ///SET_SHARPNESS_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x06, 0xff },              		///GET_SHARPNESS_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x07, 0xff },              		///GET_BACKLIGHT_SENSOR_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x08, 0xff },                    ///GET_THERMAL_SENSOR_INDEX,
            //new byte[] { 0x06, 0x01, 0xE0, 0x08, 0xff, 0xff },            ///GET_THERMAL_SENSOR_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x09, 0xff, 0xff },              ///SET_SPI_PORT_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x0A, 0xff },              		///GET_SPI_PORT_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x0B, 0xff, 0xff },              ///SET_UART_PORT_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x0C, 0xff },              		///GET_UART_PORT_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x0D, 0xff, 0xff },              ///SET_BRIGHTNESS_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x0E, 0xff },        		    ///GET_BRIGHTNESS_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x0F, 0xff, 0xff },              ///SET_CONTRAST_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x10, 0xff },              		///GET_CONTRAST_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x11, 0xff, 0xff },              ///SET_MAIN_INPUT_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x12, 0xff },              		///GET_MAIN_INPUT_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x13, 0xff, 0xff },              ///SET_SUB_INPUT_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x14, 0xff },              		///GET_SUB_INPUT_INDEX,
            new byte[] { 0x06, 0x01, 0xE0, 0x15, 0xff, 0xff },              ///SET_PIP_MODE_INDEX,
            new byte[] { 0x05, 0x01, 0xE0, 0x16, 0xff },              		///GET_PIP_MODE_INDEX,

            // Write Data Command
            new byte[] { 0x05, 0x02, 0xE0, 0x00, 0xff },              		///GET_SCALER_TYPE_INDEX,
            //new byte[] { 0xff, 0x02, 0xE0, 0x01, 0xff, 0xff, 0xff },      ///SET_MODEL_NAME_INDEX,
            new byte[] { 0x05, 0x02, 0xE0, 0x02, 0xff },              		///GET_MODEL_NAME_INDEX,
            //new byte[] { 0xff, 0x02, 0xE0, 0x03, 0xff, 0xff, 0xff, 0xff, 0xff },              ///SET_EDID_INDEX,
            new byte[] { 0x06, 0x02, 0xE0, 0x04, 0xff, 0xff },              ///GET_EDID_INDEX,
            //new byte[] { 0xff, 0x02, 0xE0, 0x05, 0xff, 0xff, 0xff, 0xff },///SET_HDCP14_INDEX,
            new byte[] { 0x05, 0x02, 0xE0, 0x06, 0xff },              		///GET_HDCP14_INDEX,
            //new byte[] { 0xff, 0x02, 0xE0, 0x07, 0xff, 0xff, 0xff, 0xff },///SET_HDCP2x_INDEX,
            new byte[] { 0x06, 0x02, 0xE0, 0x08, 0xff },              		///GET_HDCP2x_INDEX,
            //new byte[] { 0xff, 0x02, 0xE0, 0x09, 0xff, 0xff, 0xff },      ///SET_SERIAL_NUMBER_INDEX,
            new byte[] { 0x05, 0x02, 0xE0, 0x0A, 0xff },              		///GET_SERIAL_NUMBER_INDEX,
            new byte[] { 0x05, 0x02, 0xE0, 0x0B, 0xff },              		///GET_FW_VERSION_INDEX,
            new byte[] { 0x05, 0x02, 0xE0, 0x0C, 0xff },              		///GET_FAC_EEPROM_DATA_INDEX,

            // BenQ Command
            //new byte[] { 0xff, 0x00, 0xE0, 0x00, 0xff, 0xff, 0xff },      ///SET_BENQ_MODEL_NAME_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x01, 0xff, 0xff, 0xff },      ///SET_BENQ_SERIAL_NAME_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x02, 0xff, 0xff, 0xff },      ///SET_BENQ_FW_VERSION_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x03, 0xff, 0xff, 0xff },      ///SET_BENQ_MONITOR_ID_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x04, 0xff, 0xff, 0xff },      ///SET_BENQ_DNA_VERSION_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x05, 0xff, 0xff, 0xff },      ///SET_BENQ_MANUFACTURE_YEARANDDATE_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x06, 0xff, 0xff, 0xff },      ///SET_BENQ_EEPROM_INIT_INDEX,
            //new byte[] { 0xff, 0x00, 0xE0, 0x07, 0xff, 0xff, 0xff },      ///GET_BENQ_EEPROM_INDEX,
        };
        // copy to mei

        byte[][] Parsing_Packet =
        {
            //// Calibration Command
            new byte [] { }, //SET_GAMMA_INDEX = 0,
            new byte [] { 0x06, 0x00, 0xE0, 0x01 }, //GET_GAMMA_INDEX,
            //new byte [] { }, //SET_OUTPUT_GAMMA_TABLE_INDEX,
            new byte [] { }, //GET_OUTPUT_GAMMA_TABLE_INDEX,
            new byte [] { }, //SET_COLOR_GAMUT_INDEX,
            new byte [] { 0x06, 0x00, 0xE0, 0x05 }, //GET_COLOR_GAMUT_INDEX,
            //new byte [] { }, //SET_INPUT_GAMMA_TABLE_INDEX,
            new byte [] { }, //GET_INPUT_GAMMA_TABLE_INDEX,
            //new byte [] { }, //SET_PCM_MARTIX_TABLE_INDEX,
            new byte [] { }, //GET_PCM_MARTIX_TABLE_INDEX,
            new byte [] { }, //SET_COLOR_TEMP_INDEX,
            new byte [] { 0x06, 0x00, 0xE0, 0x0B }, //GET_COLOR_TEMP_INDEX,
            new byte [] { }, //SET_RGB_GAIN_INDEX,
            new byte [] { 0x08, 0x00, 0xE0, 0x0D }, //GET_RGB_GAIN_INDEX,
            
            //// Control Command
            new byte [] { }, //SET_BACKLIGHT_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x01 },  //GET_BACKLIGHT_INDEX,
            new byte [] { }, //SET_PQ_ONOFF_INDEX,
            new byte [] { }, //SET_INTERNAL_PATTERN_INDEX,
            new byte [] { }, //SET_PATTERN_RGB_INDEX,
            new byte [] { }, //SET_SHARPNESS_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x06 }, //GET_SHARPNESS_INDEX,
            new byte [] { 0x07, 0x01, 0xE0, 0x07 }, //GET_BACKLIGHT_SENSOR_INDEX,
            new byte [] { 0x09, 0x01, 0xE0, 0x08 }, //GET_THERMAL_SENSOR_INDEX,
            new byte [] { }, //SET_SPI_PORT_INDEX,
            new byte [] { 0x07, 0x01, 0xE0, 0x0A }, //GET_SPI_PORT_INDEX,
            new byte [] { }, //SET_UART_PORT_INDEX,
            new byte [] { 0x07, 0x01, 0xE0, 0x0C }, //GET_UART_PORT_INDEX,
            new byte [] { }, //SET_BRIGHTNESS_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x0E }, //GET_BRIGHTNESS_INDEX,
            new byte [] { }, //SET_CONTRAST_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x10 }, //GET_CONTRAST_INDEX,
            new byte [] { }, //SET_MAIN_INPUT_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x12 }, //GET_MAIN_INPUT_INDEX,
            new byte [] { }, //SET_SUB_INPUT_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x14 }, //GET_SUB_INPUT_INDEX,
            new byte [] { }, //SET_PIP_MODE_INDEX,
            new byte [] { 0x06, 0x01, 0xE0, 0x16 }, //GET_PIP_MODE_INDEX,
            
            //new byte [] { }, // Write Data Command
            new byte [] { }, //GET_SCALER_TYPE_INDEX,
            //new byte [] { }, //SET_MODEL_NAME_INDEX,
            new byte [] { }, //GET_MODEL_NAME_INDEX,
            //new byte [] { }, //SET_EDID_INDEX,
            new byte [] { }, //GET_EDID_INDEX,
            //new byte [] { }, //SET_HDCP14_INDEX,
            new byte [] { }, //GET_HDCP14_INDEX,
            //new byte [] { }, //SET_HDCP2x_INDEX,
            new byte [] { }, //GET_HDCP2x_INDEX,
            //new byte [] { }, //SET_SERIAL_NUMBER_INDEX,
            new byte [] { }, //GET_SERIAL_NUMBER_INDEX,
            new byte [] { }, //GET_FW_VERSION_INDEX,
            new byte [] { }, //GET_FAC_EEPROM_DATA_INDEX,
            
            //// BenQ Command
            //new byte [] { }, //SET_BENQ_MODEL_NAME_INDEX,
            //new byte [] { }, //SET_BENQ_SERIAL_NAME_INDEX,
            //new byte [] { }, //SET_BENQ_FW_VERSION_INDEX,
            //new byte [] { }, //SET_BENQ_MONITOR_ID_INDEX,
            //new byte [] { }, //SET_BENQ_DNA_VERSION_INDEX,
            //new byte [] { }, //SET_BENQ_MANUFACTURE_YEARANDDATE_INDEX,
            //new byte [] { }, //SET_BENQ_EEPROM_INIT_INDEX,
            //new byte [] { }, //GET_BENQ_EEPROM_INDEX,
       };

        //  RK2797_cmd_list
        public void Package_add_queue(Mod_RS232 serialPort)
        {
            List<byte> BackupDataList = new List<byte>();

            while (serialPort.ReceiveQueue.Count > 0)                   // Queue有資料就收取
            {
                //  Queue一個byte一個byte取出來被丟入List
                byte serial_byte = (byte)serialPort.ReceiveQueue.Dequeue();
                serialPort.ReceiveList.Add(serial_byte);                    // Queue一個byte一個byte取出來被丟入List
                //  BackupDataList.Add(serial_byte);                          // Queue debug list content
            }
        }

        public void Package_queue_to_list(Mod_RS232 serialPort)
        {
            Algorithm algorithm = new Algorithm();
            if (serialPort.ReceiveList.Count >= 3)
            {
                if (serialPort.ReceiveList.ElementAt(2) != 0xE0)
                {
                    serialPort.ReceiveList.RemoveAt(0);
                }
                else
                {
                    byte packet_len = serialPort.ReceiveList.ElementAt(0);
                    if (packet_len >= 4)
                    {
                        if (serialPort.ReceiveList.Count >= packet_len)
                        {
                            byte calculate_checksum;
                            calculate_checksum = algorithm.XOR_List(serialPort.ReceiveList, packet_len);

                            if (calculate_checksum == 0)
                            {
                                List<byte> CurrentDataList = new List<byte>();
                                CurrentDataList = serialPort.ReceiveList.GetRange(0, packet_len);
                                serialPort.ReceiveQueueList.Enqueue(CurrentDataList);                  // Enqueue list byte data
                                serialPort.ReceiveList.RemoveRange(0, packet_len);
                            }
                            else
                            {
                                serialPort.ReceiveList.RemoveAt(0);
                            }
                        }
                    }
                    else
                    {
                        serialPort.ReceiveList.RemoveAt(0);
                    }
                }
            }
        }


        public bool Package_queue_to_catch(Mod_RS232 serialPort)
        {
            bool status = false;
            List<byte> CurrentDataList = new List<byte>();
            int BACKLIGHT_SENSOR_INDEX = (int)command_index.GET_BACKLIGHT_SENSOR_INDEX;
            int THERMAL_SENSOR_INDEX = (int)command_index.GET_THERMAL_SENSOR_INDEX;

            if (serialPort.ReceiveQueueList.Count > 0)
            {
                CurrentDataList = serialPort.ReceiveQueueList.Dequeue();
                // update parsing here
                if (Parse_packet(CurrentDataList, Parsing_Packet[BACKLIGHT_SENSOR_INDEX].ToList()) == true)
                {
                    GlobalData.Measure_Backlight = raw_data(CurrentDataList);
                }
                else if (Parse_packet(CurrentDataList, Parsing_Packet[THERMAL_SENSOR_INDEX].ToList()) == true)
                {
                    GlobalData.Measure_Thermal = raw_data(CurrentDataList);
                }
            }
            return status;
        }

        private bool Parse_packet(List<byte> input_packet, List<byte> original_packet)
        {
            bool ret_value = true;

            for (int index = 0; index < original_packet.Count; index++)
            {
                if (input_packet.ElementAt(index) != original_packet[index])
                {
                    ret_value = false;
                    break;
                }
            }
            return ret_value;
        }

        public string GetDUTSensor(string measure_remark = "")
        {
            string log_content = "";
            int i = 1, measure_times = 1;
            try
            {
                DateTime dt = DateTime.Now;
                string DisplayMode = "None";
                string log = "None" + "," + "None" + "," +
                             "None" + "," + "None" + "," +
                             "None" + "," + DisplayMode + "," +
                             "None" + "," + "None" + "," + "None" + "," +
                             dt.ToString("yyyy/MM/dd") + "," + dt.ToString("HH:mm:ss") + "," +
                             measure_remark + "," + i + "," + measure_times + "," +
                             GlobalData.Measure_Backlight + "," + GlobalData.Measure_Thermal + "," + "\r\n";
                log_content = string.Concat(log_content, log);
            }
            catch (Exception)
            {
                log_content = "";
            }
            return log_content;
        }

        private string raw_data(List<byte> data)
        {
            string HexString = "";
            if (data != null)
            {
                foreach (byte sum in data)
                {
                    HexString += (sum.ToString("X2"));
                }
            }
            return HexString;
        }
    }
}


/*
private void initComboboxSaveLog()
{
    List<string> portList = new List<string> { "Port A", "Port B", "Port C", "Port D", "Port E", "Kline", "Canbus" };

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

    if (ini12.INIRead(MainSettingPath, "Device", "CA310Exist", "") == "1")
        comboBox_savelog.Items.Add("CA310");

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
*/


// OPTT debug function 
/*
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
*/



/*
private void Log(string msg)
{
    textBox_serial.Invoke(new EventHandler(delegate
    {
        textBox_serial.Text = msg.Trim();
        PortA.WriteLine(msg.Trim());
        //PortA.We
        //myPortA .WriteLine(msg.Trim());
    }));
}


/*
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
#endregion
*/



/*
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
#endregion

/*
private void button_savelog_Click(object sender, EventArgs e)
{
    string save_option = comboBox_savelog.Text;
    switch (save_option)
    {
        case "Port A":
            Serialportsave("A");
            MessageBox.Show("Port A is saved.", "Reminder");
            break;
        case "Port B":
            Serialportsave("B");
            MessageBox.Show("Port B is saved.", "Reminder");
            break;
        case "Port C":
            Serialportsave("C");
            MessageBox.Show("Port C is saved.", "Reminder");
            break;
        case "Port D":
            Serialportsave("D");
            MessageBox.Show("Port D is saved.", "Reminder");
            break;
        case "Port E":
            Serialportsave("E");
            MessageBox.Show("Port E is saved.", "Reminder");
            break;
        case "CA310":
            Serialportsave("CA310");
            MessageBox.Show("CA310 is saved.", "Reminder");
            break;
        case "Canbus":
            Serialportsave("Canbus");
            MessageBox.Show("Canbus is saved.", "Reminder");
            break;
        case "Kline":
            Serialportsave("KlinePort");
            MessageBox.Show("Kline Port is saved.", "Reminder");
            break;
        case "Port All":
            Serialportsave("All");
            MessageBox.Show("All Port is saved.", "Reminder");
            break;
        default:
            break;
    }
}
*/

/*
private void logB_analysis()
{
    while (PortB.IsOpen == true)
    {
        int data_to_read = PortB.BytesToRead;
        if (data_to_read > 0)
        {
            byte[] dataset = new byte[data_to_read];
            PortB.Read(dataset, 0, data_to_read);

            for (int index = 0; index < data_to_read; index++)
            {
                byte input_ch = dataset[index];
                logB_recorder(input_ch);
                if (TemperatureIsFound == true)
                {
                    log_temperature(input_ch);
                }
            }
        }
        //else
        //{
        //    logB_recorder(0x00,true); // tell log_recorder no more data for now.
        //}
    }
}
*/
/*
private void log_process(string port, string log)
{
    try
    {
        switch (port)
        {
            case "A":
                logA_text = string.Concat(logA_text, log);
                break;
            case "B":
                logB_text = string.Concat(logB_text, log);
                break;
            case "C":
                logC_text = string.Concat(logC_text, log);
                break;
            case "D":
                logD_text = string.Concat(logD_text, log);
                break;
            case "E":
                logE_text = string.Concat(logE_text, log);
                break;
            case "CA310":
                ca310_text = string.Concat(ca310_text, log);
                break;
            case "Canbus":
                canbus_text = string.Concat(canbus_text, log);
                break;
            case "KlinePort":
                kline_text = string.Concat(kline_text, log);
                break;
            case "All":
                logAll_text = string.Concat(logAll_text, log);
                break;
        }
    }
    catch (Exception ex)
    {
        Console.Write(ex.Message.ToString());
        Serialportsave("All");
    }
}
*/
