using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Woodpecker
{
    class Konica_Minolta
    {
        #region Connect CA310/210

        private CA200SRVRLib.Ca200 objCa200;
        private CA200SRVRLib.Ca objCa;
        private CA200SRVRLib.Probe objProbe;
        private CA200SRVRLib.Memory objMemory;
        private CA200SRVRLib.IProbeInfo objProbeInfo;
        private uint isMsr = 0;

        public bool Status()
        {
            bool status = false;
            GlobalData.log.Debug("CA210 status: " + isMsr + "(0: Exception, 1: Disconnect, 2: Connect, 3: CalZero)");
            if (isMsr > 1)
                status = true;
            else
                status = false;

            return status;
        }

        public uint Connect()
        {
            try
            {
                objCa200 = new CA200SRVRLib.Ca200();
                objCa200.AutoConnect();
                objCa = objCa200.SingleCa;
                objProbe = objCa.SingleProbe;
                isMsr = 2;
            }
            catch (Exception)
            {
                isMsr = 0;
            }
			return isMsr;
        }

        public uint DisConnect()
        {
            try
            {
                objCa.RemoteMode = 0;
                objCa200 = null;
                objCa = null;
                objProbe = null;
                objMemory = null;
                objProbeInfo = null;
                isMsr = 1;
            }
            catch (Exception)
            {
                isMsr = 0;
            }
            return isMsr;
        }

        public uint CalZero()
        {
            try
            {
                objCa.CalZero();
                isMsr = 3;
            }
            catch (Exception)
            {
                isMsr = 0;
            }
			return isMsr;
        }

        public string Measure_Once(string measure_remark = "")
        {
            string log_content = "";
            int i = 1, measure_times = 1;
            if (Status() == true)
            {
                try
                {
                    objCa.Measure();
                    DateTime dt = DateTime.Now;
                    string DisplayMode = "";
                    switch (objCa.DisplayMode)
                    {
                        case 0:
                            DisplayMode = "Lvxy";
                            break;
                        case 1:
                            DisplayMode = "Tdudv";
                            break;
                        case 2:
                            DisplayMode = "no display";
                            break;
                        case 3:
                            DisplayMode = "G standard";
                            break;
                        case 4:
                            DisplayMode = "R standard";
                            break;
                        case 5:
                            DisplayMode = "u'v'";
                            break;
                        case 6:
                            DisplayMode = "FMA flicker";
                            break;
                        case 7:
                            DisplayMode = "XYZ";
                            break;
                        case 8:
                            DisplayMode = "JEITA flicker";
                            break;
                    }

                    string log = objProbe.sx.ToString("0.000000") + "," + objProbe.sy.ToString("0.000000") + "," +
                                 objProbe.Lv.ToString("##0.0000") + "," + objProbe.T.ToString("####") + "," +
                                 objProbe.duv.ToString("0.000000") + "," + DisplayMode + "," +
                                 objProbe.X.ToString("##0.0000") + "," + objProbe.Y.ToString("##0.0000") + "," + objProbe.Z.ToString("##0.0000") + "," +
                                 dt.ToString("yyyy/MM/dd") + "," + dt.ToString("HH:mm:ss") + "," +
                                 measure_remark + "," + i + "," + measure_times + "," +
                                 GlobalData.Measure_Backlight + "," + GlobalData.Measure_Thermal + "," + "\r\n";
                    log_content = string.Concat(log_content, log);
                }
                catch (Exception)
                {
                    log_content = "";
					isMsr = 0;
                }
            }
            return log_content;
        }

        public string Measure_Multi(int measure_times = 1, int measure_interval = 0, string measure_remark = "")
        {
            string log_content = "";
            if (Status() == true)
            {
                for (int i = 1; i <= measure_times; i++)
                {
                    try
                    {
                        objCa.Measure();
                        DateTime dt = DateTime.Now;
                        string DisplayMode = "";
                        switch (objCa.DisplayMode)
                        {
                            case 0:
                                DisplayMode = "Lvxy";
                                break;
                            case 1:
                                DisplayMode = "Tdudv";
                                break;
                            case 2:
                                DisplayMode = "no display";
                                break;
                            case 3:
                                DisplayMode = "G standard";
                                break;
                            case 4:
                                DisplayMode = "R standard";
                                break;
                            case 5:
                                DisplayMode = "u'v'";
                                break;
                            case 6:
                                DisplayMode = "FMA flicker";
                                break;
                            case 7:
                                DisplayMode = "XYZ";
                                break;
                            case 8:
                                DisplayMode = "JEITA flicker";
                                break;
                        }

                        string log = objProbe.sx.ToString("0.000000") + "," + objProbe.sy.ToString("0.000000") + "," +
                                     objProbe.Lv.ToString("##0.0000") + "," + objProbe.T.ToString("####") + "," +
                                     objProbe.duv.ToString("0.000000") + "," + DisplayMode + "," +
                                     objProbe.X.ToString("##0.00") + "," + objProbe.Y.ToString("##0.00") + "," + objProbe.Z.ToString("##0.00") + "," +
                                     dt.ToString("yyyy/MM/dd") + "," + dt.ToString("HH:mm:ss") + "," +
                                     measure_remark + "," + i + "," + measure_times + "," + 
                                     GlobalData.Measure_Backlight + "," + GlobalData.Measure_Thermal + "," + "\r\n";
                        log_content = string.Concat(log_content, log);
                    }
                    catch (Exception)
                    {
                        log_content = "";
                        isMsr = 0;
                    }
                    Thread.Sleep(measure_interval);
                }
            }
            return log_content;
        }

        public void DisplayMode(int mode_number)
        {
            if (Status() == true)
            {
                try
                {
                    switch (mode_number)
                    {
                        case 0:
                            objCa.DisplayMode = 0;          // 0. Lvxy.
                            break;
                        case 1:
                            objCa.DisplayMode = 1;          // 1. Tdudv.
                            break;
                        case 2:
                            objCa.DisplayMode = 2;          // 2. no display.
                            break;
                        case 3:
                            objCa.DisplayMode = 3;          // 3. G standard.
                            break;
                        case 4:
                            objCa.DisplayMode = 4;          // 4. R standard.
                            break;
                        case 5:
                            objCa.DisplayMode = 5;          // 5. u'v'.
                            break;
                        case 6:
                            objCa.DisplayMode = 6;          // 6. FMA flicker.
                            break;
                        case 7:
                            objCa.DisplayMode = 7;          // 7. XYZ.
                            break;
                        case 8:
                            objCa.DisplayMode = 8;          // 8. JEITA flicker. 
                            break;
                    }
                }
                catch (Exception)
                {
                    isMsr = 0;
                }
            }
        }
        #endregion
    }
}
