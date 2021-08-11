using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Xml;
using System.Xml.Linq;
using System.Windows;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Universal_Toolkit.Types;
using Woodpecker;

namespace Universal_Toolkit
{
    public class FTDI_Lib
    {
        internal static int _initializations = 0;

        /// <summary>
        /// LibMPSSE.dll Import 
        /// </summary>
        /// <param name="numChannels"></param>
        /// <returns></returns>
        /// 

        

        ////DLL Import I2C Function
        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult I2C_GetNumChannels(out uint numChannels);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult I2C_GetChannelInfo(uint index, out FtDeviceInfo chaninfo);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult I2C_OpenChannel(uint index, out IntPtr ftHandle);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult I2C_CloseChannel(IntPtr ftHandle);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult I2C_InitChannel(IntPtr ftHandle, ref Ft_I2C_ChannelConfig config);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult I2C_DeviceRead(
           IntPtr handle,
           int deviceAddress,
           int sizeToTransfer,
           byte[] buffer,
           out int sizeTransfered,
           FtI2cTransferOptions options);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult I2C_DeviceWrite(
           IntPtr handle,
           int deviceAddress,
           int sizeToTransfer,
           byte[] buffer,
           out int sizeTransfered,
           FtI2cTransferOptions options);
        

        ////DLL Import SPI Function

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult SPI_GetNumChannels(out uint numChannels);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult SPI_GetChannelInfo(uint index, out FtDeviceInfo chaninfo);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult SPI_OpenChannel(uint index, out IntPtr ftHandle);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult SPI_CloseChannel(IntPtr ftHandle);


        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern FtResult SPI_InitChannel(IntPtr ftHandle, ref FtChannelConfig config);
        

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult SPI_ChangeCS(IntPtr handle, FtConfigOptions configOptions);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult SPI_IsBusy(IntPtr handle, out bool state);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult SPI_Read(
           IntPtr handle,
           byte[] buffer,
           int sizeToTransfer,
           out int sizeTransfered,
           FtSpiTransferOptions options);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult SPI_ReadWrite(
          IntPtr handle,
          byte[] inBuffer,
          byte[] outBuffer,
          int sizeToTransfer,
          out int sizeTransferred,
          FtSpiTransferOptions transferOptions);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult SPI_Write(
            IntPtr handle,
            byte[] buffer,
            int sizeToTransfer,
            out int sizeTransfered,
            FtSpiTransferOptions options);


        //// Global Functions
        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult FT_WriteGPIO(
            IntPtr handle,
            byte dir,
            byte value);

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public extern static FtResult FT_ReadGPIO(
            IntPtr handle,
            out byte value);

        public static void Init()
        {
            if (Interlocked.Increment(ref _initializations) == 1)
                Init_libMPSSE();
        }

        public static void Cleanup()
        {
            if (Interlocked.Decrement(ref _initializations) == 0)
                Cleanup_libMPSSE();
        }

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        private extern static void Init_libMPSSE();

        [DllImportAttribute("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        private extern static void Cleanup_libMPSSE();




        /// <summary>
        /// Custom I2C Function
        /// </summary>
        /// 
        public void CleanDLL()
        {
            Cleanup();
        }

        public FtResult CheckDeviceExist(out uint CH_IN_Host)
        {
            FtResult FtStatus = SPI_GetNumChannels(out CH_IN_Host);
            return FtStatus;
        }

        public FtResult CheckI2CCheannel(out uint CH_IN_Host)
        {
            FtResult FtStatus = I2C_GetNumChannels(out CH_IN_Host);
            return FtStatus;
        }

        public FtResult I2C_Init(out IntPtr ftHandle, PortInfo _PortInfo)
        {
            Init_libMPSSE();
            if (I2C_OpenChannel(0, out ftHandle) == FtResult.Ok)
            {
                FtResult ftStatus= I2C_InitChannel(ftHandle, ref _PortInfo.I2C_Channel_Conf);
                return ftStatus;
            }
            else
            {
                return FtResult.InvalidHandle;
            }
        }

        public FtResult I2C_DeInit(IntPtr ftHandle)
        {
            FtResult ftStatus = FtResult.InvalidHandle;
            try
            {
                ftStatus = I2C_CloseChannel(ftHandle);
                Cleanup_libMPSSE();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }
            return ftStatus;
        }

        public FtResult I2C_SEQ_Write(IntPtr FtHandle, byte DeviceAddr, byte[] W_Data)
        {
            int ByteToTrans;
            int ByteTransfered;
            FtResult FtStatus;

            ByteToTrans = W_Data.Length;
            FtStatus = I2C_DeviceWrite(FtHandle, (byte)(DeviceAddr >> 1), ByteToTrans, W_Data, out ByteTransfered, FtI2cTransferOptions.START_BIT | FtI2cTransferOptions.STOP_BIT | FtI2cTransferOptions.NACK_LAST_BYTE);
            return FtStatus;
        }

        public FtResult I2C_SEQ_Read(IntPtr FtHandle, byte DeviceAddr, byte[] W_Data, byte[] R_Data, out byte RDataLength)
        {
            FtResult FtStatus;
            int ByteTransfered;
            byte[] buffer = new byte[128];

            FtStatus = I2C_SEQ_Write(FtHandle, DeviceAddr, W_Data);
            if (FtStatus != FtResult.Ok)
            {
                RDataLength = 0;
                return FtStatus;
            }
                
            Thread.Sleep(300);
            FtStatus = I2C_DeviceRead(FtHandle, (byte)(DeviceAddr >> 1), 128, buffer, out ByteTransfered, FtI2cTransferOptions.START_BIT | FtI2cTransferOptions.NACK_LAST_BYTE);
            string inputstring = BitConverter.ToString(buffer).Replace("-", " ");
            GlobalData.log.Debug("I2C_DeviceRead: " + inputstring);
            RDataLength = buffer[1];
            if (RDataLength > 0x7D)
            {
                RDataLength -= 0x7D;
                Array.Copy(buffer, R_Data, RDataLength);
            }
            else
            {
                GlobalData.log.Debug("Error reading return length");
                MessageBox.Show("Read 回傳長度錯誤");
                FtStatus = FtResult.InvalidHandle;
            }
            return FtStatus;

        }

        public FtResult I2C_Device_Info(uint index,out FtDeviceInfo devicenfo)
        {
            FtResult FtStatus = I2C_GetChannelInfo(index, out devicenfo);
            return FtStatus;
        }
    }
}
