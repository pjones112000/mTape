using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;
using System.IO;
using System.Management;
using System.Collections.Generic;

namespace mtape
{

    #region Typedefenitions
    //this is for compatibility with API types
    using BOOL = System.Int32;
    #endregion

    /// <summary>
    /// Low level Tape Engine based on Win API
    /// </summary>
    public class DiskWinAPI
    {
        private String m_diskDrive = string.Empty;
        private Boolean m_diskOpen = false;


        #region Private variables

        public string DiskDrive
        {
            get { return m_diskDrive; }
            set { m_diskDrive = value; }
        }

        /// <summary>
        /// Retrieves a dictionary of information about the specified drive
        /// </summary>
        /// <param name="DriveLetter"></param>
        /// <returns>Dictionary of the properites of the specified drive</returns>
        public Dictionary<String,String> getDriveInfo(String DriveLetter)
        {
            Dictionary<String, String> retValue = new Dictionary<string, string>();
            DriveInfo[] drives = DriveInfo.GetDrives();
            

            foreach (DriveInfo drive in drives)
            {
                Dictionary<String, String> tmpDict = new Dictionary<string, string>();
                if (drive.Name.ToLower() == DriveLetter.ToLower() + @":\")
                {
                    tmpDict.Add("Name", drive.Name);
                    tmpDict.Add("VolumeLabel", drive.VolumeLabel);
                    tmpDict.Add("RootDirectory", String.Format($"{drive.RootDirectory}"));
                    tmpDict.Add("DriveType", String.Format($"{drive.DriveType}"));
                    tmpDict.Add("DriveFormat", String.Format($"{drive.DriveFormat}"));
                    tmpDict.Add("IsReady", String.Format($"{drive.IsReady}"));
                    tmpDict.Add("TotalSize", String.Format($"{drive.TotalSize}"));
                    tmpDict.Add("TotalFreeSpace", String.Format($"{drive.TotalFreeSpace}"));
                    tmpDict.Add("AvailableFreeSpace", String.Format($"{drive.AvailableFreeSpace}"));
                    retValue = tmpDict;
                    break;
                }
                
            }
              //          ManagementScope scope = new ManagementScope("\\\\.\\root\\cimv2");
              //          scope.Connect();
              //          ObjectQuery query = new ObjectQuery(@"SELECT DeviceID, Model, Name, Size, Status, Partitions, 
              //TotalTracks, TotalSectors, BytesPerSector, SectorsPerTrack, TotalCylinders, TotalHeads, 
              //TracksPerCylinder, CapabilityDescriptions 
              //FROM Win32_DiskDrive WHERE MediaType = 'Fixed hard disk media'");
              //          ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
              //          foreach (ManagementObject wmi_HD in searcher.Get())
              //          {
              //              foreach (PropertyData property in wmi_HD.Properties)
              //                  Console.WriteLine($"{property.Name} = {property.Value}");
              //              var capabilities = wmi_HD["CapabilityDescriptions"] as string[];
              //              if (capabilities != null)
              //              {
              //                  Console.WriteLine("Capabilities");
              //                  foreach (var capability in capabilities)
              //                      Console.WriteLine($"  {capability}");
              //              }
              //              Console.WriteLine("-----------------------------------");
              //          }
            return retValue;
        }

        #endregion

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr CreateFile(
         string lpFileName,
         uint dwDesiredAccess,
         uint dwShareMode,
         IntPtr SecurityAttributes,
         uint dwCreationDisposition,
         uint dwFlagsAndAttributes,
         IntPtr hTemplateFile
    );

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool DeviceIoControl(
            IntPtr hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            IntPtr lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped
        );

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool DeviceIoControl(
            IntPtr hDevice,
            uint dwIoControlCode,
            byte[] lpInBuffer,
            uint nInBufferSize,
            IntPtr lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);

        private IntPtr handle = IntPtr.Zero;

        const uint GENERIC_READ = 0x80000000;
        const uint GENERIC_WRITE = 0x40000000;
        const int FILE_SHARE_READ = 0x1;
        const int FILE_SHARE_WRITE = 0x2;
        const int FSCTL_LOCK_VOLUME = 0x00090018;
        const int FSCTL_DISMOUNT_VOLUME = 0x00090020;
        const int IOCTL_STORAGE_EJECT_MEDIA = 0x2D4808;
        const int IOCTL_STORAGE_MEDIA_REMOVAL = 0x002D4804;

        /// <summary>
        /// Constructor for the USBEject class
        /// </summary>
        /// <param name="driveLetter">This should be the drive letter. Format: F:/, C:/..</param>

        public bool EjectDisk(string driveLetter)
        {
            
            string filename = @"\\.\" + driveLetter[0] + ":";
            IntPtr devicePtr = CreateFile(filename, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, 0x3, 0, IntPtr.Zero);
            if (devicePtr != IntPtr.Zero)
            {
                return Eject(devicePtr);
            }
            else
            {
                return false;
            }
        }

        private bool Eject(IntPtr handle)
        {
            bool result = false;

            if (LockVolume(handle) && DismountVolume(handle))
            {
                PreventRemovalOfVolume(handle, false);
                result = AutoEjectVolume(handle);
            }
            CloseHandle(handle);
            return result;
        }

        private bool LockVolume(IntPtr handle)
        {
            uint byteReturned;

            for (int i = 0; i < 10; i++)
            {
                if (DeviceIoControl(handle, FSCTL_LOCK_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out byteReturned, IntPtr.Zero))
                {
                    return true;
                }
                Thread.Sleep(500);
            }
            return false;
        }

        private bool PreventRemovalOfVolume(IntPtr handle, bool prevent)
        {
            byte[] buf = new byte[1];
            uint retVal;

            buf[0] = (prevent) ? (byte)1 : (byte)0;
            return DeviceIoControl(handle, IOCTL_STORAGE_MEDIA_REMOVAL, buf, 1, IntPtr.Zero, 0, out retVal, IntPtr.Zero);
        }

        private bool DismountVolume(IntPtr handle)
        {
            uint byteReturned;
            return DeviceIoControl(handle, FSCTL_DISMOUNT_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out byteReturned, IntPtr.Zero);
        }

        private bool AutoEjectVolume(IntPtr handle)
        {
            uint byteReturned;
            return DeviceIoControl(handle, IOCTL_STORAGE_EJECT_MEDIA, IntPtr.Zero, 0, IntPtr.Zero, 0, out byteReturned, IntPtr.Zero);
        }

        private bool CloseVolume(IntPtr handle)
        {
            return CloseHandle(handle);
        }
    }

    #region ApplicationException
    /// <summary>
    /// Exception that will be thrown by tape
    /// Engine when one of WIN32 APIs terminates 
    /// with error code 
    /// </summary>
    public class DiskWinAPIException : ApplicationException
    {
        public DiskWinAPIException(string methodName, int win32ErroCode) :
            base(string.Format(
               "WIN32 API method failed : {0} failed with error code {1}",
               methodName,
               win32ErroCode
           ))
        { }
    }
    #endregion
}