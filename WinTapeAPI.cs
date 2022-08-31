using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;


namespace mtape
{

    #region Typedefenitions
    //this is for compatibility with API types
    using BOOL = System.Int32;
    #endregion

    /// <summary>
    /// Low level Tape Engine based on Win API
    /// </summary>
    public class TapeWinAPI
    {
        private String m_tapeDrive = string.Empty;
        private Boolean m_TapeOpen = false;


        #region Private variables

        public string TapeDrive
        {
            get { return m_tapeDrive; }
            set { m_tapeDrive = value; }
        }

        /// <summary>
        /// Tape device handle
        /// </summary>
        private SafeFileHandle m_handleDeviceValue = null;

        /// <summary>
        /// Used by GetTapeParameters API call for Tape drive
        /// </summary>
        private Nullable<TapeGetDriveParameters> m_getDriveParamsStruct = null;

        /// <summary>
        /// Used by GetMediaParameters API call for Tape media
        /// </summary>
        private Nullable<TapeGetMediaParameters> m_getMediaParamsStruct = null;

        /// <summary>
        /// Used by SetTapeParameters API call for Tape media
        /// </summary>
        private TapeSetMediaParameters m_setMediaParamsStruct;

        /// <summary>
        /// Used by SetTapeParameters API call for Tape drive 
        /// </summary>
        private TapeSetDriveParameters m_setDriveParamsStruct;

        /// <summary>
        /// Member holding Tape drive number in the system
        /// i.e. for TAPE0 it's 0
        /// </summary>
        private int m_tapeDriveNumber;

        /// <summary>
        /// Member holding Tape drive symbolic name
        /// i.e. \\.\TAPE0
        /// </summary>
        private string m_tapeDriveName;

        #endregion

        #region Nested Types

        /// <summary>
        /// TAPE_GAET_MEDIA_PARAMTERS structure
        /// See MSDN for details
        /// </summary>
        [StructLayout(LayoutKind.Sequential/*, Pack = 1*/)]
        private struct TapeGetMediaParameters
        {
            public long Capacity;
            public long Remaining;
            public uint BlockSize;
            public uint PartitionCount;

            public byte IsWriteProtected;
        }

        /// <summary>
        /// TAPE_SET_MEDIA_PARAMETERS structure
        /// See MSDN for details
        /// </summary>
        [StructLayout(LayoutKind.Sequential/*, Pack = 1*/)]
        public struct TapeSetMediaParameters
        {
            public uint BlockSize;
        }

        /// <summary>
        /// TAPE_GET_DRIVE_PARAMETERS structure
        /// See MSDN for details
        /// </summary>
        [StructLayout(LayoutKind.Sequential/*, Pack = 1*/)]
        private struct TapeGetDriveParameters
        {
            public byte ECC;
            public byte driveCompression;
            public byte DataPadding;
            public byte ReportSetMarks;

            public uint DefaultBlockSize;
            public uint MaximumBlockSize;
            public uint MinimumBlockSize;
            public uint PartitionCount;

            public uint FeaturesLow;
            public uint FeaturesHigh;
            public uint EATWarningZone;
        }

        /// <summary>
        /// TAPE_SET_DRIVE_PARAMETERS strucutre
        /// See MSDN for details
        /// </summary>
        [StructLayout(LayoutKind.Sequential/*, Pack = 1*/)]
        public struct TapeSetDriveParameters
        {
            public byte ECC;
            public byte driveCompression;
            public byte DataPadding;
            public byte ReportSetmarks;

            public uint EOTWarningZoneSize;
        }
        #endregion

        #region Public constants
        /// <summary>
        /// General constants
        /// </summary>
        public const int FALSE = 0;
        public const int TRUE = 1;
        public const int NULL = 0;
        public const int DEVICE_TYPE_TAPE = 0;
        public const int DEVICE_TYPE_FILE = 1;

        public const short INVALID_HANDLE_VALUE = -1;

        public bool IsTapeOpen()
        {
            return m_TapeOpen;
        }

        /// <summary>
        /// file share modes
        /// </summary>
        public const uint FILE_SHARE_READ = 0x00000001;
        public const uint FILE_SHARE_WRITE = 0x00000002;
        public const uint FILE_SHARE_DELETE = 0x00000004;

        /// <summary>
        /// file creation disposition
        /// </summary>
        public const uint CREATE_NEW = 1;
        public const uint CREATE_ALWAYS = 2;
        public const uint OPEN_EXISTING = 3;
        public const uint OPEN_ALWAYS = 4;
        public const uint TRUNCATE_EXISTING = 5;

        /// <summary>
        /// file attributes
        /// </summary>
        public const uint FILE_ATTRIBUTE_READONLY = 0x00000001;
        public const uint FILE_ATTRIBUTE_HIDDEN = 0x00000002;
        public const uint FILE_ATTRIBUTE_SYSTEM = 0x00000004;
        public const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
        public const uint FILE_ATTRIBUTE_ARCHIVE = 0x00000020;
        public const uint FILE_ATTRIBUTE_DEVICE = 0x00000040;
        public const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;
        public const uint FILE_ATTRIBUTE_TEMPORARY = 0x00000100;
        public const uint FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200;
        public const uint FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400;
        public const uint FILE_ATTRIBUTE_COMPRESSED = 0x00000800;
        public const uint FILE_ATTRIBUTE_OFFLINE = 0x00001000;
        public const uint FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000;
        public const uint FILE_ATTRIBUTE_ENCRYPTED = 0x00004000;
        public const uint FILE_ATTRIBUTE_VIRTUAL = 0x00010000;

        /// <summary>
        /// file flags
        /// </summary>
        public const uint FILE_FLAG_WRITE_THROUGH = 0x80000000;
        public const uint FILE_FLAG_OVERLAPPED = 0x40000000;
        public const uint FILE_FLAG_NO_BUFFERING = 0x20000000;
        public const uint FILE_FLAG_RANDOM_ACCESS = 0x10000000;
        public const uint FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000;
        public const uint FILE_FLAG_DELETE_ON_CLOSE = 0x04000000;
        public const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        public const uint FILE_FLAG_POSIX_SEMANTICS = 0x01000000;
        public const uint FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000;
        public const uint FILE_FLAG_OPEN_NO_RECALL = 0x00100000;
        public const uint FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000;

        /*
        /// <summary>
        /// file security features
        /// </summary>
        public const uint SECURITY_ANONYMOUS
        public const uint SECURITY_CONTEXT_TRACKING
        public const uint SECURITY_DELEGATION
        public const uint SECURITY_EFFECTIVE_ONLY
        public const uint SECURITY_IDENTIFICATION
        public const uint SECURITY_IMPERSONATION
        */

        /// <summary>
        /// file desired access
        /// </summary>
        public const uint GENERIC_READ = 0x80000000;
        public const uint GENERIC_WRITE = 0x40000000;

        /// <summary>
        /// Tape device preparation Operation
        /// </summary>
        public const uint TAPE_LOAD = 0;
        public const uint TAPE_UNLOAD = 1;
        public const uint TAPE_TENSION = 2;
        public const uint TAPE_LOCK = 3;
        public const uint TAPE_UNLOCK = 4;
        public const uint TAPE_FORMAT = 5;

        /// <summary>
        /// return codes
        /// </summary>
        public const uint NO_ERROR = 0;
        public const uint ERROR_BEGINNING_OF_MEDIA = 1102;
        public const uint ERROR_BUS_RESET = 1111;
        public const uint ERROR_DEVICE_NOT_PARTITIONED = 1107;
        public const uint ERROR_END_OF_MEDIA = 1100;
        public const uint ERROR_FILEMARK_DETECTED = 1101;
        public const uint ERROR_INVALID_BLOCK_LENGTH = 1106;
        public const uint ERROR_MEDIA_CHANGED = 1110;
        public const uint ERROR_NO_DATA_DETECTED = 1104;
        public const uint ERROR_NO_MEDIA_IN_DRIVE = 1112;
        public const uint ERROR_NOT_SUPPORTED = 50;
        public const uint ERROR_PARTITION_FAILURE = 1105;
        public const uint ERROR_SETMARK_DETECTED = 1103;
        public const uint ERROR_UNABLE_TO_LOCK_MEDIA = 1108;
        public const uint ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109;
        public const uint ERROR_WRITE_PROTECT = 19;
        public const uint ERROR_DEVICE_REQUIRES_CLEANING = 1165;

        /// <summary>
        /// types of positioning
        /// </summary>
        public const uint TAPE_REWIND = 0;
        public const uint TAPE_ABSOLUTE_BLOCK = 1;
        public const uint TAPE_LOGICAL_BLOCK = 2;
        public const uint TAPE_PSEUDO_LOGICAL_BLOCK = 3;
        public const uint TAPE_SPACE_END_OF_DATA = 4;
        public const uint TAPE_SPACE_RELATIVE_BLOCKS = 5;
        public const uint TAPE_SPACE_FILEMARKS = 6;
        public const uint TAPE_SPACE_SEQUENTIAL_FMKS = 7;
        public const uint TAPE_SPACE_SETMARKS = 8;
        public const uint TAPE_SPACE_SEQUENTIAL_SMKS = 9;

        /// <summary>
        /// position type
        /// </summary>
        public const uint TAPE_ABSOLUTE_POSITION = 0;
        public const uint TAPE_LOGICAL_POSITION = 1;
        public const uint TAPE_PSEUDO_LOGICAL_POSITION = 2;

        /// <summary>
        /// get tape parameters operation type
        /// </summary>
        public const uint GET_TAPE_MEDIA_INFORMATION = 0;
        public const uint GET_TAPE_DRIVE_INFORMATION = 1;

        /// <summary>
        /// Drive parameters features low
        /// </summary>
        public const uint TAPE_DRIVE_FIXED = 0x00000001;
        public const uint TAPE_DRIVE_SELECT = 0x00000002;
        public const uint TAPE_DRIVE_INITIATOR = 0x00000004;
        public const uint TAPE_DRIVE_ERASE_SHORT = 0x00000010;
        public const uint TAPE_DRIVE_ERASE_LONG = 0x00000020;
        public const uint TAPE_DRIVE_ERASE_BOP_ONLY = 0x00000040;
        public const uint TAPE_DRIVE_ERASE_IMMEDIATE = 0x00000080;
        public const uint TAPE_DRIVE_TAPE_CAPACITY = 0x00000100;
        public const uint TAPE_DRIVE_TAPE_REMAINING = 0x00000200;
        public const uint TAPE_DRIVE_FIXED_BLOCK = 0x00000400;
        public const uint TAPE_DRIVE_VARIABLE_BLOCK = 0x00000800;
        public const uint TAPE_DRIVE_WRITE_PROTECT = 0x00001000;
        public const uint TAPE_DRIVE_EOT_WZ_SIZE = 0x00002000;
        public const uint TAPE_DRIVE_ECC = 0x00010000;
        public const uint TAPE_DRIVE_COMPRESSION = 0x00020000;
        public const uint TAPE_DRIVE_PADDING = 0x00040000;
        public const uint TAPE_DRIVE_REPORT_SMKS = 0x00080000;
        public const uint TAPE_DRIVE_GET_ABSOLUTE_BLK = 0x00100000;
        public const uint TAPE_DRIVE_GET_LOGICAL_BLK = 0x00200000;
        public const uint TAPE_DRIVE_SET_EOT_WZ_SIZE = 0x00400000;
        public const uint TAPE_DRIVE_EJECT_MEDIA = 0x01000000;
        public const uint TAPE_DRIVE_CLEAN_REQUESTS = 0x02000000;
        public const uint TAPE_DRIVE_SET_CMP_BOP_ONLY = 0x04000000;

        /// <summary>
        /// Drive parameters features high
        /// </summary>
        public const uint TAPE_DRIVE_LOAD_UNLOAD = 0x80000001;
        public const uint TAPE_DRIVE_TENSION = 0x80000002;
        public const uint TAPE_DRIVE_LOCK_UNLOCK = 0x80000004;
        public const uint TAPE_DRIVE_REWIND_IMMEDIATE = 0x80000008;
        public const uint TAPE_DRIVE_SET_BLOCK_SIZE = 0x80000010;
        public const uint TAPE_DRIVE_LOAD_UNLD_IMMED = 0x80000020;
        public const uint TAPE_DRIVE_TENSION_IMMED = 0x80000040;
        public const uint TAPE_DRIVE_LOCK_UNLK_IMMED = 0x80000080;
        public const uint TAPE_DRIVE_SET_ECC = 0x80000100;
        public const uint TAPE_DRIVE_SET_COMPRESSION = 0x80000200;
        public const uint TAPE_DRIVE_SET_PADDING = 0x80000400;
        public const uint TAPE_DRIVE_SET_REPORT_SMKS = 0x80000800;
        public const uint TAPE_DRIVE_ABSOLUTE_BLK = 0x80001000;
        public const uint TAPE_DRIVE_ABS_BLK_IMMED = 0x80002000;
        public const uint TAPE_DRIVE_LOGICAL_BLK = 0x80004000;
        public const uint TAPE_DRIVE_LOG_BLK_IMMED = 0x80008000;
        public const uint TAPE_DRIVE_END_OF_DATA = 0x80010000;
        public const uint TAPE_DRIVE_RELATIVE_BLKS = 0x80020000;
        public const uint TAPE_DRIVE_FILEMARKS = 0x80040000;
        public const uint TAPE_DRIVE_SEQUENTIAL_FMKS = 0x80080000;
        public const uint TAPE_DRIVE_SETMARKS = 0x80100000;
        public const uint TAPE_DRIVE_SEQUENTIAL_SMKS = 0x80200000;
        public const uint TAPE_DRIVE_REVERSE_POSITION = 0x80400000;
        public const uint TAPE_DRIVE_SPACE_IMMEDIATE = 0x80800000;
        public const uint TAPE_DRIVE_WRITE_SETMARKS = 0x81000000;
        public const uint TAPE_DRIVE_WRITE_FILEMARKS = 0x82000000;
        public const uint TAPE_DRIVE_WRITE_SHORT_FMKS = 0x84000000;
        public const uint TAPE_DRIVE_WRITE_LONG_FMKS = 0x88000000;
        public const uint TAPE_DRIVE_WRITE_MARK_IMMED = 0x90000000;
        public const uint TAPE_DRIVE_FORMAT = 0xA0000000;
        public const uint TAPE_DRIVE_FORMAT_IMMEDIATE = 0xC0000000;

        /// <summary>
        /// set tape parameters operation type
        /// </summary>
        public const uint SET_TAPE_MEDIA_INFORMATION = 0;
        public const uint SET_TAPE_DRIVE_INFORMATION = 1;

        /// <summary>
        /// partition method
        /// </summary>
        public const uint TAPE_FIXED_PARTITIONS = 0;
        public const uint TAPE_SELECT_PARTITIONS = 1;
        public const uint TAPE_INITIATOR_PARTITIONS = 2;

        /// <summary>
        /// Erase type
        /// </summary>
        public const uint TAPE_ERASE_SHORT = 0;
        public const uint TAPE_ERASE_LONG = 1;

        /// <summary>
        /// tapemark types
        /// </summary>
        public const uint TAPE_SETMARKS = 0;
        public const uint TAPE_FILEMARKS = 1;
        public const uint TAPE_SHORT_FILEMARKS = 2;
        public const uint TAPE_LONG_FILEMARKS = 3;
        #endregion

        #region PInvoke
        //Provides access to Win32 API DLLs

        /// <summary>
        /// CreateFile Win API
        /// </summary>
        /// <param name="lpFileName"></param>
        /// <param name="dwDesiredAccess"></param>
        /// <param name="dwShareMode"></param>
        /// <param name="lpSecurityAttributes"></param>
        /// <param name="dwCreationDisposition"></param>
        /// <param name="dwFlagsAndAttributes"></param>
        /// <param name="hTemplateFile"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile
            );

        /// <summary>
        /// PrepareTape Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="prepareType"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern int PrepareTape(
            SafeFileHandle handle,
            int prepareType,
            BOOL isImmediate
            );

        /// <summary>
        /// SetTapePosition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="positionMethod"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint SetTapePosition(
            SafeFileHandle handle,
            int positionMethod,
            int partition,
            int offsetLow,
            int offsetHigh,
            BOOL isImmediate
            );

        /// <summary>
        /// GetTapePosition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="positionType"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapePosition(
            SafeFileHandle handle,
            uint positionType,
            ref uint partition,
            ref uint offsetLow,
            ref uint offsetHigh
            );

        /// <summary>
        /// GetTapeParameters Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="operationType">
        /// GET_TAPE_DRIVE_INFORMATION
        /// or
        /// GET_TAPE_MEDIA_INFORMATION
        /// </param>
        /// <param name="size"></param>
        /// <param name="mediaOrDriveInfo"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapeParameters(
           SafeFileHandle handle,
           uint operationType,
           ref uint size,
           IntPtr mediaOrDriveInfo
           );

        /// <summary>
        /// SetTapeParameters Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="operationType"></param>
        /// <param name="mediaOrDriveInfo"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint SetTapeParameters(
           SafeFileHandle handle,
           uint operationType,
           IntPtr mediaOrDriveInfo
            );

        /// <summary>
        /// CreateTapePartition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="PartitionMethod"></param>
        /// <param name="count"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint CreateTapePartition(
            SafeFileHandle handle,
            uint PartitionMethod,
            uint count,
            uint size
            );

        /// <summary>
        /// EraseTape Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="eraseType"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint EraseTape(
        SafeFileHandle handle,
        uint eraseType,
        BOOL isImmediate
        );

        /// <summary>
        /// WriteTapemark Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="tapemarkType"></param>
        /// <param name="tapemarkCount"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint WriteTapemark(
        SafeFileHandle handle,
        uint tapemarkType,
        uint tapemarkCount,
        BOOL isImmediate
        );

        /// <summary>
        /// GetTapeStatus Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapeStatus(
        SafeFileHandle handle
        );

        /// <summary>
        /// ReadFile Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="buffer"></param>
        /// <param name="numberOfBytesToRead"></param>
        /// <param name="numberOfBytesRead"></param>
        /// <param name="overlappedBuffer"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern bool ReadFile(
        SafeFileHandle handle,
        IntPtr buffer,
        uint numberOfBytesToRead,
        ref uint numberOfBytesRead,
        IntPtr overlappedBuffer
        );

        /// <summary>
        /// WriteFile Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="buffer"></param>
        /// <param name="numberOfBytesToWrite"></param>
        /// <param name="numberOfBytesWritten"></param>
        /// <param name="overlappedBuffer"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern bool WriteFile(
        SafeFileHandle handle,
        byte[] lpBuffer,
        uint numberOfBytesToWrite,
        ref uint numberOfBytesWritten,
        IntPtr overlappedBuffer
        );


        /// <summary>
        /// GetLastError Win API
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetLastError();

        #endregion

        #region Constructor
        /// <summary>
        /// Default constructor
        /// </summary>
        public TapeWinAPI()
        {
            //            m_getMediaParamsStruct = new TapeGetMediaParameters();
            //            m_getDriveParamsStruct = new TapeGetDriveParameters();
            //            m_setMediaParamsStruct = new TapeSetMediaParameters();
            //            m_setDriveParamsStruct = new TapeSetDriveParameters();
        }
        #endregion

        #region Public methods

        /// <summary>
        /// Attempts to open tape drive
        /// </summary>
        /// <param name="strTapeDriveName"></param>
        public bool Open(string strTapeDriveName)
        {
            // initialize private member with passed string 
            m_tapeDriveName = strTapeDriveName;
            if (m_handleDeviceValue == null || m_handleDeviceValue.IsClosed)
            {
                // try to open device
                m_handleDeviceValue = CreateFile(
                                m_tapeDriveName,
                                GENERIC_READ | GENERIC_WRITE,
                                0,
                                IntPtr.Zero,
                                OPEN_EXISTING,
                                FILE_ATTRIBUTE_NORMAL,
                                IntPtr.Zero
                                );

                if (m_handleDeviceValue.IsInvalid)
                {
                    // could not open
                    m_tapeDriveName = null;
                    return false;
                }

                // initialize private member
                m_tapeDriveNumber = int.Parse(strTapeDriveName.Remove(0, 8));
                this.GetTapeStatus();// this is to reset ERROR_MEDIA_CHANGED status
                m_TapeOpen = true;
                m_tapeDrive = strTapeDriveName.Substring(strTapeDriveName.Length - 5);
                Thread.Sleep(500);
            }
            else
            {
                m_TapeOpen = true;
            }
            return true;
        }

        /// <summary>
        /// Attempts to open tape drive
        /// </summary>
        /// <param name="nTapeDevice"></param>
        public bool Open(uint nTapeDevice)
        {
            return Open("\\\\.\\TAPE" + nTapeDevice);
        }

        /// <summary>
        /// Loads tape
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Load(ref uint returnCode)
        {
            return Prepare(TAPE_LOAD, ref returnCode, TRUE);
        }

        /// <summary>
        /// Unloads tape
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Unload(ref uint returnCode)
        {
            return Prepare(TAPE_UNLOAD, ref returnCode, TRUE);
        }

        /// <summary>
        /// Locks tape in a drive
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Lock(ref uint returnCode)
        {
            return Prepare(TAPE_LOCK, ref returnCode, TRUE);
        }

        /// <summary>
        /// Unlocks tape in a drive
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Unlock(ref uint returnCode)
        {
            return Prepare(TAPE_UNLOCK, ref returnCode, TRUE);
        }

        /// <summary>
        /// Formats tape
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool FormatTape(ref uint returnCode)
        {
            return Prepare(TAPE_FORMAT, ref returnCode, TRUE);
        }

        /// <summary>
        /// Adjusts tape tension
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool AdjustTension(ref uint returnCode)
        {
            return Prepare(TAPE_TENSION, ref returnCode, TRUE);
        }

        /// <summary>
        /// Rewinds tape to BOD
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Rewind(ref uint returnCode)
        {
            return SetTapePosition(TAPE_REWIND, 0, 0, 0, FALSE, ref returnCode);
        }

        /// <summary>
        /// Rewinds tape to EOD
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SeekToEOD(ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_END_OF_DATA, 0, 0, 0, FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning tape to absolute block
        /// </summary>
        /// <param name="blockAddress"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SeekToAbsoluteBlock(long blockAddress, ref uint returnCode)
        {
            return SetTapePosition(TAPE_ABSOLUTE_BLOCK, 0, (uint)blockAddress, (uint)(blockAddress >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning tape to logical block
        /// </summary>
        /// <param name="blockAddress"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SeekToLogicalBlock(long blockAddress, ref uint returnCode)
        {
            return SetTapePosition(TAPE_LOGICAL_BLOCK, 0, (uint)blockAddress, (uint)(blockAddress >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning (backward or forward) by spacing number of filemarks
        /// </summary>
        /// <param name="filemarksToSpace"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SpaceFileMarks(long filemarksToSpace, ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_FILEMARKS, 0, (uint)filemarksToSpace, (uint)(filemarksToSpace >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning (backward or forward) by spacing number of blocks
        /// </summary>
        /// <param name="blocksToSpace"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SpaceBlocks(long blocksToSpace, ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_FILEMARKS, 0, (uint)blocksToSpace, (uint)(blocksToSpace >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Reads from the tape
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="bytesToRead"></param>
        /// <param name="bytesRead"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Read(ref byte[] buffer, uint bytesToRead, ref uint bytesRead, ref uint returnCode)
        {// consider using try-catch if this method introduces exceptions...

            returnCode = NO_ERROR;
            try
            {
                //allocate unmanaged memory safely
                GCHandle h = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                IntPtr ptrBuffer = Marshal.UnsafeAddrOfPinnedArrayElement(buffer, 0);
                h.Free();
                System.Console.WriteLine("About to read");
                // read
                if (ReadFile(m_handleDeviceValue, ptrBuffer, bytesToRead, ref bytesRead, IntPtr.Zero))
                {
                    System.Console.WriteLine("Just reading cleanly");
                    System.Console.WriteLine("Read: " + bytesRead);
                    return true;
                }
                returnCode = GetLastError();
                if (returnCode != 1101)
                {
                    System.Console.WriteLine("Just read with error.");
                    System.Console.WriteLine("Error : " + returnCode);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Read Error  " + ex.Message);
            }
            return false;
        }

        /// <summary>
        /// Writes to tape
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="bytesToWrite"></param>
        /// <param name="bytesWritten"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Write(ref byte[] buffer, uint bytesToWrite, ref uint bytesWritten, ref uint returnCode)
        {// consider using try-catch if this method introduces exceptions...

            returnCode = NO_ERROR;
            //write
            if (!WriteFile(m_handleDeviceValue, buffer, bytesToWrite, ref bytesWritten, IntPtr.Zero))
            {
                returnCode = GetLastError();
                return false;
            }
            return true;
        }

        /// <summary>
        /// Closes device handle and zero all associated members
        /// </summary>
        public void Close()
        {
            if (m_handleDeviceValue != null &&
                !m_handleDeviceValue.IsInvalid &&
                !m_handleDeviceValue.IsClosed)
            {
                m_handleDeviceValue.Close();
                m_tapeDriveName = null;
                m_tapeDriveNumber = 0;
                m_TapeOpen = false;
            }
        }

        /// <summary>
        /// Erase tape
        /// </summary>
        /// <param name="eraseType"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool Erase(uint eraseType, ref uint returnCode)
        {
            returnCode = EraseTape(m_handleDeviceValue, eraseType, FALSE);

            if (returnCode == NO_ERROR) return true;
            return false;
        }

        /// <summary>
        /// Partitions (reformats) the tape
        /// partMethod: 0 - fixed
        ///             1 - select
        ///             2 - initiator
        /// </summary>
        /// <param name="partMethod"></param>
        /// <param name="partCount"></param>
        /// <param name="partSizeMB"></param>
        /// <returns></returns>
        public bool CreatePartition(uint partMethod, uint partCount, uint partSizeMB, ref uint returnCode)
        {
            returnCode = CreateTapePartition(m_handleDeviceValue, partMethod, partCount, partSizeMB);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// Writes the number of tape marks
        /// </summary>
        /// <param name="marksCount"></param>
        /// <param name="marksType"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool WriteTapemark(uint marksCount, uint marksType, ref uint returnCode)
        {
            returnCode = WriteTapemark(m_handleDeviceValue, marksType, marksCount, FALSE);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// returns tape drive status
        /// </summary>
        public int GetTapeStatus()
        {
            return (int)GetTapeStatus(m_handleDeviceValue);
        }

        /// <summary>
        /// Returns current tape's position in logical or absolute blocks
        /// </summary>
        /// <param name="positionType"></param>
        /// <param name="partitionNumber"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool GetTapePosition(uint positionType, ref uint partitionNumber, ref uint offsetLow, ref uint offsetHigh, ref uint returnCode)
        {
            returnCode = GetTapePosition(m_handleDeviceValue, positionType, ref partitionNumber, ref offsetLow, ref offsetHigh);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        public Boolean GetHandleValid()
        {
            return !m_handleDeviceValue.IsInvalid;
        }

        /// <summary>
        /// Returns current tape's logical position
        /// </summary>
        /// <param name="Position"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool GetTapePosition(ref long Position, ref uint returnCode)
        {
            uint partition = 0;
            uint offsetLow = 0;
            uint offsetHigh = 0;

            returnCode = GetTapePosition(
                m_handleDeviceValue,
                TAPE_LOGICAL_POSITION,
                ref partition,
                ref offsetLow,
                ref offsetHigh);

            Position = (long)(offsetHigh * Math.Pow(2, 32) + offsetLow);

            if (returnCode == NO_ERROR)
            {
                return true;
            }
            else
            {
                Position = -1;
                return false;
            }

            return false;
        }

        /// <summary>
        /// This is test version
        /// It should be replaced by appropriate ApplicationException class later
        /// </summary>
        /// <param name="errcode"></param>
        /// <returns></returns>
        public static string ConvertErrCode(uint errcode)
        {
            string errorMessage = errcode.ToString() + ": " + new Win32Exception((int)errcode).Message;

            return errorMessage;
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns Device number
        /// </summary>
        public int TapeDriveNumber
        {
            get { return m_tapeDriveNumber; }
        }

        /// <summary>
        /// Returns string with the full device name/path
        /// </summary>
        public string TapeDriveName
        {
            get { return m_tapeDriveName; }
        }

        /// <summary>
        /// Returns Device Handle
        /// </summary>
        public SafeFileHandle HandleDeviceValue
        {
            get { return m_handleDeviceValue; }
        }

        /// <summary>
        /// Gets and Sets tape block size
        /// </summary>
        public int BlockSizeTape
        {
            get
            {
                if (!m_getMediaParamsStruct.HasValue)
                {
                    m_getMediaParamsStruct = new TapeGetMediaParameters();

                    uint returnCode = 0;

                    if (!this.GetMediaParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetMediaParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getMediaParamsStruct.Value.BlockSize;

            }
            set
            {
                uint returnCode = 0;

                m_setMediaParamsStruct.BlockSize = (uint)value;

                if (!this.SetMediaParameters(m_setMediaParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetMediaParameters", Marshal.GetLastWin32Error());
                }

            }
        }

        /// <summary>
        /// Returns total number of bytes on the current tape partition
        /// </summary>
        public long Capacity
        {
            get
            {
                if (!m_getMediaParamsStruct.HasValue)
                {
                    m_getMediaParamsStruct = new TapeGetMediaParameters();

                    uint returnCode = 0;

                    if (!this.GetMediaParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetMediaParameters", Marshal.GetLastWin32Error());

                    }
                }

                return m_getMediaParamsStruct.Value.Capacity;

            }
        }

        /// <summary>
        /// Returns number of bytes between the current position 
        /// and the end of the current tape partition
        /// </summary>
        public long Remaining
        {
            get
            {
                if (!m_getMediaParamsStruct.HasValue)
                {
                    m_getMediaParamsStruct = new TapeGetMediaParameters();

                    uint returnCode = 0;

                    if (!this.GetMediaParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetMediaParameters", Marshal.GetLastWin32Error());

                    }
                }

                return m_getMediaParamsStruct.Value.Remaining;

            }
        }

        /// <summary>
        /// Returns number of partitions on the tape
        /// </summary>
        public int PartitionCountTape
        {
            get
            {
                if (!m_getMediaParamsStruct.HasValue)
                {
                    m_getMediaParamsStruct = new TapeGetMediaParameters();

                    uint returnCode = 0;

                    if (!this.GetMediaParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetMediaParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getMediaParamsStruct.Value.PartitionCount;

            }
        }

        /// <summary>
        /// Returns true if tape is write-protected
        /// </summary>
        public bool IsWriteProtected
        {
            get
            {
                if (!m_getMediaParamsStruct.HasValue)
                {
                    m_getMediaParamsStruct = new TapeGetMediaParameters();

                    uint returnCode = 0;

                    if (!this.GetMediaParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetMediaParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (m_getMediaParamsStruct.Value.IsWriteProtected != 0);

            }
        }

        /// <summary>
        /// Returns default block size for the drive
        /// </summary>
        public int BlockSizeDrive
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.DefaultBlockSize;

            }
        }

        /// <summary>
        /// Returns maximum block size for the drive
        /// </summary>
        public int MaximumBlockSizeDrive
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.MaximumBlockSize;

            }
        }

        /// <summary>
        /// Returns minimum block size for the drive
        /// </summary>
        public int MinimumBlockSizeDrive
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.MinimumBlockSize;

            }
        }

        /// <summary>
        /// Returns maximum number of partitions that can be created on the drive
        /// </summary>
        public int PartitionCountDriveMaximum
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.PartitionCount;

            }
        }

        /// <summary>
        /// Returns Low-order bits of the device features flags
        /// </summary>
        public int FeaturesLowDrive
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.FeaturesLow;

            }
        }

        /// <summary>
        /// Returns High-order bits of the device features flags
        /// </summary>
        public int FeaturesHighDrive
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.FeaturesHigh;

            }
        }

        /// <summary>
        /// Returns the number of bytes between the end-of-tape warning
        /// and the physical end of the tape
        /// </summary>
        public int EOTWarningZoneSize
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (int)m_getDriveParamsStruct.Value.EATWarningZone;

            }
            set
            {
                uint returnCode = 0;

                // get current parameters
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                // transfer current parameters to the set structure
                CopyCurrentParamsDriveHelper();

                // initialize parameter
                m_setDriveParamsStruct.EOTWarningZoneSize = (uint)value;

                // set it finally
                if (!this.SetDriveParameters(m_setDriveParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetDriveParameters", Marshal.GetLastWin32Error());
                }
            }


        }

        /// <summary>
        /// Returns true if device supports hardware error correction
        /// Enables/Disables ECC support
        /// </summary>
        public bool ECC
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (m_getDriveParamsStruct.Value.ECC != 0);

            }
            set
            {
                uint returnCode = 0;

                // get current parameters
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                // transfer current parameters to the set structure
                CopyCurrentParamsDriveHelper();

                // initialize parameter
                if (value)
                {
                    m_setDriveParamsStruct.ECC = TRUE;
                }
                else
                {
                    m_setDriveParamsStruct.ECC = FALSE;
                }

                // set it finally
                if (!this.SetDriveParameters(m_setDriveParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetDriveParameters", Marshal.GetLastWin32Error());
                }
            }
        }

        /// <summary>
        /// Returns true if hardware compression is enabled
        /// Enables/Disables hardware compression
        /// </summary>
        public bool Compression
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (m_getDriveParamsStruct.Value.driveCompression != 0);

            }
            set
            {
                uint returnCode = 0;

                // get current parameters
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                // transfer current parameters to the set structure
                CopyCurrentParamsDriveHelper();

                // initialize parameter
                if (value)
                {
                    m_setDriveParamsStruct.driveCompression = TRUE;
                }
                else
                {
                    m_setDriveParamsStruct.driveCompression = FALSE;
                }

                // set it finally
                if (!this.SetDriveParameters(m_setDriveParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetDriveParameters", Marshal.GetLastWin32Error());
                }
            }
        }

        /// <summary>
        /// Returns true if data padding is enabled
        /// Enables/Disables data padding
        /// </summary>
        public bool DataPadding
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (m_getDriveParamsStruct.Value.DataPadding != 0);

            }
            set
            {
                uint returnCode = 0;

                // get current parameters
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                // transfer current parameters to the set structure
                CopyCurrentParamsDriveHelper();

                // initialize parameter
                if (value)
                {
                    m_setDriveParamsStruct.DataPadding = TRUE;
                }
                else
                {
                    m_setDriveParamsStruct.DataPadding = FALSE;
                }

                // set it finally
                if (!this.SetDriveParameters(m_setDriveParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetDriveParameters", Marshal.GetLastWin32Error());
                }
            }
        }

        /// <summary>
        /// Returns true if setmark reporting is enabled
        /// Enables/Disables setmark reporting
        /// </summary>
        public bool ReportSetmarks
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return (m_getDriveParamsStruct.Value.ReportSetMarks != 0);

            }
            set
            {
                uint returnCode = 0;

                // get current parameters
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                CopyCurrentParamsDriveHelper();

                if (value)
                {
                    m_setDriveParamsStruct.ReportSetmarks = TRUE;
                }
                else
                {
                    m_setDriveParamsStruct.ReportSetmarks = FALSE;
                }


                if (!this.SetDriveParameters(m_setDriveParamsStruct, ref returnCode))
                {
                    throw new TapeWinAPIException(
                        "SetDriveParameters", Marshal.GetLastWin32Error());
                }


            }
        }

        /// <summary>
        /// returns true if device supports data compression
        /// </summary>
        public bool IsCompressionCapable
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return ((m_getDriveParamsStruct.Value.FeaturesLow & TAPE_DRIVE_COMPRESSION) != 0);

            }
        }

        /// <summary>
        /// returns true if device supports variable block length mode
        /// </summary>
        public bool IsVariableBlockCapable
        {
            get
            {
                if (!m_getDriveParamsStruct.HasValue)
                {
                    m_getDriveParamsStruct = new TapeGetDriveParameters();

                    uint returnCode = 0;

                    if (!this.GetDriveParameters(ref returnCode))
                    {
                        throw new TapeWinAPIException(
                            "GetDriveParameters", Marshal.GetLastWin32Error());

                    }
                }

                return ((m_getDriveParamsStruct.Value.FeaturesLow & TAPE_DRIVE_VARIABLE_BLOCK) != 0);

            }
        }


        #endregion

        #region Private methods

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="prepareType"></param>
        /// <param name="returnCode"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        private bool Prepare(uint prepareType, ref uint returnCode, BOOL isImmediate)
        {
            //returnCode = PrepareTape(
            //    m_handleDeviceValue,
            //    prepareType,
            //    isImmediate
            //    );
            int tmp = 0;
            tmp = Convert.ToInt32(prepareType);
            int tmpr = 0;
            tmpr = Convert.ToInt32(returnCode);
                        
            tmpr = PrepareTape(
                m_handleDeviceValue,
                tmp,
                isImmediate
                );

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="positionMethod"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <param name="isImmediate"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool SetTapePosition(uint positionMethod, uint partition, uint offsetLow, uint offsetHigh, BOOL isImmediate, ref uint returnCode)
        {
            int m_method = Convert.ToInt32(positionMethod);
            int m_partition = Convert.ToInt32(partition);
            int m_offsetLow = Convert.ToInt32(offsetLow);
            int m_offsetHigh = Convert.ToInt32(offsetHigh);

            returnCode = SetTapePosition(
                m_handleDeviceValue,
                m_method,
                m_partition,
                m_offsetLow,
                m_offsetHigh,
                isImmediate
                );

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SetDriveParameters(TapeSetDriveParameters parameters, ref uint returnCode)
        {
            //allocate unmanaged memory safely
            GCHandle h = GCHandle.Alloc(parameters, GCHandleType.Pinned);
            IntPtr ptrParameters = h.AddrOfPinnedObject();
            h.Free();

            returnCode = SetTapeParameters(
            m_handleDeviceValue,
            SET_TAPE_DRIVE_INFORMATION,
            ptrParameters
            );

            m_getDriveParamsStruct = null;// previous data is not valid anymore

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        public bool SetMediaParameters(TapeSetMediaParameters parameters, ref uint returnCode)
        {
            //allocate unmanaged memory safely
            GCHandle h = GCHandle.Alloc(parameters, GCHandleType.Pinned);
            IntPtr ptrParameters = h.AddrOfPinnedObject();
            h.Free();

            returnCode = SetTapeParameters(
            m_handleDeviceValue,
            SET_TAPE_MEDIA_INFORMATION,
            ptrParameters
            );

            m_getMediaParamsStruct = null;// previous data is not valid anymore

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// Retrieves Media parameters into private member
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool GetMediaParameters(ref uint returnCode)
        {
            IntPtr ptr = IntPtr.Zero;

            returnCode = NO_ERROR;

            //allocate unmanaged memory safely
            uint size = (uint)Marshal.SizeOf(m_getMediaParamsStruct);
            ptr = Marshal.AllocHGlobal((int)size);
            Marshal.StructureToPtr(m_getMediaParamsStruct, ptr, false);

            returnCode = GetTapeParameters(
            m_handleDeviceValue,
            GET_TAPE_MEDIA_INFORMATION,
            ref size,
            ptr);

            // Get managed media Info
            m_getMediaParamsStruct = (TapeGetMediaParameters)
                Marshal.PtrToStructure(ptr, typeof(TapeGetMediaParameters));

            Marshal.FreeHGlobal(ptr);//release memory

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// Retrieves Drive parameters into private member
        /// </summary>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool GetDriveParameters(ref uint returnCode)
        {
            IntPtr ptr = IntPtr.Zero;

            returnCode = NO_ERROR;

            //allocate unmanaged memory
            uint size = (uint)Marshal.SizeOf(m_getDriveParamsStruct);
            ptr = Marshal.AllocHGlobal((int)size);
            Marshal.StructureToPtr(m_getDriveParamsStruct, ptr, false);

            returnCode = GetTapeParameters(
            m_handleDeviceValue,
            GET_TAPE_DRIVE_INFORMATION,
            ref size,
            ptr);

            // Get managed media Info
            m_getDriveParamsStruct = (TapeGetDriveParameters)
                Marshal.PtrToStructure(ptr, typeof(TapeGetDriveParameters));

            Marshal.FreeHGlobal(ptr);//release memory

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// Initializes m_setDriveParamsStruct with the current parameters
        /// </summary>
        private void CopyCurrentParamsDriveHelper()
        {
            m_setDriveParamsStruct.ECC = m_getDriveParamsStruct.Value.ECC;
            m_setDriveParamsStruct.driveCompression = m_getDriveParamsStruct.Value.driveCompression;
            m_setDriveParamsStruct.DataPadding = m_getDriveParamsStruct.Value.DataPadding;
            m_setDriveParamsStruct.ReportSetmarks = m_getDriveParamsStruct.Value.ReportSetMarks;
            m_setDriveParamsStruct.EOTWarningZoneSize = m_getDriveParamsStruct.Value.EATWarningZone;
        }

        #endregion

    }

    #region ApplicationException
    /// <summary>
    /// Exception that will be thrown by tape
    /// Engine when one of WIN32 APIs terminates 
    /// with error code 
    /// </summary>
    public class TapeWinAPIException : ApplicationException
    {
        public TapeWinAPIException(string methodName, int win32ErroCode) :
            base(string.Format(
               "WIN32 API method failed : {0} failed with error code {1}",
               methodName,
               win32ErroCode
           ))
        { }
    }
    #endregion
}