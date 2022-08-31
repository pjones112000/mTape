using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Data.SQLite;
using System.Web;
using System.IO.Compression;

namespace mtape
{
    class Program
    {
        private static TapeWinAPI API = new TapeWinAPI();
        private static int m_Count = 0;
        private static String m_Filename = string.Empty;
        private static string m_Output = string.Empty;
        private static String newTapeScript = string.Empty;
        private static int m_VolumeNumber = 0;
        private static long m_NeededPosition = 0;
        private static long m_CurrentPosition = 0;
        private static long m_OverridePosition = -1;
        private static string m_LogFile = string.Empty;
        private static string m_ErrorLog = string.Empty;


        private const int NewTape = 0;
        private const int StartBackup = 1;
        private const int EndBackup = 2;
        private const int Alert = 3;
        private const int TapeNeeded = 4;
        private const int Clean = 5;
        private static string m_Message = string.Empty;
        private static string m_tapeDrive = string.Empty;


        private static Dictionary<String, String> m_GlobalVariables = new Dictionary<string, string>();
        private static Dictionary<String, List<String>> mFunctions = new Dictionary<string, List<string>>();
        private static Dictionary<String, List<String>> mFunctParams = new Dictionary<string, List<string>>();


        /// <summary>
        /// Entry point of the application
        /// </summary>
        /// <param name="string[] args"></param>
        static void Main(string[] args)
        {
            SetTape("tape0");

            int argMax = args.Length;
            for (int argCnt = 0; argCnt < argMax; argCnt++)
            {
                switch (args[argCnt])
                {
                    case "-h":
                        QuickHelp();
                        break;
                    case "--help":
                        DetailedHelp();
                        break;
                    case "-f":
                        if (API.TapeDrive != args[argCnt + 1])
                        {
                            System.Console.WriteLine("Changing tape device " + args[argCnt + 1]);
                        }
                        else
                        {
                            System.Console.WriteLine("Specified tape device is the same as the default.  Not changed.");
                        }
                        SetTape(args[argCnt + 1]); // next value must be the drive name
                        argCnt++; // increment counter to skip next value.
                        break;
                    case "-L":
                        m_LogFile = args[argCnt + 1];
                        argCnt++;
                        break;
                    case "-LE":
                        m_ErrorLog = args[argCnt + 1];
                        argCnt++;
                        break;
                    case "--version":
                        System.Reflection.Assembly assm = System.Reflection.Assembly.GetExecutingAssembly();

                        System.Console.WriteLine("mTape v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

                        var versionInfo = FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetEntryAssembly().Location);

                        var companyName = versionInfo.CompanyName;
                        var description = versionInfo.Comments;
                        System.Console.WriteLine(description);
                        var copyright = versionInfo.LegalCopyright;
                        System.Console.WriteLine(copyright);


                        return;
                        break;
                    case "--newtape":
                        newTapeScript = args[argCnt + 1];
                        System.Console.WriteLine("New tape script:" + newTapeScript);
                        argCnt++;
                        break;
                    default:
                        int tmpInt = 0;
                        if (int.TryParse(args[argCnt], out tmpInt)) // Is this value a number?  If so, then it is the count
                        {
                            m_Count = tmpInt;
                            argCnt = argMax;  // Exit before another number is found.
                        }
                        else
                        {
                            if (!command(args[argCnt]))
                            {
                                if (m_Filename == string.Empty)
                                {
                                    m_Filename = args[argCnt];
                                }
                                else
                                {
                                    m_Output = args[argCnt];
                                }
                            }
                        }
                        break;
                }
            }

            LogMessage("Main", "Tape device " + API.TapeDrive);

            for (int argCnt = 0; argCnt < argMax; argCnt++)
            {

                switch (args[argCnt])
                {
                    case "info":
                        Info();
                        // tape information
                        break;
                    case "status":
                        // Print status information about the tape unit.  (If the density code is "not translation" in the status output, this does not affect working of the tape drive.)
                        Status();
                        break;
                    case "lock":
                        // lock tape
                        Lock();
                        break;
                    case "unlock":
                        // unlock tape
                        Unlock();
                        break;
                    case "rewind":
                        // rewind tape
                        Rewind();
                        break;
                    case "fsf":
                        // forward space count files
                        FSF();
                        break;
                    case "bsf":
                        // back space count files
                        BSF();
                        break;
                    case "fss":
                        // forward space count setmarks
                        FSS();
                        break;
                    case "bss":
                        // Back space count setmarks
                        BSS();
                        break;
                    case "fsfm":
                        // forward count space file marks, then forward space 1 record.
                        FSFM();
                        break;
                    case "bsfm":
                        // Back space count files, then forward space one record.  Leaves the tape at the first block of the file that is count - 1
                        BSFM();
                        break;
                    case "asf":
                        // Tape is positioned at the the beginning of the count file.  Positioning is done by first rewinding the tape and then spacing forward over count filemarks.
                        ASF();
                        break;
                    case "fsr":
                        // Forward space count records.
                        FSR();
                        break;
                    case "bsr":
                        // Back space count records.
                        BSR();
                        break;
                    case "eod":
                    case "seod":
                        // Space ot end of valid data.  Used on streamer tape drives to append data to the logical end of tape.
                        EOD();
                        break;
                    case "offline":
                    case "rewoffl":
                    case "eject":
                        // Rewind the tape and, if applicable, unload the tape.
                        EJECT();
                        break;
                    case "retention":
                        // Rewind the tape, then wind it to the end of the reel, then rewind it again.
                        RETENTION();
                        break;
                    case "weof":
                    case "eof":
                        // Write count EOF marks at current position.
                        EOF();
                        break;
                    case "wset":
                        // (SCSI tapes) Write count setmarks at current position (only SCSI tape);
                        WSET();
                        break;
                    case "erase":
                        // Erase the tape.  Note that this is a long erase, which on modern (high-capacity) tapes can take many hours, and which usually can't be aborted.
                        ERASE();
                        break;
                    case "seeka":
                        // (SCSI tapes) Seek to the count block on the tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.  The block address should be obtained from a tell call earlier.
                        SEEKA();
                        break;
                    case "seekl":
                        // (SCSI tapes) Seek to the count block on the tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.  The block address should be obtained from a tell call earlier.
                        SEEKL();
                        break;
                    case "tell":
                        // (SCSI tapes) Tell the current block on tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.
                        TELL();
                        break;
                    case "setpartition":
                        // (SCSI tapes) Switch to the parition determined by count.  The default data partition of the tape is numbered zero.  Switching partition is available only if enabled for the device, the device supports multiple partitions, and the tape is formatted with multiple partitions.
                        SETPARTITION();
                        break;
                    case "partseek":
                        // (SCSI tapes) The tape position is set to block count in the partition given by the argument after count.  The default partition is zero.
                        PARTSEEK();
                        break;
                    case "mkpartition":
                        // (SCSI tapes) Format the tape with one (count is zero) or two partitions (count gives the size of the second partition in megabytes).  If the count
                        // is positive, it specifies the size of paritition 1.  From kernel version 4.6, if the count is negative, it specifies the size of partition 0.
                        // With older kernels, a negative arguemnt formats the tape with one parition.  The tape drive must be able to format partitioned tapes with
                        // initiator-specified partition size and partition support must be enabled for the drive.
                        MKPARTITION();
                        break;
                    case "load":
                        // (SCSI tapes) Send the load command to the tape drive.  The drives usually load the tape when a new cartridge is inserted.  The argument
                        // count can usually be omitted.  Some HP changers load tape n if the cunt 100000 + n is given (a special function in the Linux st driver).
                        Load();
                        break;
                    case "setblk":
                        // (SCSI tapes) Set the tape density code to count.  The proper codes to use with each drive should be looked up from the drive documentation.
                        SETBLK();
                        break;
                    case "densities":
                        // (SCSI tapes) Write explanation of some common density codes to standard output
                        DENSITIES();
                        break;
                    case "drvbuffer":
                        // (SCSI tapes) Set the tape drive buffer code to number.  The proper value for unbuffered operation is zero and "normal" buffered operation
                        // one.  The meanings of other values can be found in the drive documentation or, in the case of a SCSI-2 drive, from the SCSI-2 standard.
                        DRVBUFFER();
                        break;
                    case "compression":
                        // SCSI tapes) The compression within the drive can be switched on or off using the MTCOMPRESSION ioctl.  Note that this method is not
                        // supported by all drives implementing compression.  For instance, the Exabyte 8 mm drives use densitiy codes to select compression.
                        COMPRESSION();
                        break;
                    case "stoptions":
                        // (SCSI tapes) Set the driver options bits for the device to the defined values.  Allowed only for the superuser.  The bits can be set either by
                        //  ORing the option bits from the file /usr/include/linux/mtio.h to count, or by using the following keywords (as many keywords can be used on
                        // the same line as necessary, unambiguous abbreviatons allowed):
                        // buffer-writes    buffered writes enabled
                        // async-writes     asynchronous writes enabled
                        // read-ahead       read-ahead for fixed block size
                        // debug            debugging (if compiled into driver)
                        // two-fms          write two filemarks when file closed
                        // fast-eod         space directly to eod (and lose file number)
                        // no-wait          don't wait until rewind, etc. complete
                        // auto-lock        automatically lock/unlock drive door
                        // def-writes       the block size and density are for writes
                        // can-bsr          drive can space backwards as well
                        // no-blklimits     drive doesn't support read block limits
                        // can-parititions  drive can handle partitioned tapes
                        // scsi2logical     seek and tell use SCSI-2 logical block addresses instead of device dependent addresses
                        // sili             Set the SILI bit is when reading in variable block mode.  This may speed up reading blocks shorter than the read byte count.
                        //                  Set this option only if you know that the drive supports SILI and the HBA reliably returens transfer resicual byte counts.  Re-
                        //                  quires kernel version >= 2.6.26
                        // sysv             Enable the System V semantics
                        STOPTIONS();
                        break;
                    case "stsetoptions":
                        // (SCSI tapes) Set selected driver options bits.  The methods to specify the bits to set are given above in the description of stoptions.  Al-
                        // lowed only for the superuser.
                        STSETOPTIONS();
                        break;
                    case "stclearoptions":
                        // (SCSI tapes) Clear selected driver option bits.  The methods to specify the bits to clear are given above in the description of stoptions.  Al-
                        // lowed only for the superuser.
                        STCLEAROPTIONS();
                        break;
                    case "stshowoptions":
                        // (SCSI tapes) Print the currently enabled options for the device.  Requires kernel version >= 2.6.26 and sysfs must be mounted at /sys.
                        STSHOWOPTIONS();
                        break;
                    case "stwrthreshold":
                        // (SCSI tapes) The write threshold for the tape device is set to count kilobytes.  The value must be smaller than or equal to the driver buffer
                        // size.  Allowed only for the superuser.
                        STWRTHRESHOLD();
                        break;
                    case "defblksize":
                        // (SCSI tapes) Set the default block size of the device to count bytes.  The value -1 disables the default block size.  The block size set by
                        // setblk overrides the default until a new tape is inserted.  Allowed only for the superuser.
                        DEFBLKSIZE();
                        break;
                    case "defdensity":
                        // (SCIS tapes) Set the default density code.  The value -1 disables the default desnity.  The density set by setdensity overrides the default un-
                        // til a new tape is inserted.  Allowed only for the superuser.
                        DEFDENSITY();
                        break;
                    case "defdrvbuffer":
                        // (SCSI tapes) Set the default derive buffer code.  The value -1 disables the default drive buffer code.  The drive buffer code set by drvbuffer
                        // overrides the default until a new tape is inserted.  Allowed only for the superuser.
                        DEFDRVBUFFER();
                        break;
                    case "defcompression":
                        // (SCSI tapes) Set the default compression state.  The value -1 disables the default compression.  The compression state set by compression over-
                        // reides the default until a new tape is inserted.  Allowed only for the superuser.
                        DEFCOMPRESSION();
                        break;
                    case "sttimeout":
                        // Sets the normal timeout for the device.  The value is given in seconds.  Allowedo nly for the superuser.
                        STTIMEOUT();
                        break;
                    case "stlongtimeout":
                        // Sets the long timeout for the device.  The value is given in seconds.  Alloed only for the superuser.
                        STLONGTIMEOUT();
                        break;
                    case "stsetcln":
                        // Sets the cleaning request interpretation parameters.
                        STSETCLN();
                        break;
                    case "test":
                        NextTape();
                        Environment.Exit(1);
                        break;
                    case "locate":
                        Locate();
                        break;
                    case "init":
                        InitializeLibrary();
                        break;
                    case "read":
                        ReadTape();
                        break;
                    case "append":
                        AppendtoLibrary();
                        break;
                    case "resume":
                        ResumeWrite();
                        break;
                    case "write":
                        WriteTape();
                        break;
                    case "writelist":
                        WriteTapeFromList();
                        break;
                }
            }
            if (API.IsTapeOpen())
            {
                API.Close();
            }
        }



        /// <summary>
        /// Locates a file in the tape library and requests the volume that that file is located on, then positions the tape at that file.
        /// </summary>
        /// <param name="FilePath"></param>
        private static void LoadVolumeByPath(String FilePath)
        {
            bool retValue = false;
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Files' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    cmd = new SQLiteCommand("select Volume, Position from Files where Name = @name and Directory = @directory", conn);
                    System.IO.FileInfo FI = new FileInfo(FilePath);

                    cmd.Parameters.AddWithValue("@name", FI.Name);
                    cmd.Parameters.AddWithValue("@directory", FI.DirectoryName);
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    if (rdr2.HasRows)
                    {
                        while (rdr2.Read())
                        {
                            int neededVolume = rdr2.GetInt32(0);
                            long m_NeedPostion = rdr2.GetInt64(1);
                            m_VolumeNumber = neededVolume;
                            NextTape(false);
                            LogMessage("LoadVolumeByPath", "Seeking to the appropriate position on tape...");
                            SEEKL(m_NeededPosition);
                        }
                    }
                    else
                    {
                        LogError("DeleteFileInDatabase", "Specified file is does not exist in the database '" + FilePath + "'");
                        retValue = false;
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("DeleteFileInDatabase", "Files table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
        }

        /// <summary>
        /// Virtually identical to ContinueWrite....however uses a user supplied list of files to write to tape.
        /// </summary>
        private static void WriteTapeFromList()
        {
            if (newTapeScript != string.Empty)
            {
                int retCode = 0;
                //  Will need to return 0 (true) or -1 (false) to continue or not.
                runMTapeScript(StartBackup, newTapeScript, out retCode, "");
                addVolumeToDatabase();
            }

            StreamReader sr = new StreamReader(m_Filename);
            using (StreamWriter transWrite = new StreamWriter("transWrite.trn", true))
            {
                while (!sr.EndOfStream)
                {
                    // Read from transaction file
                    String tmpFile = sr.ReadLine();
                    LogMessage("WriteList", "Writing file (" + tmpFile + ")");
                    // Write the file to tape at current location
                    if (System.IO.File.Exists(tmpFile))
                    {
                        WriteFileHeader(tmpFile);
                        WriteFile(tmpFile);
                    }
                    // Record that it was written (last file might be duplicated...but, shouldn't be an issue)
                    transWrite.WriteLine(tmpFile);
                }
            }
            if (newTapeScript != string.Empty)
            {
                int retCode = 0;
                //  Will need to return 0 (true) or -1 (false) to continue or not.
                runMTapeScript(EndBackup, newTapeScript, out retCode, "");
            }

            LogMessage("WriteList", "Done");
        }


        private static void WriteFileHeader(String FileName)
        {
            System.IO.FileInfo FI = new FileInfo(FileName);
            String JSONString = string.Empty;

            if (System.IO.File.Exists("FileInfo.JSON"))
            {
                System.IO.File.Delete("FileInfo.JSON");
            }
            JSONString += "{\"Name\":\"" + HttpUtility.HtmlEncode(FI.Name) + "\",\"CreationTime\":\"" + HttpUtility.HtmlEncode(FI.CreationTime) + "\",\"ModifyTime\":\"" + HttpUtility.HtmlEncode(FI.LastWriteTime) + "\",\"FileSize\":" + HttpUtility.HtmlEncode(FI.Length.ToString()) + "}";
            using (System.IO.StreamWriter SW = System.IO.File.CreateText("FileInfo.JSON"))
            {
                SW.WriteLine(JSONString);
            }
            WriteFile("FileInfo.JSON", true,false);
            System.IO.File.Delete("FileInfo.JSON");
        }

        /// <summary>
        /// Called by ResumeWrite to perform the actual writing to tape.
        /// </summary>
        private static void ContinueWrite()
        {
            StreamReader sr = new StreamReader("transRead.trn");
            using (StreamWriter transWrite = new StreamWriter("transWrite.trn", true))
            {
                while (!sr.EndOfStream)
                {
                    // Read from transaction file
                    String tmpFile = sr.ReadLine();
                    LogMessage("ContinueTape", "Writing file (" + tmpFile + ")");
                    // Write the file to tape at current location
                    WriteFile(tmpFile);
                    // Record that it was written (last file might be duplicated...but, shouldn't be an issue)
                    transWrite.WriteLine(tmpFile);
                }
            }
        }

        /// <summary>
        /// When writing, mTape writes a transaction list of files that are scheduled to be written and a list of files actually written
        /// on Resume, a list of unwritten files is generated and used to continue writing...the last written file will be seeked and replaced
        /// to ensure that the EOF marker made it to tape.
        /// </summary>
        private static void ResumeWrite()
        {
            List<String> readList = new List<string>();
            List<String> writeList = new List<string>();

            if (System.IO.File.Exists("transRead.trn") && System.IO.File.Exists("transWrite.trn"))
            {
                foreach (String tmpItem in System.IO.File.ReadLines("transRead.trn"))
                {
                    readList.Add(tmpItem);
                }
                foreach (String tmpItem in System.IO.File.ReadLines("transWrite.trn"))
                {
                    writeList.Add(tmpItem);
                }

                LogMessage("ResumeWrite", "Creating list of unrwritten files...");
                String lastWritten = string.Empty;
                using (StreamWriter SW = new StreamWriter("transDiff.trn"))
                {
                    foreach (String tmpKey in readList)
                    {
                        if (!writeList.Contains(tmpKey))
                        {
                            SW.WriteLine(tmpKey);
                        }
                        else
                        {
                            lastWritten = tmpKey;
                        }
                    }
                }
                // Add the last written file to the transaction list....this is because we don't know if the location following this file is valid or not.

                String tmp = System.IO.File.ReadAllText("transDiff.trn");
                tmp = lastWritten + Environment.NewLine + tmp;
                System.IO.File.WriteAllText("transDiff.trn", tmp);

                // Remove the old transaction read file
                System.IO.File.Delete("transRead.trn");
                // Make the differences the new transaction read file.
                System.IO.File.Move("transDiff.trn", "transRead.trn");
                // Find the volume that the last file was written to from the tape library and ask for that volume, then seek to the last file's location.
                LoadVolumeByPath(lastWritten);
                LogMessage("ResumeWrite", "Continuing the write...");
                ContinueWrite();
            }
            Dummy("Complete");
        }
        private static void AppendtoLibrary()
        {
            Dummy("Append");
            uint retCode = 0;
            if (!API.SeekToEOD(ref retCode))
            {
                LogError("Append", "Error seeking to the end of data Error:" + retCode);
            }
            else
            {
                LogMessage("Append", "Appending to the volume...");
                WriteTape();
            }
        }

        /// <summary>
        /// Is the argument a command?
        /// </summary>
        /// <param name="tmp"></param>
        /// <returns>bool Yes/No</returns>
        private static bool command(String tmp)
        {
            bool retValue = false;
            switch (tmp.ToLower())
            {
                case "weof":
                case "wset":
                case "eof":
                case "fsf":
                case "fsfm":
                case "bsf":
                case "bsfm":
                case "fsr":
                case "bsr":
                case "fss":
                case "bss":
                case "rewind":
                case "offline":
                case "rewoffl":
                case "eject":
                case "retention":
                case "eod":
                case "seod":
                case "seeka":
                case "seekl":
                case "tell":
                case "status":
                case "erase":
                case "setblk":
                case "lock":
                case "unlock":
                case "load":
                case "compression":
                case "setdensity":
                case "drvbuffer":
                case "stwrthreshold":
                case "stoptions":
                case "stsetoptions":
                case "stclearoptions":
                case "defblksize":
                case "defdensity":
                case "defdrvbuffer":
                case "defcompression":
                case "stsetcln":
                case "sttimeout":
                case "stlongtimeout":
                case "densities":
                case "setpartition":
                case "mkpartition":
                case "partseek":
                case "asf":
                case "stshowoptions":
                case "read":
                case "write":
                case "info":
                case "locate":
                case "init":
                case "delete":
                case "resume":
                case "append":
                case "writelist":
                    return true;
                    break;
            }
            return retValue;
        }

        /// <summary>
        /// Initializes the tape library (i.e. recreates the database).
        /// </summary>
        private static void InitializeLibrary()
        {
            SQLiteConnection retValue = null;
            bool createTables = false;
            try
            {
                String dbName = Path.Combine(Directory.GetCurrentDirectory(), "tapeLibrary.db");
                if (System.IO.File.Exists(dbName))
                {
                    System.IO.File.Delete(dbName);
                    createTables = true;
                }

                string cs = @"URI=file:" + dbName;

                retValue = new SQLiteConnection(cs);
                if (createTables)
                {
                    SQLiteConnection.CreateFile(dbName);
                }

                retValue.Open();
                if (createTables)
                {
                    SQLiteCommand cmd = new SQLiteCommand(retValue);
                    cmd.CommandText = @"CREATE Table Files(id INTEGER PRIMARY KEY, name TEXT,directory TEXT, Volume INTEGER, FileDate TEXT, Size INTEGER, Position INTEGER, Deleted INTEGER)";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = @"CREATE Table Volumes(id INTEGER PRIMARY KEY, name TEXT)";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                LogError("InitializeLibrary", "Failed to initialize the tape library Error:" + ex.Message);
            }
            LogMessage("InitializeLibrary", "Library has been initialized.");
        }

        /// <summary>
        /// Read a file from tape
        /// </summary>
        /// <param name="string Filename to write to"></param>
        /// <returns>bool success</returns>
        private static bool ReadFile(String Filename)
        {
            bool retValue = true;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }

            if (API.IsTapeOpen())
            {
                int MAX_BUFFER = API.MaximumBlockSizeDrive; //1MB
                int MIN_BUFFER = API.MinimumBlockSizeDrive;
                byte[] buffer = new byte[MAX_BUFFER];
                uint bytesRead = 0;
                int noOfFiles = 0;
                uint retCode = 0;
                bool Continue = true;
                LogMessage("ReadFile", "Reading from tape to file " + Filename);
                System.Console.WriteLine("Restoring to " + Filename + " if this is not correct, abort now!!!");
                System.Console.ReadLine();

                using (FileStream fs = File.Open(Filename, FileMode.Create, FileAccess.Write))
                using (BufferedStream bs = new BufferedStream(fs))
                {
                    while (Continue)
                    {
                        while (API.Read(ref buffer, (uint)MAX_BUFFER, ref bytesRead, ref retCode))
                        {
                            bs.Write(buffer, 0, (int)bytesRead);
                        }
                        bs.Flush();
                        fs.Flush();
                        Continue = false;
                        if (retCode != 0)
                        {
                            switch (retCode)
                            {
                                case 0234:  //more data is available
                                            //update error log and status
                                    LogError("ReadFile", "aborted due to SCSI Controller problem");
                                    break;
                                case 1106:  //incorrect block size
                                            //update error log and status
                                    LogError("ReadFile", "aborted due to invalid block size");
                                    break;
                                case 1101:  //reached file mark
                                    LogError("ReadFile", "File mark detected.");
                                    break;
                                case 1104:  //EOD reached
                                    LogError("ReadFile", "EOD detected");
                                    break;
                                case 1129:  // Physical EOT
                                case 1100:  // EOT marker reached
                                    LogError("ReadFile", "EOT detected.");
                                    Continue = NextTape();
                                    break;
                                default:    //any other errors
                                    LogError("ReadFile", "aborted with error code: " + retCode.ToString());
                                    break;
                            }
                        }
                        else
                        {
                            Continue = false;
                        }
                        LogMessage("ReadFile", "End of read");
                    }
                }
                LogMessage("ReadFile", "File was read from tape");
            }
            else
            {
                LogError("ReadFile", "Unable to open tape drive");
            }
            return retValue;
        }

        /// <summary>
        /// Command to read a single file from tape drive
        /// </summary>
        private static void ReadTape()
        {
            Dummy("Read");
            LogMessage("Read", "Will eventually read a file from the mounted volume from the current location.");
            ReadFile(m_Filename);
        }

        /// <summary>
        /// Calculates the number of chunks required, based on tape drive's buffer size.
        /// </summary>
        /// <param name="string filename"></param>
        /// <returns>int ChunkCount</returns>        
        private static int FileChunkCount(String FileName)
        {
            try
            {
                System.IO.FileInfo FI = new FileInfo(FileName);
                return (int)((FI.Length / API.MaximumBlockSizeDrive) + 1);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        /// <summary>
        /// Capture stdError output from external batch or executable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void p_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            Process p = sender as Process;
            if (p == null)
                return;
            Console.WriteLine(e.Data);
        }

        /// <summary>
        /// Captures standard output from executed batch/executable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void p_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Process p = sender as Process;
            if (p == null)
                return;
            Console.WriteLine(e.Data);
        }

        /// <summary>
        /// Is the script a batch file
        /// </summary>
        /// <returns>bool Is a batch file</returns>
        private static bool isScriptBatch()
        {
            bool retValue = false;
            System.IO.FileInfo FI = new FileInfo(newTapeScript);
            if (FI.Extension == ".bat")
            {
                retValue = true;
            }
            return retValue;
        }

        /// <summary>
        /// Determines if the script is an mTape script or not
        /// </summary>
        /// <returns>bool Is an mTape script</returns>
        private static bool isScriptMtape()
        {
            bool retValue = false;
            string[] text = System.IO.File.ReadAllLines(newTapeScript);
            if (text[0].Trim().ToLower() == "#!mtape")
            {
                retValue = true;
            }
            return retValue;
        }

        /// <summary>
        /// Determines if the script is an mTape script or not
        /// </summary>
        /// <returns>bool Is an mTape script</returns>
        private static bool isScriptMtape2()
        {
            bool retValue = false;
            string[] text = System.IO.File.ReadAllLines(newTapeScript);
            if (text[0].Trim().ToLower() == "#!mtape2")
            {
                retValue = true;
            }
            return retValue;
        }


        /// <summary>
        /// Keeps checking to see if the tape drive is read (i.e. tape loaded and ready to write)
        /// </summary>
        /// <returns>bool Success</returns>
        private static bool DetectNewTape()
        {
            bool retValue = false;
            bool tapeDetected = false;

            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                LogMessage("DetectNewTape", "Looking for new tape.");
                while (!tapeDetected)
                {
                    try
                    {
                        uint retCode = 0;
                        if (getCurrentPosition() != -1)
                        {
                            tapeDetected = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        int errCode = (int)(uint)Marshal.GetLastWin32Error();
                        if (errCode != TapeWinAPI.ERROR_NO_MEDIA_IN_DRIVE)
                        {
                            LogError("DetectNewTape", "Error detecting new tape Error:" + errCode);
                        }
                    }
                    System.Threading.Thread.Sleep(500);
                }
            }
            return retValue;
        }

        /// <summary>
        /// Removes double quotes from values
        /// </summary>
        /// <param name="Dictionary&lt;String,String&gt;source"></param>
        private static void cleanVariables(ref Dictionary<String, String> source)
        {
            Dictionary<String, String> tmp = new Dictionary<string, string>();

            foreach (String tmpKey in source.Keys)
            {
                String tmpString = source[tmpKey].Replace("\"", "");
                tmp.Add(tmpKey, tmpString);
            }
            source.Clear();
            foreach (String tmpKey in tmp.Keys)
            {
                source.Add(tmpKey, tmp[tmpKey]);
            }
        }

        /// <summary>
        /// Sends an SMTP email
        /// </summary>
        /// <param name="Dictionary&lt;String,String&gt;mVariables"></param>
        /// <returns>bool Success</returns>
        private static bool SendMail(Dictionary<String, String> mVariables)
        {
            bool RetValue = false;

            try
            {
                cleanVariables(ref mVariables);

                var smtpClient = new SmtpClient(mVariables["MailHost"])
                {
                    Port = int.Parse(mVariables["MailPort"]),
                    Credentials = new System.Net.NetworkCredential(mVariables["MailUser"], mVariables["MailPassword"]),
                    EnableSsl = true,
                };
                String tmpMessage = mVariables["MailMessage"];
                tmpMessage = tmpMessage.Replace("$Volume", "VOL_" + zeroFill(m_VolumeNumber,4)).Replace("$VOLUME", zeroFill(m_VolumeNumber,4)).Replace("$REASON", mVariables["$REASON"]);
                smtpClient.Send(mVariables["MailSender"], mVariables["MailRecipient"], mVariables["MailSubject"], tmpMessage);
                RetValue = true;
            }
            catch (Exception ex)
            {
                RetValue = false;
            }
            return RetValue;
        }

        /// <summary>
        /// Interpret an mTape script
        /// </summary>
        /// <returns>bool Success</returns>
        private static bool runMTapeScript(bool newTape = true)
        {
            bool retValue = false;
            int retCode = 0;
            string[] text = System.IO.File.ReadAllLines(newTapeScript);
            String regExPattern = "[A-Z$\\(\\)]+|(\\\"([^\\\"])*\\\")";

            Regex rg = new Regex(regExPattern, RegexOptions.IgnoreCase);

            Dictionary<String, String> mVariables = new Dictionary<string, string>();
            foreach (String tmpText in text)
            {
                if (!tmpText.StartsWith("#") && tmpText.Trim() != string.Empty)
                {
                    MatchCollection matches = rg.Matches(tmpText);
                    if (matches.Count > 0)
                    {
                        switch (matches[0].Value.ToLower())
                        {
                            case "eject()":
                                // Commented out for testing purposes
                                //                                EJECT();
                                break;
                            case "keypress()":
                                System.Console.ReadKey();
                                break;
                            case "detectnewtape()":
                                LogMessage("runMTapeScript", "Detecting new tape in drive.");
                                if (!DetectNewTape())
                                {
                                    return false;
                                }
                                // tape detection 
                                break;
                            case "sendmail()":
                                LogMessage("runMTapeScript", "Sending email");
                                if (!mVariables.ContainsKey("Reason"))
                                {
                                    mVariables.Add("Reason", mVariables["$REASON"]);
                                }
                                else
                                {
                                    mVariables["Reason"] = mVariables["$REASON"];
                                }
                                if (!SendMail(mVariables))
                                {
                                    return false;
                                }
                                // send mail
                                break;
                            case "print":
                                if (newTape)
                                {
                                    LogMessage(newTapeScript, matches[1].Value.Replace("\"", "").Replace("$VOLUME", "VOL_" + zeroFill(m_VolumeNumber,4)).Replace("$REASON", "Newtape"));
                                }
                                else
                                {
                                    LogMessage(newTapeScript, matches[1].Value.Replace("\"", "").Replace("$VOLUME", "VOL_" + zeroFill(m_VolumeNumber,4)).Replace("$REASON", "Needed"));
                                }
                                break;
                            case "call":
                                // Calls external application or script
                                if (!RunExternal(out retCode, "runMTapeScript"))
                                {
                                    LogMessage("runMTapeScript", "External command reported an error :" + retCode);
                                    return false;
                                }
                                break;
                            default:
                                if (tmpText.Contains("="))
                                {
                                    if (!mVariables.ContainsKey(matches[0].Value))
                                    {
                                        mVariables.Add(matches[0].Value, matches[1].Value);
                                    }
                                    else
                                    {
                                        mVariables[matches[0].Value] = matches[1].Value;
                                    }
                                }
                                else
                                {
                                    LogMessage(newTapeScript, "Unrecognized command line " + tmpText);
                                }
                                break;
                        }
                    }
                    else
                    {
                        LogError("runMTapeScript", "Unmatched command line '" + tmpText + "'");
                    }
                }
            }
            return retValue;
        }

        /// <summary>
        /// Executes an external batch file or executable
        /// </summary>
        /// <param name="ref int RetCode"></param>
        /// <param name="String Method"></param>
        /// <returns>bool Success</returns>
        private static bool RunExternal(out int RetCode, String Method = "", bool newTape = false)
        {
            bool retValue = false;
            try
            {
                using (Process p = new Process())
                {
                    // set start info
                    p.StartInfo = new ProcessStartInfo("cmd.exe")
                    {
                        RedirectStandardInput = true,
                        UseShellExecute = false,
                        WorkingDirectory = @"d:\"
                    };
                    // event handlers for output & error
                    p.OutputDataReceived += p_OutputDataReceived;
                    p.ErrorDataReceived += p_ErrorDataReceived;

                    // start process
                    p.Start();
                    // send command to its input
                    if (newTape)
                    {
                        p.StandardInput.Write(newTapeScript + " VOL_" + zeroFill(m_VolumeNumber,4) + " New" + p.StandardInput.NewLine);
                    }
                    else
                    {
                        p.StandardInput.Write(newTapeScript + " VOL_" + zeroFill(m_VolumeNumber,4) + " Needed" + p.StandardInput.NewLine);
                    }
                    //wait
                    p.WaitForExit();
                    retValue = p.ExitCode == 0;
                    RetCode = p.ExitCode;
                    return retValue;
                }
            }
            catch (Exception ex)
            {
                LogError("CallScript", "Error executing script " + ex.Message);
                RetCode = ex.HResult;
                return false;
            }
        }


        /// <summary>
        /// Logs an error.
        /// </summary>
        /// <param name="String Command">The command issuing the error</param>
        /// <param name="String Error" type="string">The comment to add to the error log</param>
        static private void LogError(String Command, String Error)
        {
            if (m_ErrorLog != string.Empty)
            {
                if (!System.IO.File.Exists(m_ErrorLog))
                {
                    var fs = System.IO.File.Create(m_ErrorLog);
                    fs.Close();
                }
                using (System.IO.StreamWriter SW = System.IO.File.AppendText(m_ErrorLog))
                {
                    SW.WriteLine("[" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + " - " + Command + "] " + Error);
                }
                System.Console.WriteLine("[" + Command + "] " + Error);
            }
        }

        /// <summary>
        /// Logs a message.
        /// </summary>
        /// <param name="String Command"></param>
        /// <param name="String Message" type="string"></param>       
        static private void LogMessage(String Command, String Message)
        {
            if (m_LogFile != string.Empty)
            {
                if (!System.IO.File.Exists(m_LogFile))
                {
                    var fs = System.IO.File.Create(m_LogFile);
                    fs.Close();
                }
                using (System.IO.StreamWriter SW = System.IO.File.AppendText(m_LogFile))
                {
                    SW.WriteLine("[" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + " - " + Command + "] " + Message);
                }
            }
            System.Console.WriteLine("[" + Command + "] " + Message);
        }


        /// <summary>
        /// Executes an mtape script or batch file
        /// </summary>
        /// <param name="Ref int RetCode"></param>
        /// <returns>bool Success</returns>
        private static bool CallScript(out int RetCode, bool newTape = false, int ReasonCode = 0)
        {
            bool retValue = false;
            int retCode = 0;
            if (isScriptBatch())
            {
                retValue = RunExternal(out RetCode, "CallScript", newTape);
            }
            else
            {
                if (isScriptMtape())
                {
                    LogMessage("CallScript", "MTape Script v1");
                    runMTapeScript(newTape);
                    retValue = true;
                    // Call my own scripting language
                }
                else
                {
                    if (isScriptMtape2())
                    {
                        LogMessage("CallScript", "MTape Script v2");
                        if (newTape)
                        {
                            runMTapeScript(NewTape, newTapeScript, out retCode, "");
                        }
                        else
                        {
                            runMTapeScript(ReasonCode, newTapeScript, out retCode, "");
                        }
                    }
                    retValue = true;
                }
            }
            RetCode = retCode;
            return retValue;
        }


        /// <summary>
        /// Executed if a new tape is required....optional user script can be executed as well.
        /// </summary>
        /// <returns>bool success</returns>
        private static bool NextTape(bool newTape = true)
        {
            bool retValue = false;

            LogMessage("NextTape", "Starting NextTape");
            // Ejecting tape
            if (newTape)
            {
                LogMessage("NextTape", "newTape set");
                EJECT();
            }

            if (newTapeScript != string.Empty)
            {
                int retCode = 0;
                //  Will need to return 0 (true) or -1 (false) to continue or not.
                LogMessage("NextTape", "Calling script");
                if (!CallScript(out retCode, newTape))
                {
                    LogError("NextTape", "Error calling script " + retCode);
                }
            }
            else
            {
                if (newTape)
                {
                    LogMessage("Next Tape", "Tape is full...please insert the next volume (VOL_" + zeroFill(m_VolumeNumber,4) + ") and press any key.");
                }
                else
                {
                    LogMessage("Next Tape", "Tape needed for restore...please insert volume (VOL_" + zeroFill(m_VolumeNumber,4) + ") and press any key.");
                }
                System.Console.ReadKey();
                // Load the new Tape.
                if (newTape)
                {
                    Load();
                    m_VolumeNumber++;
                }
                retValue = true;
            }
            return retValue;
        }

        /// <summary>
        /// Connects to the SQLite database
        /// </summary>
        /// <returns>SQLiteConnection</returns>
        private static SQLiteConnection ConnectDatabase()
        {
            SQLiteConnection retValue = null;
            bool createTables = false;
            String dbName = Path.Combine(Directory.GetCurrentDirectory(), "tapeLibrary.db");
            if (!System.IO.File.Exists(dbName))
            {
                createTables = true;
            }

            string cs = @"URI=file:" + dbName;

            retValue = new SQLiteConnection(cs);
            if (createTables)
            {
                SQLiteConnection.CreateFile(dbName);
            }

            retValue.Open();
            if (createTables)
            {
                SQLiteCommand cmd = new SQLiteCommand(retValue);
                cmd.CommandText = @"CREATE Table Files(id INTEGER PRIMARY KEY, name TEXT,directory TEXT, Volume INTEGER, FileDate TEXT, Size INTEGER, Position INTEGER, Deleted INTEGER)";
                cmd.ExecuteNonQuery();

                cmd.CommandText = @"CREATE Table Volumes(id INTEGER PRIMARY KEY, name TEXT)";
                cmd.ExecuteNonQuery();
            }

            return retValue;
        }

        /// <summary>
        /// Marks the specified file as deleted in the tape library
        /// </summary>
        /// <param name="String Filename"></param>
        /// <returns>bool Success</returns>
        private static bool DeleteFileInDatabases(String Filename)
        {
            bool retValue = false;
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Files' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    cmd = new SQLiteCommand("select * from Files where Name = @name", conn);
                    cmd.Parameters.AddWithValue("@name", Filename);
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    if (rdr2.HasRows)
                    {
                        cmd = new SQLiteCommand("update Files set Deleted = @deleted", conn);
                        cmd.Parameters.AddWithValue("@deleted", true);
                        cmd.ExecuteNonQuery();
                        retValue = true;
                    }
                    else
                    {
                        LogError("DeleteFileInDatabase", "Specified file is does not exist in the database '" + Filename + "'");
                        retValue = false;
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("DeleteFileInDatabase", "Files table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
            return retValue;
        }

        private static String zeroFill(int NumberToFill, int Width)
        {
            string retValue = NumberToFill.ToString();
            while(retValue.Length < Width)
            {
                retValue = "0" + retValue;
            }

            return retValue;
        }

        /// <summary>
        /// Adds the volume to the tape library
        /// </summary>
        /// <returns>bool Success</returns>
        private static bool addVolumeToDatabase()
        {
            bool retValue = false;
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Volumes' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    cmd = new SQLiteCommand("select * from Volumes where Name = @name", conn);
                    cmd.Parameters.AddWithValue("@name", "VOL_" + zeroFill(m_VolumeNumber,4));
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    if (!rdr2.HasRows)
                    {
                        cmd = new SQLiteCommand("insert into Volumes (Name) values(@name)", conn);
                        cmd.Parameters.AddWithValue("@name", "VOL_" + zeroFill(m_VolumeNumber,4));
                        cmd.ExecuteNonQuery();
                        retValue = true;
                    }
                    else
                    {
                        LogError("addVolumeToDatabase", "Specified volume is already in the database '" + "VOL_" + zeroFill(m_VolumeNumber, 4) + "'");
                        retValue = false;
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("addVolumeToDatabase", "Volumes table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
            return retValue;
        }

        /// <summary>
        /// Add the specified file into the tape library
        /// </summary>
        /// <param name="String Filename"></param>
        /// <returns>bool Success</returns>
        private static bool AddFileToDatabase(String Filename)
        {
            bool retValue = false;
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Files' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    cmd = new SQLiteCommand("select * from Files where Name = @name", conn);
                    cmd.Parameters.AddWithValue("@name", Filename);
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    if (!rdr2.HasRows)
                    {
                        cmd = new SQLiteCommand("insert into Files (Name,Directory,Volume,FileDate,Size,Position,Deleted) values(@name, @directory, @volume, @date, @size,@position, @deleted)", conn);
                        System.IO.FileInfo FI = new FileInfo(Filename);

                        cmd.Parameters.AddWithValue("@name", FI.Name);
                        cmd.Parameters.AddWithValue("@directory", FI.DirectoryName);
                        cmd.Parameters.AddWithValue("@volume", m_VolumeNumber);
                        cmd.Parameters.AddWithValue("@date", FI.CreationTime);
                        cmd.Parameters.AddWithValue("@size", FI.Length);
                        cmd.Parameters.AddWithValue("@position", m_CurrentPosition);
                        cmd.Parameters.AddWithValue("@deleted", false);
                        cmd.ExecuteNonQuery();
                        retValue = true;
                    }
                    else
                    {
                        LogError("AddFileToDatabase", "Specified file is already in the database '" + Filename + "'");
                        retValue = false;
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("AddFileToDatabase", "Files table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
            return retValue;
        }

        /// <summary>
        /// Writes a single file to the tape drive
        /// </summary>
        /// <param name="string filename"></param>
        /// <returns>bool Success</returns>        
        private static bool WriteFile(String Filename, Boolean SupressMessages = false, Boolean AddToDatabase = true)
        {
            bool retValue = true;
            bool tapeOpen = false;
            int tmpCode = 0;
            m_CurrentPosition = getCurrentPosition();

            // If the override psoition != -1 then use it.
            if(m_OverridePosition != -1)
            {
                m_CurrentPosition = m_OverridePosition;
            }

            // If not adding to the database (i.e. JSO file header) capture the override position.
            // The intent is that the database record points to the header and not the physical file.
            if(!AddToDatabase)
            {
                m_OverridePosition = m_CurrentPosition;
            }
            else
            {
                m_OverridePosition = -1;
            }
            if (m_Count != 0)
            {
                if (!SupressMessages)
                {
                    LogMessage("WriteFile", "Seeking to the specified location before writing.");
                }
                SEEKL((long)m_Count);
            }
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    if (!SupressMessages)
                    {
                        LogMessage("WriteFile", "Tape drive:" + @"\\.\" + API.TapeDrive);
                        LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                    }
                }
            }
            if (API.IsTapeOpen())
            {
                int MAX_BUFFER = API.MaximumBlockSizeDrive; //1MB
                int MIN_BUFFER = API.MinimumBlockSizeDrive;

                byte[] buffer = new byte[MAX_BUFFER];
                int bytesRead;
                int noOfFiles = 0;

                int chunksRequired = FileChunkCount(Filename);


                using (FileStream fs = File.Open(Filename, FileMode.Open, FileAccess.Read))
                using (BufferedStream bs = new BufferedStream(fs))
                {
                    while ((bytesRead = bs.Read(buffer, 0, MAX_BUFFER)) != 0) //reading 1mb chunks at a time
                    {
                        noOfFiles++;
                        uint bytesWritten = 0;
                        uint retCode = 0;
                        if (!API.Write(ref buffer, (uint)bytesRead, ref bytesWritten, ref retCode))
                        {
                            switch (retCode)
                            {
                                case 1112:  // Not media in drive
                                    LogError("WriteFile", "There is no tape in the drive.");
                                    Environment.Exit(1);
                                    break;
                                case 0234:  //more data is available
                                            //update error log and status
                                    LogError("WriteFile", "aborted due to SCSI Controller problem");
                                    runMTapeScript(Alert, newTapeScript, out tmpCode, "Aborted due to SCSI Controller problem.");
                                    break;
                                case 1106:  //incorrect block size
                                            //update error log and status
                                    LogError("WriteFile", "aborted due to invalid block size");
                                    runMTapeScript(Alert, newTapeScript, out tmpCode, "Aborted due to invalid block size.");
                                    break;
                                case 1129:  // Physical end of tape
                                case 1100:  // EOT Marker reached
                                    LogError("WriteFile", "EOT detected.");
                                    LogError("WriteFile", "Calling NextTape");
                                    if (NextTape())
                                    {
                                        LogError("WriteFile", "Adding volume");
                                        m_VolumeNumber++;
                                        addVolumeToDatabase();
                                        continue;
                                    }
                                    break;
                                case 1156:  // Device Requires Cleaning
                                    LogMessage("WriteFile", "Tape Drive Requires Cleaning.");
                                    runMTapeScript(Clean, newTapeScript, out tmpCode, "Tape drive requires cleaning.");
                                    break;
                                case 1110:  // Media Changed (normal if new tape was inserted.
                                    continue;
                                case 0:
                                    break;
                                default:    //any other errors
                                    LogError("WriteFile", "aborted with error code: " + retCode.ToString());
                                    runMTapeScript(Alert, newTapeScript, out tmpCode, "Aborted Error Code:" + retCode.ToString());
                                    break;
                            }
                        }

                    }
                }
                // Write End of File marker
                EOF();
                if (AddToDatabase)
                {
                    AddFileToDatabase(Filename);
                }
            }
            else
            {
                LogError("WriteFile", "Unable to open tape drive");
            }
            return retValue;
        }

        /// <summary>
        /// Writes a single file to the tape drive
        /// </summary>
        /// <param name="string filename"></param>
        /// <returns>bool Success</returns>        
        private static bool WriteFileWithCompression(String Filename, Boolean SupressMessages = false, Boolean AddToDatabase = true)
        {
            bool retValue = true;
            bool tapeOpen = false;
            

            m_CurrentPosition = getCurrentPosition();

            // If the override psoition != -1 then use it.
            if (m_OverridePosition != -1)
            {
                m_CurrentPosition = m_OverridePosition;
            }

            // If not adding to the database (i.e. JSO file header) capture the override position.
            // The intent is that the database record points to the header and not the physical file.
            if (!AddToDatabase)
            {
                m_OverridePosition = m_CurrentPosition;
            }
            else
            {
                m_OverridePosition = -1;
            }
            if (m_Count != 0)
            {
                if (!SupressMessages)
                {
                    LogMessage("WriteFile", "Seeking to the specified location before writing.");
                }
                SEEKL((long)m_Count);
            }
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    if (!SupressMessages)
                    {
                        LogMessage("WriteFile", "Tape drive:" + @"\\.\" + API.TapeDrive);
                        LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                    }
                }
            }
            if (API.IsTapeOpen())
            {
                int MAX_BUFFER = API.MaximumBlockSizeDrive; //1MB
                int MIN_BUFFER = API.MinimumBlockSizeDrive;
                int tmpOut = 0;
                byte[] buffer = new byte[MAX_BUFFER];
                int bytesRead;
                int noOfFiles = 0;

                int chunksRequired = FileChunkCount(Filename);


                using (FileStream fs = File.Open(Filename, FileMode.Open, FileAccess.Read))
                using (MemoryStream CompressedFileStream = new MemoryStream())
                using (var Compressor = new GZipStream(CompressedFileStream,CompressionMode.Compress, true))
                using (BufferedStream bs = new BufferedStream(fs))
                {
                    fs.CopyTo(Compressor);
                    while ((bytesRead = Compressor.Read(buffer, 0, MAX_BUFFER)) != 0) //reading 1mb chunks at a time

//                        while ((bytesRead = bs.Read(buffer, 0, MAX_BUFFER)) != 0) //reading 1mb chunks at a time
                    {
                        noOfFiles++;
                        uint bytesWritten = 0;
                        uint retCode = 0;
                        if (!API.Write(ref buffer, (uint)bytesRead, ref bytesWritten, ref retCode))
                        {
                            switch (retCode)
                            {
                                case 1112:  // Not media in drive
                                    LogError("WriteFile", "There is no tape in the drive.");
                                    Environment.Exit(1);
                                    break;
                                case 0234:  //more data is available
                                            //update error log and status
                                    LogError("WriteFile", "aborted due to SCSI Controller problem");
                                    runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted due to SCSI Controller problem.");

                                    break;
                                case 1106:  //incorrect block size
                                            //update error log and status
                                    LogError("WriteFile", "aborted due to invalid block size");
                                    runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted due to invalid block size.");
                                    break;
                                case 1129:  // Physical end of tape
                                case 1100:  // EOT Marker reached
                                    LogError("WriteFile", "EOT detected.");
                                    if (NextTape())
                                    {
                                        m_VolumeNumber++;
                                        addVolumeToDatabase();
                                        continue;
                                    }
                                    break;
                                case 1156:  // Device Requires Cleaning
                                    LogMessage("WriteFile", "Tape Drive Requires Cleaning.");
                                    runMTapeScript(Clean, newTapeScript, out tmpOut, "Tape drive requires cleaning.");
                                    break;

                                case 1110:  // Media Changed (normal if new tape was inserted.
                                    continue;
                                case 0:
                                    break;
                                default:    //any other errors
                                    LogError("WriteFile", "aborted with error code: " + retCode.ToString());
                                    runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted Error Code:" + retCode.ToString());
                                    break;
                            }
                        }

                    }
                }
                // Write End of File marker
                EOF();
                if (AddToDatabase)
                {
                    AddFileToDatabase(Filename);
                }
            }
            else
            {
                LogError("WriteFile", "Unable to open tape drive");
            }
            return retValue;
        }


        /// <summary>
        /// retrieves a list of files from the specified folder
        /// I am fully aware that this is not the most effecient or fastest method; however,
        /// the built-in .net method was throwing errors when accessing inaccessible files which would truncate the list 
        /// and not include files that were accessible and I could not find a workaround.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static string[] getFiles(String path)
        {
            string[] retValue = new string[0];
            try
            {
                retValue = System.IO.Directory.GetFiles(path);

                StreamWriter transRead = new StreamWriter("transRead.trn", true);
                try
                {
                    foreach (String tmpFile in retValue)
                    {
                        transRead.WriteLine(tmpFile);
                    }
                }
                catch (Exception ex)
                {
                    LogError("getFiles", "Error getting files in directory (" + path + ") error:" + ex.Message);
                }
                transRead.Flush();
                transRead.Close();
            }
            catch (Exception ex)
            {
                LogError("getFiles", "Error reading directory (" + path + ") error:" + ex.Message);
            }
            return retValue;
        }

        /// <summary>
        /// Ensures that the string is always of the maxWidth, no more and no less.
        /// </summary>
        /// <param name="inString">String to truncate</param>
        /// <param name="maxWidth">The desire length</param>
        /// <returns>Altered string</returns>
        private static string truncateRight(String inString, int maxWidth)
        {
            if (inString.Length > maxWidth)
            {
                return inString.Substring(0, maxWidth);
            }
            else
            {
                return padRight(inString, maxWidth);
            }
        }

        /// <summary>
        /// Recursively collects all the files in the directories starting at the specified folder.
        /// I am fully aware that this is not the most efficient or fastest method; however, the built-in .Net methods were throwing errors
        /// whenever there was an inaccessible folder/file that would abort the collection short of a fullly accessible list.
        /// </summary>
        /// <param name="startDir"></param>
        /// <returns></returns>
        private static String[] getDirs(String startDir)
        {
            string[] retValue = new string[0];
            try
            {
                int origRow = Console.CursorTop;
                int origCol = Console.CursorLeft;

                System.Console.Write(truncateRight(startDir, Console.WindowWidth - 5));
                System.Console.SetCursorPosition(origCol, origRow);
                string[] tmpDirs = new string[0];
                try
                {
                    string[] wrkFiles = getFiles(startDir);

                    int oldSize = retValue.Length;
                    Array.Resize(ref retValue, retValue.Length + wrkFiles.Length);
                    wrkFiles.CopyTo(retValue, oldSize);
                }
                catch (Exception ex)
                {
                    LogError("getDirs", "Error reading files from directory (" + startDir + ")");
                }
                tmpDirs = System.IO.Directory.GetDirectories(startDir);
                foreach (String tmpDir in tmpDirs)
                {
                    string[] wrkValue = getDirs(tmpDir);
                    int oldSize = retValue.Length;
                    Array.Resize(ref retValue, retValue.Length + wrkValue.Length);
                    wrkValue.CopyTo(retValue, oldSize);
                }
            }
            catch (Exception ex)
            {
                LogError("getDirs", "Error reading directory (" + startDir + ") error:" + ex.Message);
            }
            return retValue;
        }


        /// <summary>
        /// Command to write a single file to the tape drive
        /// </summary>
        private static void WriteTape()
        {
            FileAttributes attr = File.GetAttributes(m_Filename);
            System.IO.FileInfo FI = new FileInfo(m_Filename);

            if (newTapeScript != string.Empty)
            {
                int retCode = 0;
                //  Will need to return 0 (true) or -1 (false) to continue or not.
                int tmpOut = 0;
                runMTapeScript(StartBackup,newTapeScript, out tmpOut, "");
                addVolumeToDatabase();
            }

            System.Console.WriteLine("Filename\\path to write to tape:" + m_Filename);
            if (attr.HasFlag(FileAttributes.Directory))
            {
                // Directory processing

                LogMessage("WriteTape", "Collecting files to write....this may take a while.");



                System.IO.FileStream theFile = System.IO.File.Create("transRead.trn");
                theFile.Close();

                String[] files = new String[0];

                LogMessage("WriteTape", "Reading directory (" + m_Filename + ")");
                try
                {
                    // This may be the slow method of doing this; however, the faster method was throwing errors due to access denied errors.
                    files = getDirs(m_Filename);
                }
                catch (Exception ex)
                {
                    LogError("WriteTape", "Unable to read directory (" + m_Filename + ") error:" + ex.Message);
                }
                using (StreamWriter transWrite = new StreamWriter("transWrite.trn"))
                {
                    foreach (String tmpFile in files)
                    {
                        LogMessage("WriteTape", "Writing file (" + tmpFile + ")");
                        WriteFileHeader(tmpFile);
                        WriteFile(tmpFile);
                        transWrite.WriteLine(tmpFile);
                    }
                }
            }
            else
            {
                WriteFileHeader(FI.FullName);
                WriteFile(FI.FullName);
            }
            if (newTapeScript != string.Empty)
            {
                int retCode = 0;
                //  Will need to return 0 (true) or -1 (false) to continue or not.
                runMTapeScript(EndBackup, newTapeScript, out retCode, "");
            }
        }


        /// <summary>
        /// FUTURE: Command to locate a file in the library, if required prompts for volume to be loaded and then reads the file from the tape.
        /// </summary>
        private static void Locate()
        {
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Files' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    System.IO.FileInfo FI = new FileInfo(m_Filename);
                    cmd = new SQLiteCommand("select Volume, Position from Files where Name = @name and Directory = @directory", conn);
                    cmd.Parameters.AddWithValue("@name", FI.Name);
                    cmd.Parameters.AddWithValue("@directory", FI.DirectoryName);
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    if (rdr2.HasRows)
                    {
                        LogMessage("Locate", "Found file in database");
                        long newLocation = 0;
                        rdr2.Read();
                        int newVolume = rdr2.GetInt32(0);
                        newLocation = rdr2.GetInt64(1);
                        LogMessage("Locate", "Needed volume (" + newVolume + ") Position (" + newLocation + ")");
                        m_VolumeNumber = newVolume;
                        m_NeededPosition = newLocation;
                        NextTape(false);
                        uint retCode = 0;
                        LogMessage("Locate", "Seeking to the requested location on tape...");

                        if (SEEKL(newLocation))
                        {
                            LogMessage("Locate", "Reading...");
                            ReadFile(m_Output);
                        }
                    }
                    else
                    {
                        LogError("Locate", "Specified file is does not exist in the database '" + m_Filename + "'");
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("Locate", "Files table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
        }

        /// <summary>
        /// FUTURE: Command to locate a file in the library and mark for deletion.
        /// </summary>
        private static void Delete()
        {
            Dummy("Delete");
            LogMessage("Delete", "Will eventually mark file for deletion in library");
        }

        /// <summary>
        /// Dummy function to display when an unwritten command is requested
        /// </summary>
        private static void Dummy(String cmd)
        {
            System.Console.WriteLine(cmd + " is not currently implemented.");
        }


        /// <summary>
        /// Media Parameters structure
        /// </summary>
        [StructLayout(LayoutKind.Sequential/*, Pack = 1*/)]
        private struct MediaParameters
        {
            public long Capacity;
            public long Remaining;
            public uint BlockSize;
            public uint PartitionCount;

            public byte IsWriteProtected;
        }

        /// <summary>
        /// Information about the drive and media currently inserted.
        /// </summary>
        private static void Info()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }

            if (API.IsTapeOpen())
            {
                try
                {
                    int blockSize = API.BlockSizeTape;
                    long capacity = API.Capacity;
                    long remaining = API.Remaining;
                    int partcount = API.PartitionCountTape;
                    bool isWriteProtected = API.IsWriteProtected;

                    int blockSizeDrive = API.BlockSizeDrive;
                    int maxBlockSizeDrive = API.MaximumBlockSizeDrive;
                    int minBlockSizeDrive = API.MinimumBlockSizeDrive;
                    int partCountDriveMax = API.PartitionCountDriveMaximum;
                    int featuresLowDrive = API.FeaturesLowDrive;
                    int featuresHighDrive = API.FeaturesHighDrive;
                    int eotWarningZoneSize = API.EOTWarningZoneSize;
                    bool ECC = API.ECC;
                    bool Compression = API.Compression;
                    bool DataPadding = API.DataPadding;
                    bool ReportSetMarks = API.ReportSetmarks;
                    bool IsCompressionCapable = API.IsCompressionCapable;
                    bool IsVariableBlockCapable = API.IsVariableBlockCapable;

                    System.Console.WriteLine("Hardware compression : [" + Compression + "]");
                    System.Console.WriteLine("Default Block Size   : [" + blockSizeDrive + "]");
                    System.Console.WriteLine("Max Block Size       : [" + maxBlockSizeDrive + "]");
                    System.Console.WriteLine("Min Block Size       : [" + minBlockSizeDrive + "]");
                    System.Console.WriteLine("Max Partition Count  : [" + partCountDriveMax + "]");
                    System.Console.WriteLine("Capacity             : [" + capacity + "]");
                    System.Console.WriteLine("Remaining            : [" + remaining + "]");
                    System.Console.WriteLine("Block Size           : [" + blockSize + "]");
                    System.Console.WriteLine("Partition Count      : [" + partcount + "]");
                    System.Console.WriteLine("Write protected      : [" + isWriteProtected + "]");
                }
                catch (Exception ex)
                {
                    int errCode = (int)(uint)Marshal.GetLastWin32Error();
                    if (errCode == TapeWinAPI.ERROR_NO_MEDIA_IN_DRIVE)
                    {
                        LogError("Info", "There is no media in the specified drive.");
                    }
                    else
                    {
                        LogError("Info", "Error " + ex.Message);
                    }
                }
            }
            else
            {
                LogError("FSF", "Unable to open tape drive");
            }

        }

        /// <summary>
        /// FUTURE: Status of the tape library.
        /// </summary>
        private static void Status()
        {
            SQLiteConnection conn = ConnectDatabase();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                SQLiteCommand cmd = new SQLiteCommand("SELECT name FROM sqlite_schema WHERE type='table' and name = 'Files' ORDER BY name", conn);
                SQLiteDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows)
                {
                    cmd = new SQLiteCommand("select * from Files", conn);
                    SQLiteDataReader rdr2 = cmd.ExecuteReader();
                    String fileName = string.Empty;
                    String dirName = string.Empty;
                    String VolumeID = string.Empty;
                    String FileDate = string.Empty;
                    String FileSize = string.Empty;
                    String TapePosition = string.Empty;

                    if (rdr2.HasRows)
                    {
                        LogMessage("Status", "Tape Library Contents");
                        LogMessage("Status", "");
                        LogMessage("Status", padRight("Directory", 50) + padRight("File Name", 20) + padRight("Volume ID", 15) + padRight("File Date", 40) + padRight("File Size", 15) + padRight("Tape Position", 15));
                        int recCount = 0;
                        int delCount = 0;

                        List<String> volumes = new List<string>();

                        while (rdr2.Read())
                        {
                            bool isDeleted = false;
                            recCount++;
                            String tmpVolId = string.Empty;
                            for (int i = 0; i < rdr2.FieldCount; i++)
                            {
                                switch (rdr2.GetName(i))
                                {
                                    case "name":
                                        fileName = rdr2.GetString(i);
                                        break;
                                    case "directory":
                                        dirName = rdr2.GetString(i);
                                        break;
                                    case "Volume":
                                        VolumeID = "VOL_" + zeroFill(rdr2.GetInt32(i), 4);
                                        break;
                                    case "FileDate":
                                        FileDate = rdr2.GetString(i);
                                        break;
                                    case "Size":
                                        FileSize = rdr2.GetInt32(i).ToString();
                                        break;
                                    case "Position":
                                        TapePosition = rdr2.GetInt64(i).ToString();
                                        break;
                                    case "Deleted":
                                        bool tmpDel = rdr2.GetBoolean(i);
                                        if (tmpDel)
                                        {
                                            delCount++;
                                            isDeleted = true;
                                        }
                                        else
                                        {

                                            if (!volumes.Contains(VolumeID))
                                            {
                                                volumes.Add(VolumeID);
                                            }

                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                            if (!isDeleted)
                            {
                                if (dirName.Length > 50)
                                {
                                    dirName = dirName.Substring(0, 50);
                                }
                                LogMessage("Status", padRight(dirName, 50) + padRight(fileName, 20) + padRight(VolumeID, 15) + padRight(FileDate, 40) + padRight(FileSize, 15) + padRight(TapePosition, 15));
                            }
                        }
                        LogMessage("Status", "");
                        LogMessage("Status", "Total Records (" + recCount + ")");
                        LogMessage("Status", "Total Deleted Records (" + delCount + ")");
                        LogMessage("Status", "Total Volumes Used (" + volumes.Count + ")");
                    }
                    else
                    {
                        LogError("Status", "There are no records in the tape library at this time.");
                    }
                    rdr2.Close();
                }
                else
                {
                    LogError("Status", "Files table does not exist in the database");
                }
                rdr.Close();
            }
            conn.Close();
        }

        /// <summary>
        /// Lock the drive
        /// </summary>
        private static void Lock()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retCode = 0;
                if (!API.Lock(ref retCode))
                {
                    LogError("Lock", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
                else
                {
                    LogMessage("Lock", "Tape drive door locked.");
                }
            }
            else
            {
                LogError("Lock", "Unable to open tape drive");
            }

        }

        /// <summary>
        /// Unlock the drive
        /// </summary>
        private static void Unlock()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retCode = 0;
                if (!API.Unlock(ref retCode))
                {
                    LogError("Unlock", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
                else
                {
                    LogMessage("Unlock", "Tape drive door is unlockec.");
                }
            }
            else
            {
                LogError("Unlock", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// Forward Space File mark(s)
        /// </summary>
        private static void FSF()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retcode = 0;
                if (API.SpaceFileMarks(m_Count, ref retcode))
                {
                    LogMessage("FSF", "Ok");
                }
                else
                {
                    if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                    {
                        LogError("FSF", "No data detected on drive.");
                    }
                    else
                    {
                        LogError("FSF", "Undefined error code " + retcode);
                    }
                }
            }
            else
            {
                LogError("FSF", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// Backward space file mark(s)
        /// </summary>
        private static void BSF()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retcode = 0;
                if (API.SpaceFileMarks(m_Count * -1, ref retcode))
                {
                    LogMessage("BSF", "Ok");
                }
                else
                {
                    if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                    {
                        LogError("BSF", "No data detected on drive.");
                    }
                    else
                    {
                        LogError("BSF", "Undefined error code " + retcode);
                    }
                }
            }
            else
            {
                LogError("FSF", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// Forward space set mark(s)
        /// </summary>
        private static void FSS()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retcode = 0;
                if (API.SpaceBlocks(m_Count, ref retcode))
                {
                    LogMessage("FSS", "Ok");
                }
                else
                {
                    if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                    {
                        LogError("FSS", "No data detected on drive.");
                    }
                    else
                    {
                        LogError("FSS", "Undefined error code " + retcode);
                    }
                }
            }
            else
            {
                LogError("FSS", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// Backward space set mark(s)
        /// </summary>
        private static void BSS()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retcode = 0;
                if (API.SpaceFileMarks(m_Count * -1, ref retcode))
                {
                    LogMessage("BSS", "Ok");
                }
                else
                {
                    if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                    {
                        LogError("BSS", "No data detected on drive.");
                    }
                    else
                    {
                        LogError("BSS", "Undefined error code " + retcode);
                    }
                }
            }
            else
            {
                LogError("BSS", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// Forward space n file marks and then reverse one record.
        /// </summary>
        private static void FSFM()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retcode = 0;
                if (API.SpaceFileMarks(m_Count, ref retcode))
                {
                    if (API.SpaceBlocks(-1, ref retcode))
                    {
                        LogMessage("FSFM", "Ok");
                    }
                    else
                    {
                        if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                        {
                            LogError("FSFM", "No data detected on drive.");
                        }
                        else
                        {
                            LogError("FSFM", "Undefined error code " + retcode);
                        }
                    }
                    LogMessage("FSFM", "Ok");
                }
                else
                {
                    if (retcode == TapeWinAPI.ERROR_NO_DATA_DETECTED)
                    {
                        LogError("FSFM", "No data detected on drive.");
                    }
                    else
                    {
                        LogError("FSFM", "Undefined error code " + retcode);
                    }
                }
            }
            else
            {
                LogError("FSFM", "Unable to open tape drive");
            }
        }

        /// <summary>
        /// FUTURE: Backward space file mark(s) and forward 1 record.
        /// </summary>
        private static void BSFM()
        {
            Dummy("BSFM");
        }


        /// <summary>
        /// FUTURE: Rewind and then forward space file mark(s)
        /// </summary>
        private static void ASF()
        {
            Dummy("ASF");
        }

        /// <summary>
        /// FUTURE: Forward space count records 
        /// </summary>
        private static void FSR()
        {
            Dummy("FSR");
        }

        /// <summary>
        /// Rewind the tape (i.e. go to position 0)
        /// </summary>
        private static void Rewind()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                uint retCode = 0;
                if (!API.Rewind(ref retCode))
                {
                    LogError("Rewind", "Unable to rewind the tape " + new Win32Exception((int)retCode).Message);
                }
                else
                {
                    LogMessage("Rewind", "Tape rewound");
                }
            }
            else
            {
                LogError("Rewind", "Unable to open specified tape device");
            }
        }


        /// <summary>
        /// FUTURE: Backward space count records
        /// </summary>
        private static void BSR()
        {
            Dummy("BSR");
        }


        /// <summary>
        /// Got to the end of data
        /// </summary>
        private static void EOD()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.SeekToEOD(ref retCode))
                {
                    LogError("EOD", "Error seeking to EOD " + retCode);
                }
                else
                {
                    LogMessage("EOD", "Ok");
                }
            }
            else
            {
                LogError("EOD", "Unable to connect to tape drive");
            }
        }

        /// <summary>
        /// Eject the volume from the tape drive
        /// </summary>
        private static void EJECT()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.Rewind(ref retCode))
                {
                    LogError("Eject", "Error rewinding tape " + retCode);
                }
                else
                {
                    if (!API.Unload(ref retCode))
                    {
                        LogError("Eject", "Error unloading tape " + retCode);
                    }
                    else
                    {
                        LogMessage("Eject", "Tape ejected.");
                    }
                }
            }
            else
            {
                LogError("Eject", "Unable to connect to tape drive");
            }
        }

        /// <summary>
        /// Retention the tape in the drive....long process
        /// </summary>
        private static void RETENTION()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.AdjustTension(ref retCode))
                {
                    LogError("Retention", "Error retentioning tape " + retCode);
                }
                else
                {
                    LogMessage("Retention", "Tape retentioned.");
                }
            }
            else
            {
                LogError("Retention", "Unable to connect to tape drive");
            }
        }


        /// <summary>
        /// Write count end of file markers
        /// </summary>
        private static void EOF()
        {
            uint retCode = 0;
            int tmpOut = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                int newCount = m_Count;
                newCount = newCount == 0 ? 1 : newCount;
                Boolean EOFWritten = false;
                do
                {
                    if (API.WriteTapemark((uint)newCount, TapeWinAPI.TAPE_FILEMARKS, ref retCode))
                    {
                        EOFWritten = true;
                    }
                    else
                    {
                        switch (retCode)
                        {
                            case 0234:  //more data is available
                                        //update error log and status
                                LogError("EOF", "aborted due to SCSI Controller problem");
                                runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted due to SCSI Controller problem.");
                                break;
                            case 1106:  //incorrect block size
                                        //update error log and status
                                LogError("EOF", "aborted due to invalid block size");
                                runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted due to invalid block size.");
                                break;
                            case 1129:  // Physical EOT
                            case 1100:  // EOT marker reached
                                LogError("EOT", "EOT detected.");
                                if (NextTape())
                                {
                                    m_VolumeNumber++;
                                    addVolumeToDatabase();
                                    continue;
                                }
                                break;
                            case 1110:  // Media Changed (normal if new tape was inserted.
                                continue;
                            case 1156:  // Device Requires Cleaning
                                LogMessage("EOF", "Tape Drive Requires Cleaning.");
                                runMTapeScript(Clean, newTapeScript, out tmpOut, "Tape drive requires cleaning.");
                                break;
                            case 1112:
                                LogError("EOF", "There is no tape in the drive.");
                                Environment.Exit(1);
                                break;
                            case 0:
                                break;
                            default:    //any other errors
                                LogError("EOF", "aborted with error code: " + retCode.ToString());
                                runMTapeScript(Alert, newTapeScript, out tmpOut, "Aborted Error:" + retCode.ToString());
                                break;
                        }
                    }
                }
                while (!EOFWritten);
            }
            else
            {
                LogError("EOF", "Unable to connect to tape drive");
            }
        }

        /// <summary>
        /// Write count set marks
        /// </summary>
        private static void WSET()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.WriteTapemark((uint)m_Count, TapeWinAPI.TAPE_SETMARKS, ref retCode))
                {
                    LogError("WSET", "Error writing set mark(s) " + retCode + " " + new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
                else
                {
                    LogMessage("WSET", "Wrote set mark(s).");
                }
            }
            else
            {
                LogError("WSET", "Unable to connect to tape drive");
            }
        }


        /// <summary>
        /// Erase the tape...will take a long time.
        /// </summary>
        private static void ERASE()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.Erase(TapeWinAPI.TAPE_DRIVE_ERASE_IMMEDIATE, ref retCode))
                {
                    LogError("ERASE", "Error erasing tape " + retCode);
                }
                else
                {
                    LogMessage("ERASE", "Tape erased");
                }
            }
            else
            {
                LogError("ERASE", "Unable to connect to tape drive");
            }

        }


        /// <summary>
        /// Seek to either a logical or absolute position on the tape
        /// </summary>
        /// <param name="bool Absolute">If true, use absolute positioning</param>
        private static bool SEEK(bool Absolute, long Location = -1)
        {
            bool retValue = false;
            bool tapeOpen = false;
            long whereTo = m_Count;

            if (Location != -1)
            {
                whereTo = Location;
            }
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            String label = "Logical";
            String func = "SEEKL";
            if (Absolute)
            {
                label = "Absolute";
                func = "SEEKA";
            }
            if (API.IsTapeOpen())
            {
                uint retCode = 0;
                if (Absolute)
                {
                    if (!API.SeekToAbsoluteBlock(whereTo, ref retCode))
                    {
                        retValue = false;
                        LogError(func, "Unable to seek the specified " + label + " location. " + new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message + " Absolute positioning may not be supported.");
                    }
                    else
                    {
                        retValue = true;
                        LogMessage(func, "Reached specified " + label + " location.");
                    }
                }
                else
                {
                    if (!API.SeekToLogicalBlock(whereTo, ref retCode))
                    {
                        retValue = false;
                        LogError(func, "Unable to seek the specified " + label + " location. " + new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                    }
                    else
                    {
                        retValue = true;
                        LogMessage(func, "Reached specified " + label + " location.");
                    }
                }
            }
            else
            {
                LogError("Open", "Unable to connect to specified tape drive.");
            }
            return retValue;
        }


        /// <summary>
        /// Command to seek the absolute position on the tape
        /// </summary>
        private static bool SEEKA(long Location = -1)
        {
            return SEEK(true, Location);
        }


        /// <summary>
        /// Command to seek the logical position on the tape
        /// </summary>
        private static bool SEEKL(long Location = -1)
        {
            return SEEK(false, Location);
        }

        private static long getCurrentPosition()
        {
            long retValue = -1;
            //            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }

            if (API.IsTapeOpen())
            {
                long position = 0;
                uint retCode = 0;
                if (API.GetTapePosition(ref position, ref retCode))
                {
                    m_CurrentPosition = position;
                    retValue = m_CurrentPosition;
                }
            }
            return retValue;
        }



        private static int scriptProcessor(String funcName, List<String> scriptLines, List<string> Parameters = null)
        {
            int retValue = -1;
            int retCode = 0;
            String regExPattern = "[A-Z0-9$\\(\\)\\\"\\,]+| (\\\"([^\\\"])*\\\")|(^[A-Z0-9\\s\\=\\\"]*)";
            Regex rg = new Regex(regExPattern, RegexOptions.IgnoreCase);
            Dictionary<String, String> mVariables = new Dictionary<string, string>();

            try
            {
                if (mFunctParams.ContainsKey(funcName))
                {
                    int parmNum = 0;
                    try
                    {
                        foreach (String parameterName in mFunctParams[funcName])
                        {
                            if (Parameters != null)
                            {
                                if (!mVariables.ContainsKey("$" + parameterName.Replace("\"", "").ToUpper()))
                                {
                                    mVariables.Add("$" + parameterName.Replace("\"", "").ToUpper(), Parameters[parmNum]);
                                }
                                else
                                {
                                    mVariables["$" + parameterName.Replace("\"", "").ToUpper()] = Parameters[parmNum];
                                }
                            }
                            parmNum++;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError("scriptProcessor:Parameters:", ex.Message);
                    }

                }
                int i = 1;

                String switchText = String.Empty;

                bool inSwitch = false;
                bool caseTrue = false;

                foreach (String tmpText in scriptLines)
                {
                    String newText = tmpText;
                    if (!tmpText.Trim().StartsWith("#"))
                    {
                        // Replace the global variables
                        try
                        {
                            foreach (String varName in m_GlobalVariables.Keys)
                            {
                                newText = newText.Replace(varName.ToUpper(), m_GlobalVariables[varName]);
                            }
                        }
                        catch(Exception ex)
                        {
                            LogError("ScriptProcessor:Global Variables:",ex.Message);
                        }

                        // Replace the local variables
                        try
                        {
                            foreach (String varName in mVariables.Keys)
                            {
                                newText = newText.Replace(varName.ToUpper(), mVariables[varName]);
                            }
                        }
                        catch(Exception ex)
                        {
                            LogError("ScriptProcessor:Local Variables:",ex.Message);
                        }
                        if (inSwitch)
                        {
                            if (newText.ToLower() == "end select")
                            {
                                inSwitch = false;
                            }
                            else
                            {
                                if (newText.ToLower().StartsWith("case"))
                                {
                                    string condition = newText.Substring(newText.IndexOf(" ") + 1).Replace(":", "").Trim();
                                    if (condition.ToLower().Trim() == switchText.ToLower().Trim())
                                    {
                                        caseTrue = true;
                                    }
                                }
                                else
                                {
                                    if (caseTrue)
                                    {
                                        caseTrue = false;
                                        string tmpFunc = newText.Substring(0, newText.IndexOf("("));
                                        if (mFunctions.ContainsKey(tmpFunc))
                                        {
                                            retCode = retCode + scriptProcessor(tmpFunc, mFunctions[tmpFunc]);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (newText.ToLower().StartsWith("select("))
                            {
                                switchText = newText.Substring(newText.IndexOf("(") + 1).Trim().Replace(")", "").Replace("(", "");
                                inSwitch = true;
                            }
                            else
                            {
                                if (!newText.Trim().StartsWith("#") && newText.Trim() != string.Empty)
                                {
                                    if (!newText.Contains("="))
                                    {
                                        MatchCollection matches = rg.Matches(newText);
                                        if (matches.Count > 0)
                                        {
                                            String workString = matches[0].Value;

                                            if (!workString.Contains("="))
                                            {

                                                if (workString.Contains("(") && mFunctions.ContainsKey(workString.Substring(0, workString.IndexOf("("))))
                                                {
                                                    List<String> mParameters = new List<String>();
                                                    if (!newText.Contains("()"))
                                                    {
                                                        int pStart = newText.IndexOf("(") + 1;
                                                        int pEnd = newText.IndexOf(")");
                                                        String parmString = newText.Substring(pStart, pEnd - pStart);
                                                        parmString = parmString.Replace("\"", "");
                                                        string[] tParmString = parmString.Split(',');
                                                        mParameters = tParmString.ToList<String>();
                                                    }
                                                    retCode = retCode + scriptProcessor(workString.Substring(0, workString.IndexOf("(")), mFunctions[workString.Substring(0, workString.IndexOf("("))], mParameters);
                                                }
                                                else
                                                {
                                                    switch (matches[0].Value.ToLower())
                                                    {
                                                        case "eject()":
                                                            try
                                                            {
                                                                // Commented out for testing purposes
                                                                EJECT();
                                                            }
                                                            catch(Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:Eject:",ex.Message);
                                                            }
                                                            break;
                                                        case "keypress()":
                                                            System.Console.ReadKey();
                                                            break;
                                                        case "detectnewtape()":
                                                            try
                                                            {
                                                                System.Console.WriteLine("[" + funcName + "] Detecting new tape in drive.");
                                                                if (!DetectNewTape())
                                                                {
                                                                    return 1;
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:Eject:", ex.Message);
                                                            }

                                                            // tape detection 
                                                            break;
                                                        case "sendmail()":
                                                            try
                                                            {
                                                                System.Console.WriteLine("[" + funcName + "] Sending email");
                                                                if (!mVariables.ContainsKey("Reason"))
                                                                {
                                                                    mVariables.Add("Reason", m_GlobalVariables["$REASON"]);
                                                                }
                                                                else
                                                                {
                                                                    mVariables["Reason"] = m_GlobalVariables["$REASON"];
                                                                }
                                                                if (!SendMail(mVariables))
                                                                {
                                                                    return 1;
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:SendMail:", ex.Message);
                                                            }

                                                            // send mail
                                                            break;
                                                        case "print":
                                                            try
                                                            {
                                                                if (m_GlobalVariables["$REASON"] == "New Tape")
                                                                {
                                                                    System.Console.WriteLine("[" + funcName + "] " + matches[1].Value);
                                                                }
                                                                else
                                                                {
                                                                    System.Console.WriteLine("[" + funcName + "] " + matches[1].Value);
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:Print:", ex.Message);
                                                            }

                                                            break;
                                                        case "call":
                                                            // Calls external application or script
                                                            try
                                                            {
                                                                if (!RunExternal(out retCode, " + funcName + "))
                                                                {
                                                                    System.Console.WriteLine("[" + funcName + "] External command reported an error :" + retCode);
                                                                    return 1;
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:Call:", ex.Message);
                                                            }

                                                            break;
                                                        default:
                                                            try
                                                            {
                                                                if (tmpText.Contains("="))
                                                                {
                                                                    // Variables defined in the main (i.e. main script) are globally available.
                                                                    if (funcName == "main")
                                                                    {
                                                                        if (!m_GlobalVariables.ContainsKey(matches[0].Value))
                                                                        {
                                                                            m_GlobalVariables.Add(matches[0].Value, matches[1].Value);
                                                                        }
                                                                        else
                                                                        {
                                                                            m_GlobalVariables[matches[0].Value] = matches[1].Value;
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        if (!mVariables.ContainsKey(matches[0].Value))
                                                                        {
                                                                            mVariables.Add(matches[0].Value, matches[1].Value);
                                                                        }
                                                                        else
                                                                        {
                                                                            mVariables[matches[0].Value] = matches[1].Value;
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    System.Console.WriteLine("[" + funcName + "] Unrecognized command line " + tmpText);
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                LogError("ScriptProcessor:Default:", ex.Message);
                                                            }

                                                            break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (funcName == "main")
                                                {
                                                    try
                                                    {
                                                        if (!m_GlobalVariables.ContainsKey(matches[0].Value))
                                                        {
                                                            m_GlobalVariables.Add(matches[0].Value, matches[1].Value);
                                                        }
                                                        else
                                                        {
                                                            m_GlobalVariables[matches[0].Value] = matches[1].Value;
                                                        }
                                                    }
                                                    catch(Exception ex)
                                                    {
                                                        LogError("ScriptProcessor:Set Global:", ex.Message);
                                                    }

                                                }
                                                else
                                                {
                                                    try
                                                    {
                                                        if (!mVariables.ContainsKey(matches[0].Value))
                                                        {
                                                            mVariables.Add(matches[0].Value, matches[1].Value);
                                                        }
                                                        else
                                                        {
                                                            mVariables[matches[0].Value] = matches[1].Value;
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        LogError("ScriptProcessor:Set Local:", ex.Message);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            System.Console.WriteLine("[" + funcName + "] Unmatched command line '" + newText + "'");
                                        }
                                    }
                                    else
                                    {
                                        string[] splitted = newText.Split('=');
                                        if (funcName == "main")
                                        {
                                            try
                                            {
                                                if (!m_GlobalVariables.ContainsKey(splitted[0]))
                                                {
                                                    m_GlobalVariables.Add(splitted[0], splitted[1].Replace("\"", ""));
                                                }
                                                else
                                                {
                                                    m_GlobalVariables[splitted[0]] = splitted[1].Replace("\"", "");
                                                }
                                            }
                                            catch(Exception ex)
                                            {
                                                LogError("scriptProcessor", "Error setting global variable :" + splitted[0] + " Error:" + ex.Message);
                                            }
                                        }
                                        else
                                        {
                                            try
                                            {
                                                if (!mVariables.ContainsKey(splitted[0]))
                                                {
                                                    mVariables.Add(splitted[0], splitted[1].Replace("\"", ""));
                                                }
                                                else
                                                {
                                                    mVariables[splitted[0]] = splitted[1].Replace("\"", "");
                                                }
                                            }
                                            catch(Exception ex)
                                            {
                                                LogError("scriptProcessor", "Error setting local variable :" + splitted[0] + " Error:" + ex.Message);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogError("ScriptProcessor", "An unexpected error occurred during the processing of the mTape2 script.  Error:" + ex.Message);
                retValue = -1;
            }

            return retValue;
        }



        /// <summary>
        /// Interpret an mTape script
        /// </summary>
        /// <returns>bool Success</returns>
        public static bool runMTapeScript(int ReasonCode, String Script, out int returnCode, String Message = "")
        {
            bool retValue = false;
            int retCode = 0;
            string[] text = System.IO.File.ReadAllLines(Script);
            try
            {
                m_Message = Message;
                String funPattern = "[A-Z0-9$\\(\\)\\\"\\,]+| (\\\"([^\\\"])*\\\")";

                List<String> functionLines = new List<string>();
                bool ReadFunction = false;
                String functionName = string.Empty;

                String Reason = string.Empty;
                switch (ReasonCode)
                {
                    case NewTape:
                        Reason = "New Tape";
                        break;
                    case StartBackup:
                        Reason = "Archive Started";
                        break;
                    case EndBackup:
                        Reason = "Archive Ended";
                        break;
                    case Alert:
                        Reason = "Alert";
                        break;
                    case TapeNeeded:
                        Reason = "Tape Needed";
                        break;
                    case Clean:
                        Reason = "Clean";
                        break;
                }

                if (!m_GlobalVariables.ContainsKey("$VOLUME"))
                {
                    m_GlobalVariables.Add("$VOLUME", "VOL_" + zeroFill(m_VolumeNumber, 4));
                }
                else
                {
                    m_GlobalVariables["$VOLUME"] = "VOL_" + zeroFill(m_VolumeNumber, 4);
                }
                if (!m_GlobalVariables.ContainsKey("$REASON"))
                {
                    m_GlobalVariables.Add("$REASON", Reason);
                }
                else
                {
                    m_GlobalVariables["$REASON"] = Reason;
                }
                if (!m_GlobalVariables.ContainsKey("$DATETIME"))
                {
                    m_GlobalVariables.Add("$DATETIME", DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
                }
                else
                {
                    m_GlobalVariables["$DATETIME"] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
                }
                if (!m_GlobalVariables.ContainsKey("$1"))
                {
                    m_GlobalVariables.Add("$1", Reason);
                }
                else
                {
                    m_GlobalVariables["$1"] = Reason;
                }
                if (!m_GlobalVariables.ContainsKey("$2"))
                {
                    m_GlobalVariables.Add("$2", "VOL_" + zeroFill((m_VolumeNumber + 1), 4));
                }
                else
                {
                    m_GlobalVariables["$2"] = "VOL_" + zeroFill(m_VolumeNumber + 1, 4);
                }
                if (!m_GlobalVariables.ContainsKey("$0"))
                {
                    m_GlobalVariables.Add("$0", "VOL_" + zeroFill((m_VolumeNumber + 1), 4));
                }
                else
                {
                    m_GlobalVariables["$0"] = "VOL_" + zeroFill(m_VolumeNumber + 1, 4);
                }
                if (!m_GlobalVariables.ContainsKey("$3"))
                {
                    m_GlobalVariables.Add("$3", m_Message);
                }
                else
                {
                    m_GlobalVariables["$3"] = m_Message;
                }


                foreach (String tmpText in text)
                {
                    String workText = tmpText.Trim();
                    if (!ReadFunction)
                    {

                        if (!workText.StartsWith("#") && workText != string.Empty)
                        {
                            Regex rgf = new Regex(funPattern, RegexOptions.IgnoreCase);
                            MatchCollection matches = rgf.Matches(workText);
                            if (matches.Count > 0)
                            {
                                if (workText.EndsWith("{"))
                                {
                                    functionName = matches[0].Value;
                                    functionName = functionName.Substring(0, functionName.IndexOf("("));
                                    ReadFunction = true;
                                }
                                else
                                {
                                    functionName = "main";
                                    ReadFunction = true;
                                    functionLines.Add(workText);
                                }

                                if (tmpText.Contains("(") && !tmpText.Contains("()"))
                                {
                                    int pStart = tmpText.IndexOf("(") + 1;
                                    int pEnd = tmpText.IndexOf(")");
                                    String tParms = tmpText.Substring(pStart, pEnd - pStart);
                                    string[] tParmItems = tParms.Split(',');
                                    List<String> tParmNames = new List<string>();
                                    foreach (String tmp in tParmItems)
                                    {
                                        tParmNames.Add(tmp);
                                    }
                                    if (!mFunctParams.ContainsKey(functionName))
                                    {
                                        mFunctParams.Add(functionName, tParmNames);
                                    }
                                    else
                                    {
                                        mFunctParams[functionName] = tParmNames;
                                    }
                                }

                            }
                            else
                            {
                                if (!mFunctions.ContainsKey("main"))
                                {
                                    functionName = "main";
                                    ReadFunction = true;
                                    functionLines.Add(workText);

                                }
                                else
                                {
                                    mFunctions["main"].Add(workText);

                                }
                            }
                        }
                    }
                    else
                    {
                        if (workText.StartsWith("}"))
                        {
                            ReadFunction = false;
                            if (!mFunctions.ContainsKey(functionName))
                            {
                                mFunctions.Add(functionName, functionLines);
                            }
                            else
                            {
                                mFunctions[functionName] = functionLines;
                            }
                            functionLines = new List<string>();
                        }
                        else
                        {
                            functionLines.Add(workText);
                            retCode = 0;
                        }

                    }
                }
                if (ReadFunction)
                {
                    if (!mFunctions.ContainsKey("main"))
                    {
                        mFunctions.Add("main", functionLines);
                    }
                    else
                    {
                        mFunctions["main"] = functionLines;
                    }
                    if (mFunctions.ContainsKey("main"))
                    {
                        retCode = scriptProcessor("main", mFunctions["main"]);
                    }
                    else
                    {
                        retCode = 0;
                    }
                }
            }
            catch(Exception ex)
            {
                LogError("runMTapeScript", "An unexpected error occurred during script execution. Error:" + ex.Message);
                var w32ex = ex as Win32Exception;
                if (w32ex == null)
                {
                    w32ex = ex.InnerException as Win32Exception;
                }
                if (w32ex != null)
                {
                    returnCode = w32ex.ErrorCode;
                    // do stuff
                }
                retValue = false;
            }
            returnCode = retCode;
            return retValue;
        }


        /// <summary>
        /// Prints the current position on the tape
        /// </summary>
        private static void TELL()
        {
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }

            if (API.IsTapeOpen())
            {
                long position = 0;
                uint retCode = 0;
                if (API.GetTapePosition(ref position, ref retCode))
                {
                    LogMessage("TELL", "Position: " + position);
                }
                else
                {
                    if(retCode == 1112)
                    {
                        LogMessage("TELL", "There isn't a tape in the drive.");
                    }
                }
            }
            else
            {
                System.Console.WriteLine("TELL Error: Tape is not open");
            }
        }

        /// <summary>
        /// FUTURE: Set the current partition to use
        /// </summary>
        private static void SETPARTITION()
        {
            Dummy("Setpartition");
        }

        /// <summary>
        /// FUTURE: Seek in this partition
        /// </summary>
        private static void PARTSEEK()
        {
            Dummy("Partseek");
        }

        /// <summary>
        /// FUTURE: Make a partition
        /// </summary>
        private static void MKPARTITION()
        {
            Dummy("Makepartition");
        }

        /// <summary>
        /// Load a tape volume into drive.
        /// </summary>
        private static void Load()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                if (!API.Load(ref retCode))
                {
                    LogError("Load", "Error loading tape " + retCode);
                }
                else
                {
                    LogMessage("Load", "Tape loaded.");
                }
            }
            else
            {
                LogError("Load", "Unable to connect to tape drive");
            }
        }


        /// <summary>
        /// Set the block size
        /// </summary>
        private static void SETBLK()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                TapeWinAPI.TapeSetMediaParameters TParms = new TapeWinAPI.TapeSetMediaParameters();

                TParms.BlockSize = (uint)m_Count;
                if (!API.SetMediaParameters(TParms, ref retCode))
                {
                    LogError("SETBLK", "Error setting the block size " + retCode);
                }
                else
                {
                    LogMessage("SETBLK", "Block size set.");
                }
            }
            else
            {
                LogError("SETBLK", "Unable to connect to tape drive");
            }

        }

        /// <summary>
        /// FUTURE: Print the densities supported
        /// </summary>
        private static void DENSITIES()
        {
            Dummy("Densities");
        }

        /// <summary>
        /// FUTURE: Set the tape drive buffer code to number.
        /// </summary>
        private static void DRVBUFFER()
        {
            Dummy("Drvbuffer");
        }

        /// <summary>
        /// Enable/disable the compression on the drive.
        /// </summary>
        private static void COMPRESSION()
        {
            uint retCode = 0;
            bool tapeOpen = false;
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
            }
            if (API.IsTapeOpen())
            {
                TapeWinAPI.TapeSetDriveParameters TParms = new TapeWinAPI.TapeSetDriveParameters();


                TParms.driveCompression = (byte)m_Count;

                if (!API.SetDriveParameters(TParms, ref retCode))
                {
                    LogError("COMPRESSION", "Error setting the compression " + retCode);
                }
                else
                {
                    LogMessage("COMPRESSION", "Compression set.");
                }
            }
            else
            {
                LogError("COMPRESSION", "Unable to connect to tape drive");
            }

        }

        /// <summary>
        /// FUTURE: Sets the driver options (likely unsupported in Windows)
        /// </summary>
        private static void STOPTIONS()
        {
            Dummy("Stoptions");
        }

        /// <summary>
        /// FUTURE: Sets the select drive options (likely unsupported in Windows)
        /// </summary>
        private static void STSETOPTIONS()
        {
            Dummy("Stsetoptions");
        }

        /// <summary>
        /// FUTURE: Clears the selected driver options (likely unsupported in Windows)
        /// </summary>
        private static void STCLEAROPTIONS()
        {
            Dummy("Stclearoptions");
        }

        /// <summary>
        /// FUTURE: Prints the currently enabled options for the device. (Likely unsupported in Windows); 
        /// </summary>
        private static void STSHOWOPTIONS()
        {
            Dummy("Stshowoptions");
        }

        /// <summary>
        /// FUTURE: The write threshold for the tape device in count kilobytes.  (Likely unsupported in Windows);
        /// </summary>
        private static void STWRTHRESHOLD()
        {
            Dummy("Stwrthreshold");
        }

        /// <summary>
        /// FUTURE: Sets the default block size on the device to count bytes
        /// </summary>
        private static void DEFBLKSIZE()
        {
            Dummy("Defblksize");
        }

        /// <summary>
        /// FUTURE: Sets the default density code. -1 to reset to default
        /// </summary>
        private static void DEFDENSITY()
        {
            Dummy("Defdensity");
        }

        /// <summary>
        /// FUTURE: Sets the default drive buffer code.  -1 to reset to default.
        /// </summary>
        private static void DEFDRVBUFFER()
        {
            Dummy("Defdrvbuffer");
        }

        /// <summary>
        /// FUTURE: Sets the normal timeout for the device.  (Likely unsupported by Windows)
        /// </summary>
        private static void STTIMEOUT()
        {
            Dummy("Sttimeout");
        }

        /// <summary>
        /// FUTURE: Sets the default compression state.  -1 to reset to default.
        /// </summary>
        private static void DEFCOMPRESSION()
        {
            Dummy("Defcompression");
        }

        /// <summary>
        /// FUTURE: Sets the long timeout for the device.  (Likely unsupported in Windows)
        /// </summary>
        private static void STLONGTIMEOUT()
        {
            Dummy("Stlongtimeout");
        }

        /// <summary>
        /// FUTURE: Sets the cleaning request interpretation parameters (Likely unsupported in Windows)
        /// </summary>
        private static void STSETCLN()
        {
            Dummy("Stsetcln");
        }

        /// <summary>
        /// Initiate tape connection and set tape drive to the appropriate name.
        /// </summary>
        /// <param name="string tapeDrive"></param>
        private static void SetTape(String tapeDrive)
        {
            API.TapeDrive = tapeDrive;
        }

        /// <summary>
        /// Pads a string with the width number of spaces at the end.
        /// </summary>
        /// <param name="string text"></param>
        /// <param name="int width"></param>
        /// <returns>The padded string</returns>
        private static String padRight(String text, int width)
        {
            String RetValue = text;
            while (RetValue.Length < width)
            {
                RetValue = RetValue + " ";
            }
            return RetValue;
        }

        /// <summary>
        /// Pads a string with the width number of spaces at the beginning of the string
        /// </summary>
        /// <param name="string text"></param>
        /// <param name="int width"></param>
        /// <returns>The padded string</returns>
        private static String padLeft(String text, int width)
        {
            String RetValue = text;
            while (RetValue.Length < width)
            {
                RetValue = " " + RetValue;
            }
            return RetValue;
        }

        /// <summary>
        /// Determines the which number is the smaller of the two
        /// </summary>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <returns>the value of the smaller</returns>
        private static int min(int i, int j)
        {
            if (i < j)
            {
                return i;
            }
            else
            {
                return j;
            }
        }

        /// <summary>
        /// Algorithm to wrap text displayed on the screen to fit the screen.
        /// 
        /// Will optionally ignore the tab position on the first line of wrapped text...this is useful when the cursor is already tabbed.
        /// Will optionally ignore the tab on the first line completely.
        /// </summary>
        /// <param name="string text"></param>
        /// <param name="int tabPosition"></param>
        /// <param name="bool ignoreFirst"></param>
        /// <param name="bool ignoreTabOnFirst"></param>
        private static void wrapText(String text, int tabPosition = 10, bool ignoreFirst = true, bool ignoreTabOnFirst = false)
        {
            int screenWidth = System.Console.WindowWidth - tabPosition - 5;
            String tmpString = text;
            String wrkString = string.Empty;
            String[] wrdSplit = text.Split(' ');
            string outString = string.Empty;
            bool newLine = true;
            bool firstRun = true;

            foreach (String tmpWrd in wrdSplit)
            {
                if (!ignoreFirst)
                {
                    if (newLine)
                    {
                        if (!ignoreTabOnFirst)
                        {
                            outString = padLeft(" ", tabPosition - 1) + outString;
                            ignoreTabOnFirst = false;
                        }
                        newLine = false;
                    }
                    ignoreFirst = false;
                }
                if (outString.Length + 1 + tmpWrd.Length <= screenWidth)
                {
                    if (!firstRun)
                    {
                        if (outString != string.Empty)
                        {
                            outString = outString + " " + tmpWrd;
                        }
                        else
                        {
                            outString = tmpWrd;
                        }
                    }
                    else
                    {
                        outString = outString + tmpWrd;
                        firstRun = false;
                    }
                    newLine = false;
                }
                else
                {
                    System.Console.WriteLine(outString);
                    outString = string.Empty;
                    newLine = true;
                    outString = padLeft(" ", tabPosition - 1) + tmpWrd;
                    ignoreFirst = true;
                }
            }
            if (outString != string.Empty)
            {
                System.Console.WriteLine(outString);
            }
        }

        /// <summary>
        /// Pseudo replacement for System.Console.Write/WriteLine
        /// This method is useful for simulating the manpage type display of help to the screen.
        /// </summary>
        /// <param name="string titleText (i.e. first column)"></param>
        /// <param name="string cmdText (i.e. second column)"></param>
        /// <param name="string detailText (i.e. third column)"></param>
        private static void ScreenWriter(String titleText, String cmdText = "", String detailText = "")
        {
            bool ignoreTab = false;
            bool newLine = false;
            if (titleText.Length <= 10)
            {
                if (!ignoreTab)
                {
                    System.Console.Write(padRight(titleText, 9));
                    ignoreTab = true;
                }
                else
                {
                    System.Console.Write(titleText);
                }
            }
            else
            {
                System.Console.WriteLine(titleText);
                ignoreTab = true;
            }
            if (cmdText.Length <= 15)
            {
                if (cmdText == string.Empty && detailText == string.Empty)
                {
                    System.Console.WriteLine();
                }
                else
                {
                    System.Console.Write(padRight(cmdText, 15));
                    ignoreTab = false;
                }
            }
            else
            {
                System.Console.WriteLine(cmdText);
                newLine = true;
            }
            if (detailText != string.Empty)
            {
                if (!newLine)
                {
                    wrapText(detailText, 25, true);
                }
                else
                {
                    wrapText(detailText, 25, false);
                }
            }
        }

        /// <summary>
        /// Displays the basic syntax of the utility
        /// </summary>
        private static void QuickHelp()
        {
            ScreenWriter("mt v. 1.0");
            ScreenWriter("usage: mtape [-v] [--version] [-h] [--help] [-f device] command [count | sourcefile | sourcefile targetfile]");
            ScreenWriter("commands:", "weof, wset, eof, fsf, fsfm, bsf, bsfm, fsr, bsr, fss, bss, rewind,");
            ScreenWriter("", "offline, rewoffl, eject, retention, eod, seod, seeka, seekl, tell, status,");
            ScreenWriter("", "erase, setblk, lock, unlock, load, compression, setdensity,");
            ScreenWriter("", "drvbuffer, stwrthreshold, stoptions, stsetoptions, stclearoptions,");
            ScreenWriter("", "defblksize, defdensity, defdrvbuffer, defcompression, stsetcln,");
            ScreenWriter("", "sttimeout, stlongtimeout, densities, setpartition, mkpartition");
            ScreenWriter("", "partseek, asf, stshowoptions, locate, read, write, continue");
            ScreenWriter("", "resume, writelist, info, init.");
        }

        private static void testMessage()
        {
            ScreenWriter("DESCRIPTION", "This manual page documents the tape control program mtape.  mtape performs the given operation, which must be one of the tape operations listed below, on a tape drive.  The commands can also be listed by running the program with the -h option.");

        }
        /// <summary>
        /// Print detailed help on all commands and syntax
        /// </summary>
        private static void DetailedHelp()
        {
            // Need to change from System.Console.Writeline to a custom writeline that supports recognizing screen widths and wrapping accordingly.
            ScreenWriter("MTAPE(1)");
            ScreenWriter("");
            ScreenWriter("NAME", "mtape - control magnetic tape drive operation");
            ScreenWriter("");
            ScreenWriter("SYNOPSIS", "mtape [-h] [-f device] operation [count] [arguments...]");
            ScreenWriter("");
            ScreenWriter("DESCRIPTION", "This manual page documents the tape control program mtape.  mtape performs the given operation, which must be one of the tape operations listed below, on a tape drive.  The commands can also be listed by running the program with the -h option.");
            ScreenWriter("");
            ScreenWriter("", "Some operations optionally take an argument or request count, which can be given after the operation name and defaults to 1.  The postfix k, M, or G can be used to give counts in units of 1024, 1024 * 1024, or 1024 * 1024 * 1024, respectively.");
            ScreenWriter("");
            ScreenWriter("", "The available operations are listed below.  Unique abbreviates are accepted.  Not all operations are avilable on all systems, or work on all types of tape drives.");
            ScreenWriter("");
            ScreenWriter("", "fsf", "Forward space count files.  The tape is positioned on the first block of the next file.");
            ScreenWriter("");
            ScreenWriter("", "fsfm", "Forward space count files, then backward space one record.  This leaves the tape positioned at the last block of the file that is count - 1 files past the current file.");
            ScreenWriter("");
            ScreenWriter("", "bsf", "Backward space count files.  The tape is positioned on the last block of the previous file.");
            ScreenWriter("");
            ScreenWriter("", "bsfm", "Backward space count files, then forward space one record.  This leaves the tape positioned at the first block of the file that is count - 1 files before the current file.");

            ScreenWriter("");
            ScreenWriter("", "asf", "The tape is positioned at the beginning of the count file.  Positioning is done by first rewinding the tape and then spacing forward over count filemarks.");
            ScreenWriter("");
            ScreenWriter("", "fsr", "Forward space count records.");
            ScreenWriter("");
            ScreenWriter("", "bsr", "Backward space count records.");
            ScreenWriter("");
            ScreenWriter("", "fss", "(SCSI tapes) Forward space count setmarks.");
            ScreenWriter("");
            ScreenWriter("", "bss", "(SCSI tapes) Backward space count setmarks.");
            ScreenWriter("");
            ScreenWriter("", "eod, seod", "Space to end of valid data.  Used on streamer tape drives to append data to the logical end of tape.");

            ScreenWriter("");
            ScreenWriter("", "rewind", "Rewind the tape.");
            ScreenWriter("");
            ScreenWriter("", "offline, rewoffl, eject", "Rewind the tape and, if appicable, unload the tape.");
            ScreenWriter("");
            ScreenWriter("", "retention", "Rewind the tape, then wind it to the end of the reel, then rewind it again.");
            ScreenWriter("");
            ScreenWriter("", "weof, eof", "Write count EOF marks at current position.");
            ScreenWriter("");
            ScreenWriter("", "wset", "(SCSI tapes) Write count setmarks at current position (only SCSI tape).");
            ScreenWriter("");
            ScreenWriter("", "erase", "Erase the tape.  Note that this is a long erase, which on modern (high-capacity) tapes can take many hours, and which usually can't be aborted");
            ScreenWriter("");
            ScreenWriter("", "status", "Print status information about the tape unit.  (f the density code is \"no translation\" in the status output, this does not affect working of the tape drive.");

            ScreenWriter("");
            ScreenWriter("", "seeka", "(SCSI tapes) Seek to the count absolute location on the tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.  The block address should be obtained from a tell call earlier.");
            ScreenWriter("");
            ScreenWriter("", "seekl", "(SCSI tapes) Seek to the count logical location on the tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.  The block address should be obtained from a tell call earlier.");

            ScreenWriter("");
            ScreenWriter("", "tell", "(SCSI tapes) Tell the current block on tape.  This operation is available on some Tandberg and Wangtek streamers and some SCSI-2 tape drives.");
            ScreenWriter("");
            ScreenWriter("", "setpartition", "(SCSI tapes) Switch to the partition determined by count.  The default data partition of the tape is numbered zero.  Switching partition is available only if enabled for the device, the device supports multiple partitions, and the tape is formatted with multiple partitions.");

            ScreenWriter("");
            ScreenWriter("", "partseek", "(SCSI tapes) The tape position is set to block count in the parition given by the argument after count.  The default parition is zero.");

            ScreenWriter("");
            ScreenWriter("", "mkpartition", "(SCSI tapes) Format the tape with one (count is zero) or two partitions (count gives the size of the second partition in megabytes). If the count is positive, it specifies the size of partition 1.  From kernel version 4.6, if the count is negative, it specifies the size of partition 0.  With older kernels, a negative argument formats the tape with one partition.  The tape drive must be able to format partitioned tapes with initiator-specified partition size and partition support must be enabled for the drive.");
            ScreenWriter("");
            ScreenWriter("", "load", "(SCSI tapes) Send the load command to the tape drive.  The drives usually load the tape when a new cartridge is inserted.  The argument count can usually be omitted.  Some HP changers load tape n if the count 10000 + n is given (a special function in the Linux st driver).");

            ScreenWriter("");
            ScreenWriter("", "lock", "(SCSI tapes) Lock the tape drive door.");
            ScreenWriter("");
            ScreenWriter("", "unlock", "(SCSI tapes) Unlock the tape drive door.");
            ScreenWriter("");
            ScreenWriter("", "setblk", "(SCSI tapes) Set the block size of the drive to count bytes per record");
            ScreenWriter("");
            ScreenWriter("", "setdensity", "(SCSI tapes) Set the tape density code to count.  The proper codes to use with each drive should be looked up from the drive documentation.");
            ScreenWriter("");
            ScreenWriter("", "densities", "(SCSI tapes) Write explanation of some common desnity codes to standard output.");

            ScreenWriter("");
            ScreenWriter("", "drvbuffer", "(SCSI tapes) Set the tape drive buffer code to number.  The proper value for unbuffered operation is zero and \"normal\" buffered operation one.  The meanings of other values can be found in the drive documentation or, in the case of a SCSI-2 drive, from the SCSI-2 standard.");

            ScreenWriter("", "compression", "(SCSI tapes) The compression within the drive can be switched on or off using the MTCOMPRESSION ioctl.  Note that this method is not supported by all drives implementing compression.  For instance, the Exabyte 8 mm drivers use density codes to select compression.");
            ScreenWriter("");
            ScreenWriter("", "stoptions", "(SCSI tapes) Set the driver options bits for the device to the defined values.  Allowed only for the superuser.  The bits can be set either by ORing the option bits from the file /usr/include/linux/mtio.h to count, or by using the following keywords (as many keywords can be used on the same line as necessary, unambiguous abbreviations allowed):");
            ScreenWriter("");
            ScreenWriter("", " ", "     buffer-writes   buffered writes enabled");
            ScreenWriter("");
            ScreenWriter("", " ", "     async-writes    asynchronous writes enabled");
            ScreenWriter("");
            ScreenWriter("", " ", "     read-ahead      read-ahead for fixed block size");
            ScreenWriter("");
            ScreenWriter("", " ", "     debug           debugging (if compiled into driver)");
            ScreenWriter("");
            ScreenWriter("", " ", "     two-fms         write two filemarks when file closed");
            ScreenWriter("");
            ScreenWriter("", " ", "     fast-eod        space directly to eod (and lose file number)");
            ScreenWriter("");
            ScreenWriter("", " ", "     no-wait         don't wait until rewind, etc. complete");
            ScreenWriter("");
            ScreenWriter("", " ", "     auto-lock       automatically lock/unlock drive door");
            ScreenWriter("");
            ScreenWriter("", " ", "     def-writes      the block size and desntity are for writes");
            ScreenWriter("");
            ScreenWriter("", " ", "     can-bsr         drive can space backwards as well");
            ScreenWriter("");
            ScreenWriter("", " ", "     no-blklimits    drive doesn't support read block limits");
            ScreenWriter("");
            ScreenWriter("", " ", "     can-partitions  drive can handle partitioned tapes");
            ScreenWriter("");
            ScreenWriter("", " ", "     scsi2logical    seek and tell use SCSI-2 logical block addresses instead of device dependent addresses");
            ScreenWriter("");
            ScreenWriter("", " ", "     sili            Set the SILI bit is when reading in variable block mode.  This may speed up reading blocks shorter than the read byte count.  Set this option only if you know that the drive supports SILI and the HBA reliably returns transfer residual byte counts.  Requires kernel version >= 2.6.26.");
            ScreenWriter("");
            ScreenWriter("", " ", "     sysv            Enable the System V semantics");
            ScreenWriter("");
            ScreenWriter("", "stsetoptions", "(SCSI tapes) Set selected driver options bits.  The methods to specify the bits to set are given above in the description of stoptions.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "stclearoptions", "(SCSI tapes) Clear selected driver option bits.  The methods to specify the bits to clear are given above in deescription of stoptions.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "stshowoptions", "(SCSI tapes) Print the currently enabled options for the device.  Requires kernel version >= 2.6.26 and sysfs must be moutned at /sys.");
            ScreenWriter("");
            ScreenWriter("", "stwrthreshold", "(SCSI tapes) The write threshold for the tape device is set to count kilobytes.  The value must be smaller than or equal to the driver buffer size.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "defblksize", "(SCSI tapes) Set the default block size of the device to count ytes.  The value -1 disables the default block size.  The block size set by setblk overrides the default until a new tape is inserted.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "defdensity", "(SCSI tapes) Set the default density code.  The value -1 disables the default desnity.  The desnity set by setdensity overrides the default until a new tape is inserted.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "defdrvbuffer", "(SCSI tapes) Set the default drive buffer code.  The value -1 disables the default drive buffer code.  The drive buffer code set by drvbuffer overrides the default until a new tape is inserted.  Alowed only for the superuser");
            ScreenWriter("");
            ScreenWriter("", "defcompression", "(SCSI tapes) Set the default ompression state.  The value -1 disables the default compression.  The compression state set by compression over rides the default until a new tape is inserted.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "sttimeout", "sets the normal timeout for the device.  The value is given in seconds.  Allowed only for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "stlongtimeout", "sets the long timeout for the device.  The value is given in seconds.  Allowedo nly for the superuser.");
            ScreenWriter("");
            ScreenWriter("", "stsetcln", "set the cleaning request interpretation parameters.");
            ScreenWriter("");
            ScreenWriter("", "read", "reads a file from tape and optionally writes to specified file.  Default writes to stdout.");
            ScreenWriter("");
            ScreenWriter("", "write", "writes specified file to tape.  If a location is specifed, it mtape will advance to that filemark before writing.");
            ScreenWriter("");
            ScreenWriter("", "writelist", "writes files from specified file list to tape. If a location is specifed, it mtape will advance to that filemark before writing.");
            ScreenWriter("");
            ScreenWriter("", "locate", "Locates file in the tape library, requests the appropriate volume, seeks to the appropriate location on tape and restores the the specified target file name.");
            ScreenWriter("");
            ScreenWriter("", "info", "mtape prints a list of tape and device parameters");
            ScreenWriter("");
            ScreenWriter(" ", "init", "Initialize tape library.");
            ScreenWriter(" ");
            ScreenWriter("", "append", "Appends one or more files to the current volume.");
            ScreenWriter("");
            ScreenWriter("", "resume", "Resumes writing to the current volume by identifying what was scheduled to be written but did not complete and then appending.");
            ScreenWriter("");
            ScreenWriter("", "", "mtape exits with a status of 0 if the operation succeeded, 1 if the operation or device name given was invalid, or 2 if the operation failed.");
            ScreenWriter("");
            ScreenWriter("OPTIONS");
            ScreenWriter("", "-h, --help", "Print a usage message on standard output and exit successfully.");
            ScreenWriter("");
            ScreenWriter("", "-v, --version", "Prints the current mtape version number.");
            ScreenWriter("");
            ScreenWriter("", "-f, -t", "The path of the tape device on which to operate.  If neither of those options is given, and the environment variable TAPE is set, it is used.  Otherwise, a default device defined in the file /usr/include/sys/mtio.h is used (note that the actual path to mtio.h can vary per architecture and/or distribution).");
            ScreenWriter("");
            ScreenWriter("NOTES", "", "The argument of mkpartition specifies the size of the partition in megabytes.  If you add a postfix, it applies to this definition.  For example, argument 1G means 1 giga megabytes, which probably is not what the user is anticipating.");
            ScreenWriter("");
            ScreenWriter("AUTHOR", "", "The program is written by Phillip Jones (based on Kai Makisara's mt), and is currently maintained by Phillip Jones<jones.phillip.a@gmail.com>.");
            ScreenWriter("");
            ScreenWriter("COPYRIGHT", "", "The program and the manual page are copyrighted by Phillip Jones, 2020-.  They can be distributed according to the GNU Copyleft.");
            ScreenWriter("");
            ScreenWriter("BUGS", "", "Please report bugs to <jones.phillip.a@gmail.com>.");
            ScreenWriter("");
            ScreenWriter("SEE ALSO", "", "");
            ScreenWriter("");
            // Print a usage message on standard output and exit successfully.

        }
    }
}
