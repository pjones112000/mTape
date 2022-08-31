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

namespace mtape
{
    class MTape2Script
    {

        public const int NewTape = 0;
        public const int StartBackup = 1;
        public const int EndBackup = 2;
        public const int Alert = 3;
        public const int TapeNeeded = 4;
        public const int Clean = 5;

        private string m_LogFile = string.Empty;
        private string m_ErrorLog = string.Empty;
        private long m_CurrentPosition = -1;

        private int m_VolumeNumber;
        private string m_Message = string.Empty;
        private string m_tapeDrive = string.Empty;



        private Dictionary<String, List<String>> mFunctions = new Dictionary<string, List<string>>();
        private Dictionary<String, List<String>> mFunctParams = new Dictionary<string, List<string>>();


        public void SetLogFile(String LogFile)
        {
            m_LogFile = LogFile;
        }

        public void SetErrorLog(String LogFile)
        {
            m_ErrorLog = LogFile;
        }

        public void Init(int VolumeNumber, int ReasonCode, String Message = "", String LogFile = "", String ErrorLog = "")
        {
            m_VolumeNumber = VolumeNumber;
            m_Message = Message;
            mFunctions = new Dictionary<string, List<string>>(0);
            mFunctParams = new Dictionary<string, List<string>>();
            if (LogFile != string.Empty)
            {
                m_LogFile = LogFile;
            }
            if (ErrorLog != string.Empty)
            {
                m_ErrorLog = ErrorLog;
            }
        }

        public string TapeDrive
        {
            get { return m_tapeDrive; }
            set
            {
                m_tapeDrive = value;
                API.TapeDrive = m_tapeDrive;
            }
        }

        /// <summary>
        /// Removes double quotes from values
        /// </summary>
        /// <param name="Dictionary&lt;String,String&gt;source"></param>
        private void cleanVariables(ref Dictionary<String, String> source)
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

        private TapeWinAPI API = new TapeWinAPI();

        /// <summary>
        /// Sends an SMTP email
        /// </summary>
        /// <param name="Dictionary&lt;String,String&gt;mVariables"></param>
        /// <returns>bool Success</returns>
        private bool SendMail(Dictionary<String, String> mVariables)
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
                if (mVariables["Reason"] == "NewTape")
                {
                    tmpMessage = tmpMessage.Replace("$Volume", (m_VolumeNumber + 1).ToString()).Replace("$VOLUME", (m_VolumeNumber + 1).ToString()).Replace("$REASON", "Newtape");
                }
                else
                {
                    tmpMessage = tmpMessage.Replace("$Volume", m_VolumeNumber.ToString()).Replace("$VOLUME", m_VolumeNumber.ToString()).Replace("$REASON", "Needed");
                }
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
        /// Logs an error.
        /// </summary>
        /// <param name="String Command">The command issuing the error</param>
        /// <param name="String Error" type="string">The comment to add to the error log</param>
        private void LogError(String Command, String Error)
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
        private void LogMessage(String Command, String Message)
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

        private uint m_CurrentError = 0;
        private long getCurrentPosition()
        {
            long retValue = -1;
            //            bool tapeOpen = false;
            System.Console.WriteLine("Tapedrive:" + API.TapeDrive);
            if (!API.IsTapeOpen())
            {
                if (!API.Open(@"\\.\" + API.TapeDrive))
                {
                    System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                    LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                }
                else
                {
                    System.Console.WriteLine("Failed to open tape drive");
                }
            }

            if (API.IsTapeOpen())
            {
                System.Console.WriteLine("Before Handle:" + API.GetHandleValid());
                long position = 0;
                uint retCode = 0;
                
                if (API.GetTapePosition(ref position, ref retCode))
                {
                    m_CurrentPosition = position;
                    retValue = m_CurrentPosition;
                }
                else
                {
                    retValue = -1;
                    m_CurrentPosition = -1;
                }
                m_CurrentError = retCode;
                System.Console.WriteLine("After Handle:" + API.GetHandleValid());
            }
            return retValue;
        }


        /// <summary>
        /// Keeps checking to see if the tape drive is read (i.e. tape loaded and ready to write)
        /// </summary>
        /// <returns>bool Success</returns>
        private bool DetectNewTape()
        {
            bool retValue = false;
            bool tapeDetected = false;

            while (!tapeDetected)
            {

                if (!API.IsTapeOpen())
                {
                    if (!API.Open(@"\\.\" + API.TapeDrive))
                    {
                        System.Console.WriteLine("Tape drive:" + @"\\.\" + API.TapeDrive);
                        LogError("Open", new Win32Exception((int)(uint)Marshal.GetLastWin32Error()).Message);
                    }
                }
                else
                {
                    try
                    {
                        uint retCode = 0;
                        if (getCurrentPosition() != -1)
                        {
                            tapeDetected = true;
                        }
                        else
                        {
                            tapeDetected = false;
                            System.Console.WriteLine("Error:" + m_CurrentError);
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                System.Threading.Thread.Sleep(500);
            }
            return retValue;
        }

        /// <summary>
        /// Capture stdError output from external batch or executable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void p_ErrorDataReceived(object sender, DataReceivedEventArgs e)
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
        private void p_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Process p = sender as Process;
            if (p == null)
                return;
            Console.WriteLine(e.Data);
        }


        /// <summary>
        /// Executes an external batch file or executable
        /// </summary>
        /// <param name="ref int RetCode"></param>
        /// <param name="String Method"></param>
        /// <returns>bool Success</returns>
        private bool RunExternal(out int RetCode, String Method = "", bool newTape = false)
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
                        p.StandardInput.Write("[RunExternal] " + m_VolumeNumber + 1 + " New" + p.StandardInput.NewLine);
                    }
                    else
                    {
                        p.StandardInput.Write("[RunExternal] " + m_VolumeNumber + " Needed" + p.StandardInput.NewLine);
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
                System.Console.WriteLine("[CallScript] Error executing script " + ex.Message);
                RetCode = ex.HResult;
                return false;
            }
        }


        /// <summary>
        /// Eject the volume from the tape drive
        /// </summary>
        private void EJECT()
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



        private int scriptProcessor(String funcName, List<String> scriptLines, int ReasonCode, List<string> Parameters = null)
        {
            int retValue = -1;
            int retCode = 0;
            String regExPattern = "[A-Z0-9$\\(\\)\\\"\\,]+| (\\\"([^\\\"])*\\\")|(^[A-Z0-9\\s\\=\\\"]*)";
            Regex rg = new Regex(regExPattern, RegexOptions.IgnoreCase);
            Dictionary<String, String> mVariables = new Dictionary<string, string>();

            if (mFunctParams.ContainsKey(funcName))
            {
                int parmNum = 0;
                foreach (String parameterName in mFunctParams[funcName])
                {
                    if (Parameters != null)
                    {
                        mVariables.Add("$" + parameterName.Replace("\"", "").ToUpper(), Parameters[parmNum]);
                    }
                    parmNum++;
                }
            }

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

            mVariables.Add("$VOLUME", "VOL_" + (m_VolumeNumber - 1).ToString());
            mVariables.Add("$REASON", Reason);
            mVariables.Add("$DATETIME", DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
            mVariables.Add("$1", Reason);
            mVariables.Add("$2", "VOL_" + (m_VolumeNumber).ToString());
            mVariables.Add("$0", "VOL_" + (m_VolumeNumber).ToString());
            mVariables.Add("$3", m_Message);
            int i = 1;

            String switchText = String.Empty;

            bool inSwitch = false;
            bool caseTrue = false;

            foreach (String tmpText in scriptLines)
            {
                String newText = tmpText;

                foreach (String varName in mVariables.Keys)
                {
                    newText = newText.Replace(varName.ToUpper(), mVariables[varName]);
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
                                    retCode = retCode + scriptProcessor(tmpFunc, mFunctions[tmpFunc], ReasonCode);
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
                                            retCode = retCode + scriptProcessor(workString.Substring(0, workString.IndexOf("(")), mFunctions[workString.Substring(0, workString.IndexOf("("))], ReasonCode, mParameters);
                                        }
                                        else
                                        {
                                            switch (matches[0].Value.ToLower())
                                            {
                                                case "eject()":
                                                    // Commented out for testing purposes
                                                    EJECT();
                                                    break;
                                                case "keypress()":
                                                    System.Console.ReadKey();
                                                    break;
                                                case "detectnewtape()":
                                                    System.Console.WriteLine("[" + funcName + "] Detecting new tape in drive.");
                                                    if (!DetectNewTape())
                                                    {
                                                        return 1;
                                                    }
                                                    // tape detection 
                                                    break;
                                                case "sendmail()":
                                                    System.Console.WriteLine("[" + funcName + "] Sending email");
                                                    if (ReasonCode == NewTape)
                                                    {
                                                        mVariables.Add("Reason", "NewTape");
                                                    }
                                                    else
                                                    {
                                                        mVariables.Add("Reason", "Needed");
                                                    }
                                                    if (!SendMail(mVariables))
                                                    {
                                                        return 1;
                                                    }
                                                    // send mail
                                                    break;
                                                case "print":
                                                    if (ReasonCode == NewTape)
                                                    {
                                                        System.Console.WriteLine("[" + funcName + "] " + matches[1].Value);
                                                    }
                                                    else
                                                    {
                                                        System.Console.WriteLine("[" + funcName + "] " + matches[1].Value);
                                                    }
                                                    break;
                                                case "call":
                                                    // Calls external application or script
                                                    if (!RunExternal(out retCode, " + funcName + "))
                                                    {
                                                        System.Console.WriteLine("[" + funcName + "] External command reported an error :" + retCode);
                                                        return 1;
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
                                                        System.Console.WriteLine("[" + funcName + "] Unrecognized command line " + tmpText);
                                                    }
                                                    break;
                                            }
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
                                    System.Console.WriteLine("[" + funcName + "] Unmatched command line '" + newText + "'");
                                }
                            }
                            else
                            {
                                string[] splitted = newText.Split('=');

                                if (!mVariables.ContainsKey(splitted[0]))
                                {
                                    mVariables.Add(splitted[0], splitted[1].Replace("\"", ""));
                                }
                                else
                                {
                                    mVariables[splitted[0]] = splitted[1].Replace("\"", "");
                                }

                            }
                        }
                    }
                }
            }

            return retValue;
        }

        

        /// <summary>
        /// Interpret an mTape script
        /// </summary>
        /// <returns>bool Success</returns>
        public bool runMTapeScript(int ReasonCode, String Script, out int returnCode, String Message = "")
        {
            bool retValue = false;
            int retCode = 0;
            string[] text = System.IO.File.ReadAllLines(Script);

            m_Message = Message;
            String funPattern = "[A-Z0-9$\\(\\)\\\"\\,]+| (\\\"([^\\\"])*\\\")";

            List<String> functionLines = new List<string>();
            bool ReadFunction = false;
            String functionName = string.Empty;
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
                                mFunctParams.Add(functionName, tParmNames);
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
                        mFunctions.Add(functionName, functionLines);
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
                mFunctions.Add("main", functionLines);
                if (mFunctions.ContainsKey("main"))
                {
                    retCode = scriptProcessor("main", mFunctions["main"], ReasonCode);
                }
                else
                {
                    retCode = 0;
                }
            }
            returnCode = retCode;
            return retValue;
        }

    }
}
