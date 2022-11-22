using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;


namespace Logger
{
   public  class Logger
    {
        #region Members
        public static int lineCounter = 1;
        public static bool logged = false;
        public static StringBuilder message = new StringBuilder();
        #endregion

        #region Methods
        public static void ClearLog()
        {
            logged = false;
            lineCounter = 1; //Reset counter
            message.Clear();
        }
        public static void Log(string msg)
        {

            message.AppendLine(lineCounter + ") " + msg + "\n");
            lineCounter++;
            logged = true;
        }
        public static bool CreateConnectionErrorLog(string logPath)
        {
            string fileName = "Connection Error";
            if (logged) //Checks if there have been any log
            {
                string logFileName = "\\log_" + DateTime.Now.ToString("yyyyMMdd-hhmm") + "_" + fileName + ".txt"; //build the name of the log File

                File.AppendAllText(logPath + logFileName, message.ToString()); //Create the log file               

                return true;
            }
            else
            {
                return false;
            }
        }
        private static string _CurrentLogFileName;
        public static String CurrentLogFileName
        {
            get { return _CurrentLogFileName; }
        }
        public static bool CreateLog(string logPath, string fileName)
        {
            DeleteOldFiles(logPath);
            if (logged) //Checks if there have been any log
            {
                fileName = fileName.Substring(0, fileName.Length - 4); //remove ".txt" from the name                
                string logFileName = "\\log_" + DateTime.Now.ToString("yyyyMMdd-hhmm") + "_" + fileName + ".txt"; //build the name of the log File
                _CurrentLogFileName = logPath + logFileName;
                File.AppendAllText(logPath + logFileName, message.ToString()); //Create the log file               
                if (System.Diagnostics.Debugger.IsAttached)
                    try { Process.Start(logPath + logFileName); }
                    catch { }
                return true;
            }
            else
            {
                return false;
            }
        }
        static void DeleteOldFiles(string Directory)
        {
            System.IO.Directory.GetFiles(Directory).Select(f => new FileInfo(f))
          .Where(f => f.LastWriteTime.AddDays(100).Date < DateTime.Now.Date)
          .ToList()
          .ForEach(f => f.Delete());
        }
        #endregion
    }
}
