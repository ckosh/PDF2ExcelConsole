using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    internal class ClassEnvironmentParams
    {
        int delaySeconds = 30;
        string tempfolder = "d:\\temp";
        int opsticalMinutes = 30;
        bool debugMode = true;
        int pop3Port = 110;
        string resultExcelFile;
        bool skipMail = false;
        bool ssl = false;
        string backupFolder = "";
        ClassArgsParser argParser;
        bool printDebug = true;

        public ClassEnvironmentParams(string fileName)
        {
            //
            //
            //ComputerName = "pdf2excel";
            //
            //

            argParser = new ClassArgsParser();
            argParser.parseArgs(fileName);
            delaySeconds = argParser.getDelaySeconds();
            tempfolder = argParser.getTempFolder();
            debugMode = argParser.getDebugMode();
            pop3Port = argParser.getPort();
            skipMail = argParser.getSkipMail();
            ssl = argParser.getussl();
            backupFolder = argParser.getBackupFolder();

            if (printDebug)
            {
                Console.WriteLine("Delay " + delaySeconds);
                Console.WriteLine("temp folder " + tempfolder);
                Console.WriteLine("debug mode " + debugMode);
                Console.WriteLine("Pop " + pop3Port);
                Console.WriteLine("skip mail " + skipMail);
                Console.WriteLine("ssl " + ssl);
                Console.WriteLine("backup folder " + backupFolder);        
            }
        }
        public int DelaySeconds()
        {
            return delaySeconds;
        }
        public string Tempfolder()
        {
            return tempfolder;
        }
        public bool DebugMode()
        {
            return debugMode;
        }
        public int Pop3Port()
        {
            return pop3Port;
        }
        public bool SkipMail()
        {
            return skipMail;
        }
        public bool Ssl()
        {
            return ssl;
        }
        public string BackupFolder()
        {
            return backupFolder;
        }
    }

}
