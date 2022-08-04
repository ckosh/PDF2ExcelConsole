using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    public class ClassArgsParser
    {
        public int delaySeconds;
        public string Tempfolder;
        public bool debugMode;
        public int Pop3Port;
        public bool skipMail;
        public bool useSsl;
        string[] argsKeys = new string[6];

        public ClassArgsParser()
        {
            delaySeconds = 30;
            Tempfolder = "d:\\temp";
            debugMode = false;
            Pop3Port = 110;
            useSsl = false;
            skipMail = false;

            argsKeys[0] = "delay";
            argsKeys[1] = "tempfolder";
            argsKeys[2] = "debug";
            argsKeys[3] = "port";
            argsKeys[4] = "ssl";
            argsKeys[5] = "skipmail";
        }

        public void parseArgs(string argsFile)
        {
            var lines = File.ReadAllLines(argsFile);
            for (var i = 0; i < lines.Length; i += 1)
            {
                var line = lines[i];
                string[] subs = line.Split('=');
                subs[0] = String.Concat(subs[0].Where(c => !Char.IsWhiteSpace(c)));
                subs[1] = String.Concat(subs[1].Where(c => !Char.IsWhiteSpace(c)));

                if (subs[0] == argsKeys[0])
                {
                    delaySeconds = Convert.ToInt32(subs[1]);
                    continue;
                }
                else if (subs[0] == argsKeys[1])
                {
                    Tempfolder = subs[1];
                    continue;
                }
                else if (subs[0] == argsKeys[2])
                {
                    debugMode = Convert.ToBoolean(subs[1]);
                    continue;
                }
                else if(subs[0] == argsKeys[3])
                {
                    Pop3Port = Convert.ToInt32(subs[1]);
                    continue;
                }
                else if (subs[0] == argsKeys[4])
                {
                    useSsl = Convert.ToBoolean(subs[1]);
                    continue;
                }
                else if (subs[0] == argsKeys[5])
                {
                    skipMail = Convert.ToBoolean(subs[1]);
                    continue;
                }
            }

        }
        public int getDelaySeconds()
        {
            return delaySeconds;
        }
        public string getTempFolder()
        {
            return Tempfolder;
        }
        public bool getDebugMode()
        {
            return debugMode;
        }
        public int getPort()
        {
            return Pop3Port;
        }
        public bool getSkipMail()
        {
            return skipMail;
        }
        public bool getussl()
        {
            return useSsl;
        }
    }
}
