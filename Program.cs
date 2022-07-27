using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    internal class Program
    {
        static public int delaySeconds = 30;
        static public string Tempfolder = "d:\\temp";
        static public int opsticalMinutes = 30;
        static public bool debugMode = true;
        static public int Pop3Port = 0;
        static public string containerFolder;

        static void Main(string[] args)
        {
            int mailcount = 0;

            if (args.Length == 0)
            {
                Console.WriteLine("Invalid args");
                return;
            }
            delaySeconds = Convert.ToInt32(args[0]);
            Tempfolder = args[1];
            opsticalMinutes = Convert.ToInt32(args[2]);
            debugMode = Convert.ToBoolean(args[3]);
            Pop3Port = Convert.ToInt32(args[4]);
            if ( args.Length > 5)
            {
                containerFolder = args[5];
            }
            else
            {
                containerFolder = "";
            }


            ClassMailManager mailManager = new ClassMailManager(debugMode, Pop3Port);

            while (true)
            {
                try
                {
                    mailcount = mailManager.getNumberOfmessages();
                    ////
                    if (mailcount > 0)
                    {
                        int firstMail = 1;
                        int count = mailManager.checkMailContent  (firstMail, Tempfolder, debugMode, opsticalMinutes );
                    }



                    ///////
                }
                catch (Exception ee)
                {
 //                   Log.Info("exception while main try " + ee.ToString());
                }
                Thread.Sleep(delaySeconds*1000);                
            }

               

        }
    }
}
