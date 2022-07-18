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

            ClassMailManager mailManager = new ClassMailManager(debugMode);

            while (true)
            {
                try
                {
                    mailcount = mailManager.getNumberOfmessages();
                    ////
                    if (mailcount > 0)
                    {
                        int firstMail = 1;
                        int count = mailManager.checkMailContent  (firstMail, Tempfolder, debugMode, opsticalMinutes);
                        //count = count + Int32.Parse(labelsent.Text);
                        //labelsent.Text = count.ToString();
                        //labelsent.Refresh();

                        //int owners = batchClass.getNumberOfOwners();
                        //owners = owners + Int32.Parse(labelowners.Text);
                        //labelowners.Text = owners.ToString();
                        //labelowners.Refresh();
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
