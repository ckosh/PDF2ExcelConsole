using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    internal class Program
    {
        //static public int delaySeconds = 30;
        //static public string Tempfolder = "d:\\temp";
        //static public int opsticalMinutes = 30;
        //static public bool debugMode = true;
        //static public int Pop3Port =  110;
        static public string resultExcelFile;
        //static public bool skipMail = false;
        //static public bool ssl = false;
        //static public string backupFolder = "";
        static bool printDebug = true;

        static ClassEnvironmentParams environmentParams;
        static void Main()
         {    
            string userMail;
            List<int> listofOwners;
            string argfileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\PDF2Excel\\PDF2ExcelArgs.txt";
            environmentParams = new ClassEnvironmentParams(argfileName);
            int numberOfOwners;

            Console.WriteLine("Start 23/01/2023 fix zhuyit leasing");
            

            //if (args.Length == 0)
            //{
            //    environmentParams = new ClassEnvironmentParams("");
            //    if (!environmentParams.azureVMMode())
            //    {
            //        Console.WriteLine("Invalid args");
            //        return;
            //    }
            //}
            //else
            //{
            //    environmentParams = new ClassEnvironmentParams(args[0]);
            //}
            
            {
//                environmentParams = new ClassEnvironmentParams(args[0]);
                //string ParamFile = args[0];
                //ClassArgsParser argParser = new ClassArgsParser();
                //argParser.parseArgs(ParamFile);

                //delaySeconds = argParser.getDelaySeconds();
                //Tempfolder = argParser.getTempFolder();
                //debugMode = argParser.getDebugMode();
                //Pop3Port = argParser.getPort();
                //skipMail = argParser.getSkipMail();
                //ssl = argParser.getussl();
                //backupFolder = argParser.getBackupFolder();


                if (environmentParams.SkipMail() )
                {
                    string[] temp = Directory.GetFiles(environmentParams.Tempfolder(), "*.pdf");
                    string[] PdfFileNames = new string[temp.Length];
                    for (int k = 0; k < temp.Length; k++)
                    {
                        PdfFileNames[k] = Path.GetFileName(temp[k]);
                    }
                    ClassProcessPDF processPDF = new ClassProcessPDF(PdfFileNames, environmentParams.DebugMode(), environmentParams.Tempfolder());
                    resultExcelFile = processPDF.convert();

                }
                else
                {
                    ClassMailManager1 mailManager = new ClassMailManager1(environmentParams.DebugMode(), environmentParams.Pop3Port() , environmentParams.Ssl(), environmentParams.Tempfolder() );
                    
                    while (true)
                    {
                        userMail = mailManager.ConnectToPop3();
                        if (userMail != null && userMail.Length > 1 )
                        {
                            
                            if (userMail.Length > 0)
                            {

                                Console.WriteLine(userMail);
                                //                    Log.Info("before process files  ");
                                string[] PdfFileNames = mailManager.GetPdfFiles();
                                ClassProcessPDF processPDF = new ClassProcessPDF(PdfFileNames, environmentParams.DebugMode(), environmentParams.Tempfolder());
                                //            Log.Info("after process files  ");
                                resultExcelFile = processPDF.convert();
                                Console.WriteLine(resultExcelFile);
                                //            Log.Info("after convert files  ");

                                listofOwners = processPDF.getTotalNumberOfOwners();
                                numberOfOwners = listofOwners.Sum();
                                Console.WriteLine("owners " + numberOfOwners.ToString());
                                Console.WriteLine("files " + listofOwners.Count.ToString());

                                double totalcost = getTotalCost(listofOwners);
                                string ffff = totalcost.ToString("0.#");
                                string body = ffff + " עלות הסבה" + '\n' + PdfFileNames.Length.ToString() + " מספר נסחים " + '\n' + numberOfOwners.ToString() + " מספר בעלים " + '\n' ;
                                // copy for backup
//                                mailManager.sendMail("grabnadlan@gmail.com", "העתק תוצאות", resultExcelFile, userMail + '\n' + body);
                                mailManager.sendMail(userMail, "תוצאות הסבת נסחי טאבו", resultExcelFile, body);
                                if (environmentParams.BackupFolder()  != ""  || userMail != "chaim.koshizky@gmail.com")
                                {
                                    string Todir = environmentParams.BackupFolder() + "\\" + userMail + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                                    ClassUtils.copyFilesFromDirToDir(environmentParams.Tempfolder() , Todir);
                                }
                                ClassUtils.deleteAllFilesFromDirectory(environmentParams.Tempfolder());
                                ClassUtils.deleteAllFilesFromDirectory(environmentParams.Tempfolder() + "\\CSV");
                                if (userMail != "chaim.koshizky@gmail.com")
                                {
                                    mailManager.savecBillingData(userMail, listofOwners.Count, resultExcelFile, numberOfOwners, totalcost, mailManager.getSubject());
                                }
                            }
                        }
                        Thread.Sleep(environmentParams.DelaySeconds()  * 1000);
                    }
                }
            }
            double getTotalCost(List<int> owners)
            {
                double ret = 0.0;
                double basecost = owners.Count * 5.0; // number of files * 5 nis
                double ownercost = listofOwners.Sum() * 0.2; // number of owners * 0.2 nis
                if (ownercost > basecost)
                {
                    ret = ownercost;
                }
                else
                {
                    ret = basecost;
                }
                return ret;
            }

            //double getTotalCost(int NumberOfPDFFiles)
            //{
            //    double totalCost = 0;
            //    for (int i = 0; i < NumberOfPDFFiles; i++)
            //    {
            //        totalCost = totalCost + 5.0;
            //        if (listofOwners[i] > 25)
            //        {
            //            totalCost = totalCost + (listofOwners[i] - 25) * 0.2;
            //        }
            //    }
            //    return totalCost;

            //}
        }
    }
}
