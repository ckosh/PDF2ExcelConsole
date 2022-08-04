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
        static public int delaySeconds = 30;
        static public string Tempfolder = "d:\\temp";
        static public int opsticalMinutes = 30;
        static public bool debugMode = true;
        static public int Pop3Port =  110;
        static public string resultExcelFile;
        static public bool skipMail = false;
        static public bool ssl = false;

        static void Main(string[] args)
        {
            string userMail;
            List<int> listofOwners;
            int numberOfOwners;
            Console.WriteLine("Start");
            if (args.Length == 0)
            {
                Console.WriteLine("Invalid args");
                return;
            }
            else
            {
                string ParamFile = args[0];
                ClassArgsParser argParser = new ClassArgsParser();
                argParser.parseArgs(ParamFile);

                delaySeconds = argParser.getDelaySeconds();
                Tempfolder = argParser.getTempFolder();
                debugMode = argParser.getDebugMode();
                Pop3Port = argParser.getPort();
                skipMail = argParser.getSkipMail();
                ssl = argParser.getussl();


                if ( skipMail)
                {
                    string[] temp = Directory.GetFiles(Tempfolder, "*.pdf");
                    string[] PdfFileNames = new string[temp.Length];
                    for (int k = 0; k < temp.Length; k++)
                    {
                        PdfFileNames[k] = Path.GetFileName(temp[k]);
                    }
                    ClassProcessPDF processPDF = new ClassProcessPDF(PdfFileNames, debugMode, Tempfolder);
                    resultExcelFile = processPDF.convert();

                }
                else
                {
                    ClassMailManager1 mailManager = new ClassMailManager1(debugMode, Pop3Port, ssl , Tempfolder);
                    while (true)
                    {
                        userMail = mailManager.ConnectToPop3();
                        if (userMail != null)
                        {
                            if (userMail.Length > 0)
                            {
                                Console.WriteLine(userMail);
                                //                    Log.Info("before process files  ");
                                string[] PdfFileNames = mailManager.GetPdfFiles();
                                ClassProcessPDF processPDF = new ClassProcessPDF(PdfFileNames, debugMode, Tempfolder);
                                //            Log.Info("after process files  ");
                                resultExcelFile = processPDF.convert();
                                Console.WriteLine(resultExcelFile);
                                //            Log.Info("after convert files  ");

                                listofOwners = processPDF.getTotalNumberOfOwners();
                                numberOfOwners = listofOwners.Sum();
                                Console.WriteLine("owners " + numberOfOwners.ToString());
                                Console.WriteLine("files " + PdfFileNames.Length.ToString());

                                double totalcost = getTotalCost(PdfFileNames.Length);
                                Console.WriteLine(numberOfOwners.ToString());

                                string sstype = "";
                                sstype = "המרת נסחים - בניית טבלת מצב נכנס";
                                string body = totalcost.ToString() + " עלות הסבה" + '\n' + PdfFileNames.Length.ToString() + " מספר נסחים " + '\n' + numberOfOwners.ToString() + " מספר בעלים " + '\n' + sstype;
                                mailManager.sendMail(userMail, "תוצאות הסבת נסחי טאבו", resultExcelFile, body);
                                ClassUtils.deleteAllFilesFromDirectory(Tempfolder);
                                ClassUtils.deleteAllFilesFromDirectory(Tempfolder + "\\CSV");
                                mailManager.savecBillingData(userMail, PdfFileNames.Length, resultExcelFile, numberOfOwners, totalcost);
                            }
                        }
                        Thread.Sleep(delaySeconds * 1000);
                    }
                }
            }


            


            double getTotalCost(int NumberOfPDFFiles)
            {
                double totalCost = 0;
                for (int i = 0; i < NumberOfPDFFiles; i++)
                {
                    totalCost = totalCost + 5.0;
                    if (listofOwners[i] > 25)
                    {
                        totalCost = totalCost + (listofOwners[i] - 25) * 0.2;
                    }
                }
                return totalCost;

            }
        }
    }
}
