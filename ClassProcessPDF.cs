using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    internal class ClassProcessPDF
    {
        bool DebugMode;
        int customerType;
        public ClassClosedXML excelOperations;
        public ClassFilesHandleClosedXML fileHandler;
        ClassbatimManager1 batimManager;
        ClasszhuiotManager1 zuiotManager;

        public ClassProcessPDF(string[] sarray, bool debugMode, string tempFolder,  int custType)
        {
            DebugMode = debugMode;
            customerType = custType;
            excelOperations = new ClassClosedXML();

            fileHandler = new ClassFilesHandleClosedXML(excelOperations, DebugMode, tempFolder , sarray);
//            excelOperations.exitForDebug();
            batimManager = new ClassbatimManager1(fileHandler, excelOperations);
            zuiotManager = new ClasszhuiotManager1(fileHandler, excelOperations);
            excelOperations.createSheet(ClassClosedXML.Sheets.PDFfiles, "נסחים", Color.Black);
 //           excelOperations.exitForDebug();



            fileHandler.clearCSVFiles("batim");
            fileHandler.clearCSVFiles("zhuiot");
            excelOperations.ListPdfFiles(sarray);
//            excelOperations.exitForDebug();
        }

        public string convert()
        {
            string resultfile; 
            string sret = "";
            fileHandler.convertPDF2CSV();
 //           excelOperations.exitForDebug();
            batimManager.convertBatimtoExcel();            
            zuiotManager.convertZhuiottoExcel();
//            excelOperations.exitForDebug();


            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimProperty);
            batimManager.CreatePropertyTable();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimLeasing);
            batimManager.CreateBatimLeasing();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimMortgage);
            batimManager.CreateBatimMortgage();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimRemarks);
            batimManager.CreateBatimRemarksTables();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimAttachments);
            batimManager.createBatimAttachments();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimOwners);
            batimManager.CreateBatimOwnTable();
//            excelOperations.exitForDebug();
            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Property);
            zuiotManager.CreatePropertyTables();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Leasing);
            zuiotManager.CreateLeasingTables();
//            excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Mortgage);
            zuiotManager.CreateMortGageTables();
//            excelOperations.exitForDebug();

            //           excelOperations.deleteSheet(ClassExcelOperations.Sheets.Remark);
            zuiotManager.CreateRemarksTables();
 //           excelOperations.exitForDebug();

            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Owner);
            zuiotManager.CreateOwnersTable();
//            excelOperations.exitForDebug();
            //
            //  
            //
            ClassJoinSplitManager1 joinSplitManager;
            joinSplitManager = new ClassJoinSplitManager1(fileHandler, excelOperations, batimManager, zuiotManager);
            joinSplitManager.CreateJoinSplitTable();
            excelOperations.exitForDebug();
            //if (customerType > 80)
            //{
            //    ClassJoinSplitManager joinSplitManager;
            //    joinSplitManager = new ClassJoinSplitManager(fileHandler, excelOperations, batimManager, zuiotManager);
            //    excelOperations.deleteSheet(ClassExcelOperations.Sheets.JoinSplit);
            //    joinSplitManager.CreateJoinSplitTable();
            //}

            resultfile = fileHandler.PDFfolder + "\\Tabu_results_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            resultfile = resultfile + ".xlsx";
            excelOperations.saveResults(resultfile);

            return resultfile;

            return sret;
        }
        public List<int> getTotalNumberOfOwners()
        {
            List<int> numberOfOwners = new List<int>();


            if (batimManager.allBatim.Count > 0)
            {
                foreach (Classbatim batim in batimManager.allBatim)
                {
                    int ret = 0;
                    for (int i = 0; i < batim.tatHelkot.Count; i++)
                    {
                        ret = ret + batim.tatHelkot[i].owners.Count;
                    }
                    numberOfOwners.Add(ret);
                }
            }
            if (zuiotManager.allTaboo.Count > 0)
            {
                if (zuiotManager.allTaboo.Count > 0)
                {
                    foreach (ClassTaboo taboo in zuiotManager.allTaboo)
                    {
                        numberOfOwners.Add(taboo.zhuiotOwners.Count);
                    }
                }
            }
            return numberOfOwners;
        }

    }
}
