using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDF2excelConsole.ClassClosedXML;

namespace PDF2excelConsole
{
    internal class ClassJoinSplitManager1
    {
        public ClassFilesHandleClosedXML filesHandle;
        public ClassClosedXML excelOperations;
        public List<Classbatim> allBatim;
        public List<ClassTaboo> allTaboo;
        public ClassbatimManager1 BatimMgr;
        public ClasszhuiotManager1 TabuMgr;
        public ClassJoinSplitManager1(ClassFilesHandleClosedXML fhd, ClassClosedXML excel, ClassbatimManager1 batim, ClasszhuiotManager1 Tabu)
        {
            filesHandle = fhd;
            excelOperations = excel;
            BatimMgr = batim;
            TabuMgr = Tabu;
        }

        public void CreateJoinSplitTable()
        {
            allBatim = BatimMgr.allBatim;
            allTaboo = TabuMgr.allTaboo;

            if ( allBatim.Count == 0  && allTaboo.Count == 0)
            {
                return;
            }
            List<ClassBase> allNesachim = new List<ClassBase>();

            allNesachim = SolrtAllByGushHelkot();


            if (allNesachim.Count == 0) return;

            //           if ((allBatim is null)) return;
            excelOperations.createSheet(ClassClosedXML.Sheets.JoinSplit, "איחוד וחלוקה", Color.Yellow);
            excelOperations.setActiveSheet(ClassClosedXML.Sheets.JoinSplit);
            excelOperations.refreshAll();
            int currentrow;
            int globalCount = 0;
            int startrow = 0;
            int sectionStart = 0;
            currentrow = excelOperations.BuildJoinSplitHeader();

            sectionStart = currentrow;
            excelOperations.refreshAll();
            ClassClosedXML.Sheets splitpage;
            splitpage = ClassClosedXML.Sheets.JoinSplit;
            int presentTatHelka = 0;
            string name = "";
            string ID = "";
            string part = "";


            PutCellParameters celparam = new PutCellParameters();
            celparam.ifFrame = true;
            celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
            celparam.Weight = XLBorderStyleValues.Thin;
            celparam.xlVAlign = XLAlignmentVerticalValues.Center;
            celparam.ifmerge = false;
            XLColor oldColor;


            celparam.Columnextension = 1;
            celparam.fontSize = 10;

//            excelOperations.exitForDebug();

            foreach (ClassBase nesach in allNesachim)
            {
                celparam.colorbackground = XLColor.FromColor(  ClassUtils.GetRandomColour() );
                if (nesach.Name == "batim")
                {
                    Classbatim batim = (Classbatim)nesach;
                    try
                    {
                        if (batim.tatHelkot.Count == 0) continue;
                        globalCount++;
                        oldColor =celparam.colorbackground;
                        celparam.colorbackground = XLColor.FromColor(excelOperations.GetFromRGB(0xFF, 0xD9, 0x61));
                        celparam.ifmerge = false;
                        celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                        excelOperations.putValueWithParameter(splitpage, globalCount.ToString(), currentrow, 1, celparam);
                        celparam.colorbackground = oldColor;
                        celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                        int topHelahRow = currentrow;

                        for (int i = 0; i < batim.tatHelkot.Count; i++)
                        {
                            int partners = 0;
                            presentTatHelka = batim.tatHelkot[i].number;
                            partners = getNumberofPartnersPerTatHelka(batim, i);

                            celparam.ifmerge = true;
                            celparam.Rowextension = partners;
                            celparam.xlVAlign = XLAlignmentVerticalValues.Center;
                            celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                            
                            celparam.Weight = XLBorderStyleValues.Thin;

                            int topTatHelkaRow = currentrow;
                            excelOperations.putValueWithParameter(splitpage, presentTatHelka.ToString(), currentrow, 7, celparam);
//                            excelOperations.exitForDebug();

                            for (int own = 0; own < partners; own++)
                            {
                                if (own > 0)
                                {
                                    currentrow++;
                                    globalCount++;
                                    celparam.ifmerge = false;
                                    oldColor = celparam.colorbackground;
                                    celparam.colorbackground = XLColor.FromColor(excelOperations.GetFromRGB(0xFF, 0xD9, 0x61));
                                    excelOperations.putValueWithParameter(splitpage, globalCount.ToString(), currentrow, 1, celparam);
                                    celparam.colorbackground = oldColor;

                                }
                                if (batim.tatHelkot[i].leasings.Count > 0)
                                {
                                    name = batim.tatHelkot[i].leasings[own].Name[0];
                                    ID = batim.tatHelkot[i].leasings[own].id[0];
                                    part = batim.tatHelkot[i].leasings[own].part[0];
                                }
                                else
                                {
                                    name = batim.tatHelkot[i].owners[own].name;
                                    ID = batim.tatHelkot[i].owners[own].idNumber;
                                    part = batim.tatHelkot[i].owners[own].part;
                                }
                                celparam.ifmerge = false;
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                                excelOperations.putValueWithParameter(splitpage, name, currentrow, 8, celparam);
                                excelOperations.putValueWithParameter(splitpage, ID, currentrow, 9, celparam);
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                                excelOperations.putValueWithParameter(splitpage, part, currentrow, 10, celparam);
                            }
//                            excelOperations.exitForDebug();
                            celparam.ifmerge = true;
                            celparam.Rowextension = partners;
                            excelOperations.putValueWithParameter(splitpage, batim.tatHelkot[i].shetah, topTatHelkaRow, 11, celparam);
                            excelOperations.putValueWithParameter(splitpage, batim.tatHelkot[i].partincommon, topTatHelkaRow, 12, celparam);
                            double  percent = ClassUtils.convertPartToFraction1(batim.tatHelkot[i].partincommon);
                            excelOperations.putDoubleValueWithParameter(splitpage, percent, topTatHelkaRow, 13, celparam);
                            excelOperations.putValueWithParameter(splitpage, "", topTatHelkaRow, 14, celparam);
                            excelOperations.putValueWithParameter(splitpage, "", topTatHelkaRow, 15, celparam);
                            excelOperations.putValueWithParameter(splitpage, "", topTatHelkaRow, 16, celparam);
                            excelOperations.putValueWithParameter(splitpage, "", topTatHelkaRow, 17, celparam);
                            currentrow++;
                            celparam.ifmerge = false;

                            if (i < batim.tatHelkot.Count - 1)
                            {
                                globalCount++;
                                oldColor = celparam.colorbackground;
                                celparam.colorbackground = XLColor.FromColor(excelOperations.GetFromRGB(0xFF, 0xD9, 0x61));
                                excelOperations.putValueWithParameter(splitpage, globalCount.ToString(), currentrow, 1, celparam);
                                celparam.colorbackground = oldColor;
                            }
                        }
//                        excelOperations.exitForDebug();
                        celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                        celparam.ifmerge = true;
                        celparam.Rowextension = currentrow - topHelahRow;
                        celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                        celparam.xlVAlign = XLAlignmentVerticalValues.Top;
                        excelOperations.putValueWithParameter(splitpage, batim.header.gush, topHelahRow, 2, celparam);               // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, batim.header.helka, topHelahRow, 3, celparam);
                        excelOperations.putValueWithParameter(splitpage, batim.batimproperty.areasqmr, topHelahRow, 4, celparam);    // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, "", topHelahRow, 5, celparam);                              // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, batim.batimproperty.areasqmr, topHelahRow, 6, celparam);    // to be done at the end
                        excelOperations.refreshAll();
                    }
                    catch (Exception ex)
                    {

                    }
                }
                else
                {
                    ClassTaboo taboo = (ClassTaboo)nesach;
                    try
                    {
                        int lastone = 0;
                        int loc = 0;
                        int topcurrentRow = 0;
                        if (taboo.leasings != null)
                        {
                            lastone = taboo.leasings.Count - 1;
                            loc = taboo.leasings[taboo.leasings.Count - 1].leasingOwners.Count;
                            topcurrentRow = currentrow;
                            for (int j = 0; j < loc; j++)
                            {
                                globalCount++;
                                oldColor = celparam.colorbackground;
                                celparam.colorbackground = XLColor.FromColor(excelOperations.GetFromRGB(0xFF, 0xD9, 0x61));
                                celparam.ifmerge = false;
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                                excelOperations.putValueWithParameter(splitpage, globalCount.ToString(), topcurrentRow, 1, celparam);
                                celparam.colorbackground = oldColor;

                                name = taboo.leasings[lastone].leasingOwners[j].LeaserName;
                                ID = taboo.leasings[lastone].leasingOwners[j].idNumber;
                                part = taboo.leasings[lastone].leasingOwners[j].LeaserPart;
                                celparam.ifmerge = false;
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                                excelOperations.putValueWithParameter(splitpage, name, topcurrentRow, 8, celparam);
                                excelOperations.putValueWithParameter(splitpage, ID, topcurrentRow, 9, celparam);
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                                excelOperations.putValueWithParameter(splitpage, part, topcurrentRow, 10, celparam);
                                excelOperations.putValueWithParameter(splitpage, part, topcurrentRow, 12, celparam);
                                double percent = ClassUtils.convertPartToFraction1(part);
                                
                                excelOperations.putDoubleValueWithParameter(splitpage, percent, topcurrentRow, 13, celparam);
                                
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 14, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 15, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 16, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 17, celparam);

                                topcurrentRow++;
                            }
                        }
                        else
                        {
                            loc = taboo.zhuiotOwners.Count;
                            topcurrentRow = currentrow;
                            for (int j = 0; j < loc; j++)
                            {
                                globalCount++;
                                oldColor = celparam.colorbackground;
                                celparam.colorbackground = XLColor.FromColor(excelOperations.GetFromRGB(0xFF, 0xD9, 0x61));
                                celparam.ifmerge = false;
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                                excelOperations.putValueWithParameter(splitpage, globalCount.ToString(), topcurrentRow, 1, celparam);
                                celparam.colorbackground = oldColor;

                                name = taboo.zhuiotOwners[j].ownerName;
                                ID = taboo.zhuiotOwners[j].idNumber;
                                part = taboo.zhuiotOwners[j].ownerPart;
                                celparam.ifmerge = false;
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                                excelOperations.putValueWithParameter(splitpage, name, topcurrentRow, 8, celparam);
                                excelOperations.putValueWithParameter(splitpage, ID, topcurrentRow, 9, celparam);
                                celparam.xlHAlign = XLAlignmentHorizontalValues.Center;
                                excelOperations.putValueWithParameter(splitpage, part, topcurrentRow, 10, celparam);
                                excelOperations.putValueWithParameter(splitpage, part, topcurrentRow, 12, celparam);
                                double percent = ClassUtils.convertPartToFraction1(part);
                                excelOperations.putDoubleValueWithParameter(splitpage, percent, topcurrentRow, 13, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 14, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 15, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 16, celparam);
                                excelOperations.putValueWithParameter(splitpage, "", topcurrentRow, 17, celparam);

                                topcurrentRow++;
                            }

                        }

                        celparam.ifmerge = true;
                        celparam.Rowextension = loc;
                        celparam.xlHAlign = XLAlignmentHorizontalValues.Right;
                        celparam.xlVAlign = XLAlignmentVerticalValues.Top;

                        excelOperations.putValueWithParameter(splitpage, taboo.gush, currentrow, 2, celparam);               // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, taboo.helka, currentrow, 3, celparam);
                        excelOperations.putValueWithParameter(splitpage, taboo.description.area, currentrow, 4, celparam);    // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, "", currentrow, 5, celparam);                              // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, taboo.description.area, currentrow, 6, celparam);   // to be done at the end
                        excelOperations.putValueWithParameter(splitpage, "", currentrow, 7, celparam);
                        excelOperations.putValueWithParameter(splitpage, taboo.description.area, currentrow, 11, celparam);
                        excelOperations.refreshAll();
                        currentrow = currentrow + loc;

                    }
                    catch (Exception ex)
                    {

                    }

                }
                excelOperations.refreshAll();
            }
//            excelOperations.exitForDebug();
            excelOperations.getSheet(Sheets.JoinSplit).Columns(8, 8).AdjustToContents();
            
            excelOperations.CorrectFormatForSum(splitpage, 1, sectionStart, currentrow - 1, "@");
            excelOperations.CorrectFormatForSum(splitpage, 11, sectionStart, currentrow - 1, "#,##0.00");
            excelOperations.CorrectFormatForSum(splitpage, 13, sectionStart, currentrow - 1, "0.000%");
            excelOperations.CorrectFormatForSum(splitpage, 14, sectionStart, currentrow - 1, "#,##0.00");
            excelOperations.CorrectFormatForSum(splitpage, 15, sectionStart, currentrow - 1, "#,##0.00");
            excelOperations.CorrectFormatForSum(splitpage, 16, sectionStart, currentrow - 1, "#,##0.00");
            excelOperations.CorrectFormatForSum(splitpage, 17, sectionStart, currentrow - 1, "0.000%");
//            excelOperations.exitForDebug();
        }

        public int getNumberofPartnersOfHelkaZhuiot(ClassTaboo taboo)
        {
            int ret = 0;
            ret = taboo.zhuiotOwners.Count;
            if (taboo.leasings != null)
            {
                ret = taboo.leasings[taboo.leasings.Count - 1].leasingOwners.Count;
            }
            return ret;
        }
        public int getNumberofPartnersPerTatHelka(Classbatim bait, int tatHelka)
        {
            int ret = 0;
            ret = bait.tatHelkot[tatHelka].owners.Count;
            if (bait.tatHelkot[tatHelka].leasings.Count > 0)
            {
                ret = bait.tatHelkot[tatHelka].leasings.Count;
            }
            return ret;
        }
        public Dictionary<ClassBase, int> sortbyGush()
        {
            Dictionary<ClassBase, int> dictionary = new Dictionary<ClassBase, int>();
            Dictionary<ClassBase, int> dictionary0 = new Dictionary<ClassBase, int>();
            if (allTaboo != null)
            {
                foreach (ClassTaboo tabu in allTaboo)
                {
                    dictionary.Add(tabu, Convert.ToInt32(tabu.gush));
                }
            }
            if (allBatim != null)
            {
                foreach (Classbatim bait in allBatim)
                {
                    dictionary.Add(bait, Convert.ToInt32(bait.gush));
                }
            }
            IEnumerable<KeyValuePair<ClassBase, int>> sortedDict = from entry in dictionary orderby entry.Value ascending select entry;
            dictionary0 = sortedDict.ToDictionary(pair => pair.Key, pair => pair.Value);

            return dictionary0;
        }

        public List<ClassBase> sortDictionaryByHelka(Dictionary<ClassBase, int> sourceDic)
        {
            List<ClassBase> ret = new List<ClassBase>();
            Dictionary<ClassBase, int> helkasection = new Dictionary<ClassBase, int>();

            var first = sourceDic.First();
            int val = first.Value;

            foreach (var item in sourceDic)
            {
                if (val == item.Value)
                {
                    helkasection.Add(item.Key, Convert.ToInt32(item.Key.helka));
                    continue;
                }
                else
                {
                    val = item.Value;
                    Dictionary<ClassBase, int> dictionary0 = new Dictionary<ClassBase, int>();
                    IEnumerable<KeyValuePair<ClassBase, int>> sortedDict = from entry in helkasection orderby entry.Value ascending select entry;
                    dictionary0 = sortedDict.ToDictionary(pair => pair.Key, pair => pair.Value);
                    foreach (var item1 in dictionary0)
                    {
                        ret.Add(item1.Key);
                    }
                    helkasection.Clear();
                    helkasection.Add(item.Key, Convert.ToInt32(item.Key.helka));

                }
            }
            Dictionary<ClassBase, int> dictionary1 = new Dictionary<ClassBase, int>();
            IEnumerable<KeyValuePair<ClassBase, int>> sortedDict0 = from entry in helkasection orderby entry.Value ascending select entry;
            dictionary1 = sortedDict0.ToDictionary(pair => pair.Key, pair => pair.Value);
            foreach (var item1 in dictionary1)
            {
                ret.Add(item1.Key);
            }

            return ret;
        }
        public List<ClassBase> SortTabuFiles()
        {
            List<ClassBase> allTabooANDBatim = new List<ClassBase>();
            Dictionary<ClassBase, int> dictionary = new Dictionary<ClassBase, int>();
            Dictionary<ClassBase, int> dictionary0 = new Dictionary<ClassBase, int>();
            Dictionary<ClassBase, int> finalDictionary = new Dictionary<ClassBase, int>();

            if (allTaboo != null)
            {
                foreach (ClassTaboo tabu in allTaboo)
                {
                    dictionary.Add(tabu, Convert.ToInt32(tabu.gush));
                }
            }
            if (allBatim != null)
            {
                foreach (Classbatim bait in allBatim)
                {
                    dictionary.Add(bait, Convert.ToInt32(bait.gush));
                }
            }
            IEnumerable<KeyValuePair<ClassBase, int>> sortedDict = from entry in dictionary orderby entry.Value ascending select entry;
            dictionary0 = sortedDict.ToDictionary(pair => pair.Key, pair => pair.Value);

            var first = dictionary0.First();
            int val = first.Value;

            dictionary.Clear();

            ClassBase classBase = null;

            foreach (var item in dictionary0)
            {
                var nextone = item.Value;
                if (nextone == val)
                {
                    classBase = item.Key;

                    dictionary.Add(item.Key, Convert.ToInt32(classBase.helka));
                }
                else
                {
                    sortedDict = from entry in dictionary orderby entry.Value ascending select entry;
                    finalDictionary = sortedDict.ToDictionary(pair => pair.Key, pair => pair.Value);

                    val = nextone;
                }
            }
            return allTabooANDBatim;
        }

        public List<ClassBase> SolrtAllByGushHelkot()
        {
            List<ClassBase> all = new List<ClassBase>();

            Dictionary<ClassBase, int> dictionaryofhelkot = new Dictionary<ClassBase, int>();
            Dictionary<ClassBase, int> allnesahByGush = sortbyGush();
            foreach (var item in allnesahByGush)
            {
                dictionaryofhelkot.Add(item.Key, Convert.ToInt32(item.Value));
            }
            all = sortDictionaryByHelka(dictionaryofhelkot);

            return all;
        }

    }
}
