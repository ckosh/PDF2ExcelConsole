using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2excelConsole
{
    internal class ClassClosedXML
    {
        public IXLWorkbook xlWorkBook;

        public IXLWorksheet xlFilesWorkSheet;
        public IXLWorksheet xlOwnersSheet;
        public IXLWorksheet xlRemarksSheet;
        public IXLWorksheet xlLeasingSheet;
        public IXLWorksheet xlZikotSheet;
        public IXLWorksheet xlMortgageSheet;
        public IXLWorksheet xlPropertySheet;
        public IXLWorksheet xlBatimPropertySheet;
        public IXLWorksheet xlBatimOwnersSheet;
        public IXLWorksheet xlBatimErrorSheet;
        public IXLWorksheet xlBatimLeasingSheet;
        public IXLWorksheet xlBatimMortgageSheet;
        public IXLWorksheet xlBatimRemarksSheet;
        public IXLWorksheet xlBatimAttachmentsSheet;
        public IXLWorksheet xlJoinSplitSheet;

        public string PDFFolder;
        public enum Sheets
        {
            BatimError,
            Owner,
            Remark,
            Mortgage,
            Leasing,
            Zikot,
            PDFfiles,
            Property,
            BatimProperty,
            BatimOwners,
            BatimLeasing,
            BatimMortgage,
            BatimRemarks,
            BatimAttachments,
            JoinSplit
        }
        public ClassClosedXML()
        {
//            testHyperlink();

            xlWorkBook = new XLWorkbook();
            createSheet(Sheets.BatimError, "שגיאות", Color.Red);
            BuildBatimErrorHeader();
            // public void setBoarder(Sheets sn, int irow1, int irow2, int icol1, int icol2, int thickness)

            setBoarder(Sheets.BatimError, 9, 12, 8, 9, 0);
            setBoarder(Sheets.BatimError, 15, 17, 8, 9, 1);
            setBoarder(Sheets.BatimError, 19, 21, 8, 9, 2);
            setBoarder(Sheets.BatimError, 23, 25, 8, 9, 3);

            
            IXLWorksheet asheet = getSheet(Sheets.BatimError);
            //asheet.Column(12).Style.NumberFormat.Format = "0.00000%";

            setColumnPercentFormat(Sheets.BatimError, 12, 3);

            asheet.Cell(1, 12).Value = ClassUtils.convertPartToFraction1("1/3"); 
            asheet.Cell(2, 12).Value = ClassUtils.convertPartToFraction1("בשלמות");
            asheet.Cell(3, 12).Value = ClassUtils.convertPartToFraction1("1/7");

//            exitForDebug();
        }

        public void setColumnPercentFormat(Sheets sn, int columnNumber, int precision)
        {
            IXLWorksheet ws = getSheet(sn);
            string fff = "0.";
            for (int i = 0; i < precision; i++) fff = fff + "0";
            fff = fff + "%";
            ws.Column(columnNumber).Style.NumberFormat.Format = fff ;
        }

        public void saveResults(string resultName)
        {
            xlWorkBook.SaveAs(resultName);
        }

        public void disposeAllSheets()
        {
            if (xlOwnersSheet != null) xlOwnersSheet.Delete();
            if (xlLeasingSheet != null) xlLeasingSheet.Delete();
            if (xlMortgageSheet != null) xlMortgageSheet.Delete();
            if (xlFilesWorkSheet != null) xlFilesWorkSheet.Delete();
            if (xlZikotSheet != null) xlZikotSheet.Delete();
            if (xlPropertySheet != null) xlPropertySheet.Delete();
            if (xlBatimPropertySheet != null) xlBatimPropertySheet.Delete();
            if (xlBatimOwnersSheet != null) xlBatimOwnersSheet.Delete();
            if (xlBatimLeasingSheet != null) xlBatimLeasingSheet.Delete();
            if (xlBatimErrorSheet != null) xlBatimErrorSheet.Delete();
            if (xlBatimMortgageSheet != null) xlBatimMortgageSheet.Delete();
            if (xlBatimRemarksSheet != null) xlBatimRemarksSheet.Delete();
            if (xlBatimAttachmentsSheet != null) xlBatimAttachmentsSheet.Delete();
            if (xlJoinSplitSheet != null) xlJoinSplitSheet.Delete();
            xlWorkBook.Dispose();
        }
        public void createSheet(Sheets sn, string name, Color col)
        {
            switch (sn)
            {
                case Sheets.Owner:
                    xlOwnersSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlOwnersSheet.RightToLeft = true;
                    break;
                case Sheets.Leasing:
                    xlLeasingSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlLeasingSheet.RightToLeft = true;
                    break;
                case Sheets.Mortgage:
                    xlMortgageSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlMortgageSheet.RightToLeft = true;
                    break;
                case Sheets.PDFfiles:
                    xlFilesWorkSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlFilesWorkSheet.RightToLeft = true;
                    break;
                case Sheets.Remark:
                    xlRemarksSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlRemarksSheet.RightToLeft = true;
                    break;
                case Sheets.Zikot:
                    xlZikotSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlZikotSheet.RightToLeft = true;
                    break;
                case Sheets.Property:
                    xlPropertySheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlPropertySheet.Name = name;
                    xlPropertySheet.RightToLeft = true;
                    break;
                case Sheets.BatimProperty:
                    xlBatimPropertySheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimPropertySheet.RightToLeft = true;
                    break;
                case Sheets.BatimOwners:
                    xlBatimOwnersSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimOwnersSheet.RightToLeft = true;
                    break;
                case Sheets.BatimLeasing:
                    xlBatimLeasingSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimLeasingSheet.RightToLeft = true;
                    break;
                case Sheets.BatimError:
                    xlBatimErrorSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimErrorSheet.RightToLeft = true;
                    break;
                case Sheets.BatimMortgage:
                    xlBatimMortgageSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimMortgageSheet.RightToLeft = true;
                    break;
                case Sheets.BatimRemarks:
                    xlBatimRemarksSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimRemarksSheet.RightToLeft = true;
                    break;
                case Sheets.BatimAttachments:
                    xlBatimAttachmentsSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlBatimAttachmentsSheet.RightToLeft = true;
                    break;
                case Sheets.JoinSplit:
                    xlJoinSplitSheet = xlWorkBook.Worksheets.Add(name).SetTabColor(XLColor.FromColor(col));
                    xlJoinSplitSheet.RightToLeft = true;
                    break;
            }
        }

        public IXLWorksheet getSheet(Sheets sn)
        {
            IXLWorksheet retSheet = null;
            switch (sn)
            {
                case Sheets.Owner:
                    retSheet = xlOwnersSheet;
                    break;
                case Sheets.Leasing:
                    retSheet = xlLeasingSheet;
                    break;
                case Sheets.Mortgage:
                    retSheet = xlMortgageSheet;
                    break;
                case Sheets.PDFfiles:
                    retSheet = xlFilesWorkSheet;
                    break;
                case Sheets.Remark:
                    retSheet = xlRemarksSheet;
                    break;
                case Sheets.Zikot:
                    retSheet = xlZikotSheet;
                    break;
                case Sheets.Property:
                    retSheet = xlPropertySheet;
                    break;
                case Sheets.BatimProperty:
                    retSheet = xlBatimPropertySheet;
                    break;
                case Sheets.BatimOwners:
                    retSheet = xlBatimOwnersSheet;
                    break;
                case Sheets.BatimLeasing:
                    retSheet = xlBatimLeasingSheet;
                    break;
                case Sheets.BatimError:
                    retSheet = xlBatimErrorSheet;
                    break;
                case Sheets.BatimMortgage:
                    retSheet = xlBatimMortgageSheet;
                    break;
                case Sheets.BatimRemarks:
                    retSheet = xlBatimRemarksSheet;
                    break;
                case Sheets.BatimAttachments:
                    retSheet = xlBatimAttachmentsSheet;
                    break;
                case Sheets.JoinSplit:
                    retSheet = xlJoinSplitSheet;
                    break;
            }
            return retSheet;
        }

        public void HeadTitle(IXLWorksheet theSheet, string title, int irow, int icol, int endcol, XLAlignmentHorizontalValues hAlign, XLAlignmentVerticalValues vAlign, int fonSize, bool bBolt, System.Drawing.Color titleColor, int rowHeight, bool boarder, XLBorderStyleValues weight)
        {
            theSheet.Cell(irow,icol).Value = title;
            theSheet.Cell(irow, icol).Style.Alignment.WrapText = true;
            IXLRange titlerange = theSheet.Range(theSheet.Cell(irow, icol).Address, theSheet.Cell(irow, endcol).Address);
            titlerange.Merge();
            titlerange.Style.Alignment.Horizontal = hAlign;
            titlerange.Style.Alignment.Vertical = vAlign;
            titlerange.Style.Font.FontSize = fonSize;
            titlerange.Style.Font.Bold = bBolt;
            titlerange.Style.Font.FontName = "Tahoma";
            titlerange.Style.Fill.BackgroundColor = XLColor.FromColor(titleColor);
            theSheet.Row(irow).Height = rowHeight;
            titlerange.Style.NumberFormat.Format = "@";

 //           titleRang.RowHeight = rowHeight;
 
            if (boarder)
            {
                titlerange.Style.Border.TopBorder = weight;
                titlerange.Style.Border.LeftBorder = weight;
                titlerange.Style.Border.BottomBorder = weight;
                titlerange.Style.Border.RightBorder = weight;
 //               titleRang.BorderAround(XlLineStyle.xlContinuous, weight, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
            }
        }

        public int BuildBatimErrorHeader()
        {
            int rowNumber = 1;
            HeadTitle(xlBatimErrorSheet, "קובץ", 1, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium );
            HeadTitle(xlBatimErrorSheet, "פסקה", 1, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimErrorSheet, "תת חלקה", 1, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimErrorSheet, "פסקה", 1, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimErrorSheet, "הערות", 1, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimErrorSheet, "2", 1, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);


            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(1)).Width = 8.0;
            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(2)).Width = 8.0;
            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(3)).Width = 11.0;
            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(4)).Width = 8.0;
            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(5)).Width = 8.0;
            xlBatimErrorSheet.Column(ClassUtils.ColumnLabel(6)).Width = 5.0;
            return rowNumber;
        }

        public int BuildBatimRemarksHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimRemarksSheet, "הערות - בתים משותפים", 1, 1, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);

            HeadTitle(xlBatimRemarksSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "תת חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "סוג הערה", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "שם", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "סוג זיהוי", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "מס. זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "חלק", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "שטר", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "הערה", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimRemarksSheet, "קובץ", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            return rowNumber;
        }
        public int BuildBatimMortgageHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimMortgageSheet, "משכנתאות - בתים משותפים", 1, 1, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);

            HeadTitle(xlBatimMortgageSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "תת חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "משכנתה", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "ממשכן", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "סוג זיהוי", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "מס. זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "חלק", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "שטר", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "דרגה", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "הערות", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimMortgageSheet, "קובץ", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            return rowNumber;
        }
        public int BuildBatimAttachmentsHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimAttachmentsSheet, "הצמדות - בתים משותפים", 1, 1, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "תת חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "סימון בתשריט", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "צבע בתשריט", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "תיאור הצמדה", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "משותפת ל", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "שטח במ\"ר", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimAttachmentsSheet, "קובץ", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            return rowNumber;
        }
        public int BuildBatimLeasingHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimLeasingSheet, "חכירות - בתים משותפים", 1, 1, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "תת חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "חכירה", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "שם", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "סוג זיהוי", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "מס. זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "חלק בנכס", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "שטר", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "רמה", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "תאריך סיום", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "הערות", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "חלק בנכס", 2, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimLeasingSheet, "קובץ", 2, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);



            return rowNumber;
        }
        public int BuildBatimOwnerHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimOwnersSheet, "בעלים - בתים משותפים", 1, 1, 18, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "לחץ על סימון להגיע לדף הנתונים", 1, 19, 23, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "", 1, 24, 24, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);

            HeadTitle(xlBatimOwnersSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "שטח חלקה במ\"ר", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "תת חלקה", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "שטח במ\"ר", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "תיאור קומה", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "כניסה", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "אגף", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "מבנה", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "החלק ברכוש המשותף", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "החלק באחוזים", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);


            HeadTitle(xlBatimOwnersSheet, "קניין", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "שם", 2, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "סוג זיהוי", 2, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "מס. זיהוי", 2, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "החלק בתת חלקה", 2, 16, 16, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "החלק באחוזים", 2, 17, 17, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);


            HeadTitle(xlBatimOwnersSheet, "שטר", 2, 18, 18, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "משכנתאות", 2, 19, 19, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "הערות", 2, 20, 20, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "חכירות", 2, 21, 21, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "הצמדות", 2, 22, 22, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "זיקות הנאה", 2, 23, 23, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimOwnersSheet, "קובץ", 2, 24, 24, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);

            setColumnPercentFormat(Sheets.BatimOwners,11,5);
            setColumnPercentFormat(Sheets.BatimOwners, 17, 5);

            return rowNumber;
        }
        public int BuildBatimPropertyHeader()
        {
            int rowNumber = 4;
            HeadTitle(xlBatimPropertySheet, "רכוש משותף - בתים משותפים", 1, 1, 19, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);

            HeadTitle(xlBatimPropertySheet, "גו\"ח", 2, 1, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "הנכס נוצר", 2, 3, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "הרכוש המשותף", 2, 6, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "הערות - זיקות הנאה", 2, 15, 16, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "גירסת נסח", 2, 17, 19, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "גוש", 3, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "חלקה", 3, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "שטר", 3, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "מיום", 3, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "סוג השטר", 3, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "רשויות", 3, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "שטח במ\"ר", 3, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "תת חלקות", 3, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "תקנון", 3, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "שטר יוצר", 3, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "תיק יוצר", 3, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "תיק בית משותף", 3, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "כתובת", 3, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "הערות", 3, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "זיקות הנאה", 3, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "הערות", 3, 16, 16, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "תאריך", 3, 17, 17, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "מס. נסח", 3, 18, 18, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);
            HeadTitle(xlBatimPropertySheet, "קובץ", 3, 19, 19, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 20, true, XLBorderStyleValues.Medium);

            xlBatimPropertySheet.Columns(1,19).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            xlBatimPropertySheet.Columns(1, 19).AdjustToContents();

            return rowNumber;
        }

        public int BuildMortgageHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlMortgageSheet, "משכנתאות - פנקס הזכויות", 1, 1, 18, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlMortgageSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "מס. שטר", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "תאריך", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "מהות פעולה", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "בעלי משכנתה", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "סוג זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "מס' זיהוי", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "שם הלווה", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "סוג זיהוי", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "מס' זיהוי", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "דרגה", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "סכום", 2, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "בתנאי שטר מקורי", 2, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "החלק בנכס", 2, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "החלק בשבר", 2, 16, 16, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "הערות", 2, 17, 17, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlMortgageSheet, "שם קובץ", 2, 18, 18, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);

            xlMortgageSheet.Column(ClassUtils.ColumnLabel(1)).Width = 5.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(2)).Width = 4.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(3)).Width = 10.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(4)).Width = 9.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(5)).Width = 6.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(6)).Width = 20.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(7)).Width = 5.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(8)).Width = 9.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(9)).Width = 10.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(10)).Width = 5.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(11)).Width = 6.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(12)).Width = 6.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(13)).Width = 13.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(14)).Width = 10.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(15)).Width = 10.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(16)).Width = 30.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(17)).Width = 15.0;
            xlMortgageSheet.Column(ClassUtils.ColumnLabel(18)).Width = 15.0;

            return rowNumber;
        }
        public int BuildRemarkHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlRemarksSheet, "הערות - פנקס הזכויות", 1, 1, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlRemarksSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "מס. שטר", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "תאריך", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "מהות פעולה", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "שם המוטב", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "סוג זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "מס. זיהוי", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "הערות", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlRemarksSheet, "שם קובץ", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);


            xlRemarksSheet.Column(ClassUtils.ColumnLabel(1)).Width = 6.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(2)).Width = 5.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(3)).Width = 15.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(4)).Width = 12.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(5)).Width = 30.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(6)).Width = 20.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(7)).Width = 6.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(8)).Width = 15.0;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(9)).Width = 40.00;
            xlRemarksSheet.Column(ClassUtils.ColumnLabel(10)).Width = 10.00;
            return rowNumber;
        }
        public int BuildLeasingHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlLeasingSheet, "חכירות - פנקס הזכויות", 1, 1, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlLeasingSheet, "גוש", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "חלקה", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "מס. שטר", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "תאריך", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "מהות הפעולה", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "שם החוכר", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "סוג זיהוי", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "מס. זיהוי", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "החלק בנכס", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "רמת חכירה", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "בתנאי שטר מקורי", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "תאריך סיום", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "החלק בנכס", 2, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "הערות", 2, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlLeasingSheet, "שם קובץ", 2, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);

            xlLeasingSheet.Column(ClassUtils.ColumnLabel(1)).Width = 6.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(2)).Width = 5.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(3)).Width = 12.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(4)).Width = 12.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(5)).Width = 25.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(6)).Width = 15.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(7)).Width = 4.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(8)).Width = 10.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(9)).Width = 15.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(10)).Width = 18.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(11)).Width = 12.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(12)).Width = 10.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(13)).Width = 8.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(14)).Width = 30.00;
            xlLeasingSheet.Column(ClassUtils.ColumnLabel(15)).Width = 10.00;

            return rowNumber;
        }
        public int BuildPropertyHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlPropertySheet, "תאור הנכס - פנקס הזכויות", 1, 1, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlPropertySheet, "מס\"ד", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "גוש", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "רשויות", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "שטח במ\"ר", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "סוג המקרקעין", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "הערות רשם המקרקעין", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "המספרים הישנים של החלקה", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlPropertySheet, "שם הקובץ", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);

            xlPropertySheet.Column(ClassUtils.ColumnLabel(1)).Width = 4.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(2)).Width = 10.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(3)).Width = 10.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(4)).Width = 30.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(5)).Width = 15.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(6)).Width = 15.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(7)).Width = 40.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(8)).Width = 20.00;
            xlPropertySheet.Column(ClassUtils.ColumnLabel(9)).Width = 20.00;

            return rowNumber;
        }
        public int buildOwnerHeadr0()
        {
            int rowNumber = 1;
            HeadTitle(xlOwnersSheet, "בעלויות - פנקס הזכויות", 1, 1, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlOwnersSheet, "לחץ על הסמן לצפיה", 1, 12, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);
            HeadTitle(xlOwnersSheet, "", 1, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 12, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Medium);

            HeadTitle(xlOwnersSheet, "מס\"ד", 2, 1, 1, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "גוש", 2, 2, 2, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "חלקה", 2, 3, 3, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "מס' שטר", 2, 4, 4, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "תאריך", 2, 5, 5, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "מהות פעולה", 2, 6, 6, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "הבעלים", 2, 7, 7, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "סוג זיהוי", 2, 8, 8, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "מס' זיהוי", 2, 9, 9, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "החלק בנכס", 2, 10, 10, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "אחוז בנכס", 2, 11, 11, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "משכנתאות", 2, 12, 12, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "חכירות", 2, 13, 13, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "הערות", 2, 14, 14, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);
            HeadTitle(xlOwnersSheet, "שם קובץ", 2, 15, 15, XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 11, true, System.Drawing.Color.Aqua, 40, true, XLBorderStyleValues.Thin);

            xlOwnersSheet.Column(ClassUtils.ColumnLabel(1)).Width = 4.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(2)).Width = 8.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(3)).Width = 4.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(4)).Width = 14.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(5)).Width = 10.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(6)).Width = 20.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(7)).Width = 15.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(8)).Width = 6.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(9)).Width = 12.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(10)).Width = 20.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(11)).Width = 10.00;
            xlOwnersSheet.Column(ClassUtils.ColumnLabel(12)).Width = 20.00;

            rowNumber = 3;
            return rowNumber;
        }

        public int BuildJoinSplitHeader()
        {
            int rowNumber = 3;
            System.Drawing.Color PattensBlue = GetFromRGB(0xDA, 0xEE, 0xF3);
            string NewShekel = "\u20AA";
            string test1 = "שווי הזכויות במצב הנכנס (" + NewShekel + ")";

            HeadTitle(xlJoinSplitSheet, "פינוי בינוי", 1, 1, 17,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 13, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "נתוני המקרקעין", 2, 1, 6,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 113, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "מצב נכנס", 2, 7, 17,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 13, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "ספירה", 3, 1, 1,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "גוש", 3, 2, 2,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "חלקה", 3, 3, 3,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שטח החלקה הרשום (במ\"ר)", 3, 4, 4,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "ייעוד החלקה", 3, 5, 5,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שטח החלקה הכלול באיחוד וחלוקה", 3, 6, 6,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "מס' תת חלקה", 3, 7, 7,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שם הבעלים / חוכר הרשום", 3, 8, 8,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "ת.ז / ח.פ", 3, 9, 9,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "חלק הבעלים בנכס", 3, 10, 10,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שטח תת חלקה רשום (במ\"ר)", 3, 11, 11,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "חלק ברכוש המשותף (בשבר)", 3, 12, 12,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "חלק ברכוש המשותף (ב%)", 3, 13, 13,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שווי הזכויות במצב הנכנס (" + NewShekel + ")", 3, 14, 14,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שווי תרומת המבנים (" + NewShekel + ")", 3, 15, 15,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שווי זכויות + מחוברים (" + NewShekel + ")", 3, 16, 16,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "שווי יחסי (באחוזים)", 3, 17, 17,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 61, true, XLBorderStyleValues.Thin);

            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(1)).Width = 6.82;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(1)).Width = 6.82;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(2)).Width = 4.55;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(3)).Width = 7.0;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(4)).Width = 9.55;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(5)).Width = 8.0;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(6)).Width = 10.0;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(7)).Width = 8.82;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(8)).Width = 18.0;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(9)).Width = 13.36;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(10)).Width = 8.73;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(11)).Width = 7.36;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(12)).Width = 7.27;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(13)).Width = 7.36;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(14)).Width = 12.73;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(15)).Width = 11.27;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(16)).Width = 12.55;
            xlJoinSplitSheet.Column(ClassUtils.ColumnLabel(17)).Width = 7.36;

            HeadTitle(xlJoinSplitSheet, "", 4, 1, 1,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "a", 4, 2, 2,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "b", 4, 3, 3,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "c", 4, 4, 4,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 5, 5,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 6, 6,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "d", 4, 7, 7,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "e", 4, 8, 8,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "f", 4, 9, 9,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "g", 4, 10, 10,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "h", 4, 11, 11,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "i", 4, 12, 12,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "j", 4, 13, 13,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 14, 14,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 15, 15,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 16, 16,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            HeadTitle(xlJoinSplitSheet, "", 4, 17, 17,  XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center, 10, true, PattensBlue, 15, true, XLBorderStyleValues.Thin);
            rowNumber = 5;
            return rowNumber;
        }

        public int getBatimErrorLine()
        {
            int ret = 0;
            try
            {
                var ws1 = xlWorkBook.Worksheet("שגיאות").Cell(1, 6).Value;
                ret = Convert.ToInt32(ws1);
                return ret;
            }
            catch (Exception e)
            {
               //
            }
            return ret;
        }

        public void PutLitteralStringValue(Sheets sn, int row, int column, string val)
        {
            IXLWorksheet asheet = getSheet(sn);
            if (!(val is null))
            {
                asheet.Cell(row, column).SetValue<string>(Convert.ToString(val));
            }
        }

        public void PutLitteralPercentValue(Sheets sn, int row, int column, string val)
        {
            IXLWorksheet asheet = getSheet(sn);
            if (!(val is null))
            {
                asheet.Cell(row, column).SetValue<string>(Convert.ToString(val));
            }
        }

        public void PutDoubleValueInSheetRowColumn(Sheets sn, int row, int column, double fval)
        {
            IXLWorksheet asheet = getSheet(sn);
            asheet.Cell(row, column).Value = fval;
        }
        public void PutValueInSheetRowColumn(Sheets sn, int row, int column, string val)
        {
            IXLWorksheet asheet = getSheet(sn);
            if (!(val is null))
            {
                asheet.Cell(row, column).Value = val;
            }
        }

        public void setSheetCellWrapText(Sheets sn, bool onoff, int columns, int rows, int rowtofreez)
        {
            IXLWorksheet asheet = getSheet(sn);
            IXLRange rrr = asheet.Range(asheet.Cell(rows, columns).Address, asheet.Cell(rows, columns).Address);
            rrr.Select();
            rrr.Style.Alignment.WrapText = onoff;
            asheet.Columns(columns, columns).AdjustToContents();

            asheet.SheetView.FreezeRows(rowtofreez);
        }

        public void putParamsToTable(int row, string nesachType, string gush, string helka)
        {
            xlFilesWorkSheet.Cell(row, 2).Value = gush;
            xlFilesWorkSheet.Cell(row,3).Value = helka;
            xlFilesWorkSheet.Cell(row, 4).Value = nesachType;
        }

        public void refreshAll()
        {
        }

        public void setBoarder(Sheets sn, int irow1, int irow2, int icol1, int icol2, int thickness)
        {
            XLBorderStyleValues  thick;
            if (thickness == 0)
            {
                thick = XLBorderStyleValues.Hair;
            }
            else if (thickness == 1)
            {
                thick = XLBorderStyleValues.Thin;
            }
            else if (thickness == 2)
            {
                thick = XLBorderStyleValues.Medium;
            }
            else if (thickness == 3)
            {
                thick = XLBorderStyleValues.Thick;
            }
            else
            {
                thick = XLBorderStyleValues.Medium;
            }

            IXLWorksheet asheet = getSheet(sn);
            IXLRange frameTop = asheet.Range(asheet.Cell(irow1, icol1).Address, asheet.Cell(irow1, icol2).Address);
            IXLRange frameBottom = asheet.Range(asheet.Cell(irow2, icol1).Address, asheet.Cell(irow2, icol2).Address);
            IXLRange frameRight = asheet.Range(asheet.Cell(irow1, icol2).Address, asheet.Cell(irow2, icol2).Address);
            IXLRange frameleft = asheet.Range(asheet.Cell(irow1, icol1).Address, asheet.Cell(irow2, icol1).Address);

            frameTop.Style.Border.TopBorder = thick;
            frameleft.Style.Border.LeftBorder = thick;
            frameBottom.Style.Border.BottomBorder = thick;
            frameRight.Style.Border.RightBorder = thick;

//            frame.BorderAround2(Type.Missing, thick, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
//            frame.EntireColumn.BorderAround2(XlLineStyle.xlContinuous, XLBorderStyleValues.Medium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
        }
        public void addNameRange(Sheets sn, int irow1, int irow2, int icol1, int icol2, int gush, int helka, int tat, string prefix)
        {
            string sss = prefix + "_" + gush.ToString() + "_" + helka.ToString() + "_" + tat.ToString();
            IXLWorksheet asheet = getSheet(sn);
            IXLRange frame = asheet.Range(asheet.Cell(irow1, icol1).Address, asheet.Cell(irow2, icol2).Address);
            frame.AddToNamed(sss);
        }
        public void createHyperLink(Sheets sn, int row, int column, string gush, string helka, string tat, string prefix)
        {
            string subAddress = prefix + "_" + gush + "_" + helka + "_" + tat;
            IXLWorksheet asheet = getSheet(sn);
            XLHyperlink hyper1 = new XLHyperlink(subAddress);

            asheet.Cell(row, column).Value = "X";
            asheet.Cell(row, column).SetHyperlink(hyper1);

        }
        public void paintRow(Sheets sn, int irow, int icol, int endcol, System.Drawing.Color Color)
        {
            IXLWorksheet asheet = getSheet(sn);
            IXLRange Rang = asheet.Range(asheet.Cell(irow, endcol).Address, asheet.Cell(irow, endcol).Address);
            Rang.Style.Fill.BackgroundColor = XLColor.FromColor(Color);
        }
        public void CorrectFormatForSum(Sheets sn, int columns, int rows1, int rows2, string numFormat)
        {
            IXLWorksheet asheet = getSheet(sn);

            IXLRange Rang = asheet.Range(asheet.Cell(rows1, columns).Address, asheet.Cell(rows2, columns).Address);
            IXLRange oTargetRange = asheet.Range(asheet.Cell(rows1, columns).Address, asheet.Cell(rows2, columns).Address);

            //try
            //{
            //    Rang.TextToColumns(oTargetRange, XlTextParsingType.xlDelimited,
            //        XlTextQualifier.xlTextQualifierDoubleQuote, false, true, false, false, false, true, "-");
            //}
            //catch (Exception ex)
            //{

            //}
            //rrr.NumberFormat = numFormat;

        }

        public void ListPdfFiles(string[] files)
        {
            CreateTitle(1, 1, 1, "שם נסח", 40.0);
            CreateTitle(1, 2, 2, "גוש", 10.0);
            CreateTitle(1, 3, 3, "חלקה", 10.0);
            CreateTitle(1, 4, 4, "סוג נסח", 15.0);
            int startRow = 2;
            PDFFolder = Path.GetDirectoryName(files[0]);
            for (int i = 0; i < files.Length; i++)
            {
                string result = Path.GetFileName(files[i]);
                xlFilesWorkSheet.Cell(i+startRow , 1 ).Value = result;
            }
        }

        public void CreateTitle(int row, int column0, int column2, string title, double width)
        {
            string lRow = ClassUtils.ColumnLabel(row);
            string lColumn = ClassUtils.ColumnLabel(column0);
            string totalCellName = lColumn + ":" + lColumn;
            xlFilesWorkSheet.Column(lColumn).Width = width;

//            xlFilesWorkSheet.Column(totalCellName).Width = width;
            xlFilesWorkSheet.Cell(row, column0).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            xlFilesWorkSheet.Cell(row,column0).Value = title;
        }

        public Color GetFromRGB(byte r, byte g, byte b)
        {
            Color RGBColor = new Color();
            RGBColor = Color.FromArgb(r, g, b);
            return RGBColor;
        }

        public void setActiveSheet(Sheets sn)
        {

        }
        public class PutCellParameters
        {
            public bool ifmerge { get; set; }
            public int Rowextension { get; set; }
            public int Columnextension { get; set; }
            public XLColor colorbackground { get; set; }
            public XLAlignmentVerticalValues xlVAlign { get; set; }            
            public XLAlignmentHorizontalValues xlHAlign { get; set; }
            public bool ifFrame { get; set; }
            public XLBorderStyleValues Weight { get; set; }
            public int fontSize { get; set; }

        }
        public void putValueWithParameter(Sheets sheet, string value, int row, int col, PutCellParameters param)
        {
            IXLWorksheet asheet = getSheet(sheet);
            IXLRange  sellection;
            if (param.ifmerge)
            {
                sellection = asheet.Range(asheet.Cell(row, col).Address, asheet.Cell(row + param.Rowextension - 1, col + param.Columnextension - 1).Address);
                sellection.Merge();
            }
            else
            {
                sellection = asheet.Range(asheet.Cell(row, col).Address, asheet.Cell(row, col).Address);
            }
            sellection.Style.Alignment.Vertical = param.xlVAlign;
            sellection.Style.Alignment.Horizontal = param.xlHAlign;
            sellection.Style.Fill.BackgroundColor = param.colorbackground;

            sellection.Style.Border.BottomBorder = param.Weight;
            sellection.Style.Border.BottomBorderColor = param.colorbackground;
            sellection.Style.Border.TopBorder = param.Weight;
            sellection.Style.Border.TopBorderColor = param.colorbackground;
            sellection.Style.Border.LeftBorder = param.Weight;
            sellection.Style.Border.LeftBorderColor = param.colorbackground;
            sellection.Style.Border.RightBorder = param.Weight;
            sellection.Style.Border.RightBorderColor = param.colorbackground;
            sellection.Style.Font.FontSize =  param.fontSize;
            asheet.Cell(row, col).Value = value;
            sellection.Select();
        }

        public void setColumnsAlignments(Sheets sheet, int startcol, int endcol, XLAlignmentHorizontalValues alValue)
        {
            IXLWorksheet asheet = getSheet(sheet);
            asheet.Columns(startcol, endcol).Style.Alignment.Horizontal = alValue;

        }

        public void adjustToContent(Sheets sheet, int startcol, int endcol)
        {
            IXLWorksheet asheet = getSheet(sheet);
            asheet.Columns(startcol, endcol).AdjustToContents();
        }

        public void exitForDebug()
        {
            string resultfile = "c:\\tmp\\Tabu_results_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            saveResults(resultfile);
            disposeAllSheets();
        }

        public void testHyperlink()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Hyperlinks");
            wb.Worksheets.Add("Second Sheet");

            Int32 ro = 0;


            


            ws.Cell(++ro, 1).Value = "Link to a web page, no tooltip - Yahoo!";            
            XLHyperlink hyper1 = new XLHyperlink(@"http://www.yahoo.com");
            ws.Cell(ro, 1).SetHyperlink(hyper1);

            ws.Cell(++ro, 1).Value = "Link to a web page, with a tooltip - Yahoo!";
            XLHyperlink hyper2 = new XLHyperlink(@"http://www.yahoo.com", "Click to go to Yahoo!");
            ws.Cell(ro, 1).SetHyperlink(hyper2);

            ws.Cell(++ro, 1).Value = "Link to an address in another worksheet";
            XLHyperlink hyper3 = new XLHyperlink("'Second Sheet'!A1:C5");
            ws.Cell(ro, 1).SetHyperlink(hyper3);


            ws.Columns().AdjustToContents();

            wb.SaveAs("Hyperlinks.xlsx");
        }

    }
}
