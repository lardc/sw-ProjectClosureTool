using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ProjectClosureToolMVVM
{
    public partial class Trl : TableResp
    {
        public static List<TrelloObjectSums> sums = new List<TrelloObjectSums>();
        private static bool sumsListFilled = false;

        public static int iLabels = 0;
        //public static int iNone = 0;

        //// ввод текущей карточки
        //public static string currentCardURL;
        //public static string currentCardUnit;
        //public static string currentCardName;
        //public static double currentCardEstimate;
        //public static double currentCardPoint;
        //public static string[] cardLabels = new string[20];

        public static void SearchLabelsM(string rr)
        {
            Download.cardLabels[iLabels] = rr;
            iLabels++;
        }

        public static void SearchUnitValuesM(string rr)
        {
            if (rr.Length > 0)
            {
                int iDot = rr.IndexOf('.', 0); // позиция точки в строке rr
                int iOpeningParenthesis = rr.IndexOf('(', 0); // позиция открывающейся круглой скобки в строке rr
                int iClosingParenthesis = rr.IndexOf(')', 0); // позиция закрывающейся круглой скобки в строке rr
                int iOpeningBracket = rr.IndexOf('[', 0); // позиция открывающейся квадратной скобки в строке rr
                int iClosingBracket = rr.IndexOf(']', 0); // позиция закрывающейся квадратной скобки в строке rr

                if (iDot >= 0)
                {
                    if (iOpeningParenthesis > iDot && iClosingParenthesis > iDot)
                    {
                        string s_uiro = rr.Substring(iOpeningParenthesis + 1, iClosingParenthesis - iOpeningParenthesis - 1).Trim();
                        double d_ior = double.Parse(s_uiro);
                        Download.currentCardEstimate = d_ior;
                        //try { currentCardEstimate = d_ior; }
                        //catch (Exception e)
                        //{
                        //    Console.WriteLine(e.Message);
                        //    Console.WriteLine("Press any key");
                        //    Console.ReadKey();
                        //    return;
                        //}
                    }
                    if (iOpeningBracket > iDot && iClosingBracket > iDot)
                    {
                        string s_uisq = rr.Substring(iOpeningBracket + 1, iClosingBracket - iOpeningBracket - 1).Trim();
                        double d_isq = double.Parse(s_uisq);
                        Download.currentCardPoint = d_isq;
                        //try { currentCardPoint = d_isq; }
                        //catch (Exception e)
                        //{
                        //    Console.WriteLine(e.Message);
                        //    Console.WriteLine("Press any key");
                        //    Console.ReadKey();
                        //    return;
                        //}
                    }
                    string s_unit = rr.Substring(0, iDot).Trim();
                    Download.currentCardUnit = s_unit;
                    Download.units.Add(Download.currentCardUnit);
                    Download.unitsListFilled = true;
                    //try
                    //{
                    //    currentCardUnit = s_unit;
                    //    Program.units.Add(currentCardUnit);
                    //}
                    //catch (Exception e)
                    //{
                    //    Console.WriteLine(e.Message);
                    //    Console.WriteLine("Press any key");
                    //    Console.ReadKey();
                    //    return;
                    //}
                    string s_name = "";
                    if (iOpeningParenthesis >= 0 && iOpeningBracket >= 0)
                    {
                        if (rr.Substring(iDot + 1, Math.Min(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim().Length > 0)
                            s_name = rr.Substring(iDot + 1, Math.Min(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim();
                    }
                    else if (iOpeningParenthesis < 0 && iOpeningBracket < 0)
                    {
                        s_name = rr.Substring(iDot + 1, rr.Length - iDot - 1).Trim();
                    }
                    else if (rr.Substring(iDot + 1, Math.Max(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim().Length > 0)
                        s_name = rr.Substring(iDot + 1, Math.Max(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim();
                    //try { currentCardName = s_name; }
                    //catch (Exception e)
                    //{
                    //    Console.WriteLine(e.Message);
                    //    Console.WriteLine("Press any key");
                    //    Console.ReadKey();
                    //    return;
                    //}
                }
                //else currentCardName = rr;
                //if (currentCardEstimate == 0 && currentCardPoint == 0 && ((iDot <= Math.Min(iOpeningParenthesis, iOpeningBracket) && Math.Min(iOpeningParenthesis, iOpeningBracket) >= 0) || (Math.Max(iOpeningParenthesis, iOpeningBracket) < 0))) iNone++;
            }
        }

        public static void Sum(int i, int j)
        {
            Download.LabelCombinationsI();
            Download.sumEstimate = 0;
            Download.sumPoint = 0;
            string unit = Download.distinctUnitsList.ElementAt(i);
            string combination = Download.distinctCombinationsListI.ElementAt(j - 1);
            int sumT = 0;
            foreach (TrelloObject aCard in Download.cards)
                if (aCard.CardUnit.Equals(unit) && aCard.LabelCombinationI.Equals(combination))
                {
                    Download.sumEstimate += aCard.CardEstimate;
                    Download.sumPoint += aCard.CardPoint;
                    sumT++;
                    sums.Add(new TrelloObjectSums()
                    {
                        CardUnit = unit,
                        LabelCombinationI = combination,
                        SumEstimate = Download.sumEstimate,
                        SumPoint = Download.sumPoint
                    });
                }
            if (sumT > 0) { sumsListFilled = true; sums.Distinct(); }
        }

        //public static void FillExcel()
        //{
        //    DateTime DTnow = DateTime.Now;
        //    string DTyear = DTnow.Year.ToString();
        //    string DTmonth = DTnow.Month.ToString();
        //    string DTday = DTnow.Day.ToString();
        //    string DThour = DTnow.Hour.ToString();
        //    string DTminute = DTnow.Minute.ToString();
        //    string DTsecond = DTnow.Second.ToString();
        //    string fileName = $"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx";
        //    using (ExcelPackage excel_result = new ExcelPackage())
        //    {
        //        excel_result.Workbook.Worksheets.Add("Estimations");
        //        excel_result.Workbook.Worksheets.Add("Points");
        //        FileInfo fout = new FileInfo(fileName);
        //        if (fout.Exists)
        //        {
        //            fout.Delete();
        //            fout = new FileInfo(@"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx");
        //        }
        //        excel_result.SaveAs(fout);
        //        var estimateWorksheet = excel_result.Workbook.Worksheets["Estimations"];
        //        var pointWorksheet = excel_result.Workbook.Worksheets["Points"];
        //        for (int iU = 1; iU <= Download.distinctUnitsList.Count; iU++)
        //        {
        //            estimateWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
        //            pointWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
        //            estimateWorksheet.Cells[iU + 1, 1].Value = Download.distinctUnitsList[iU - 1];
        //            pointWorksheet.Cells[iU + 1, 1].Value = Download.distinctUnitsList[iU - 1];
        //            for (int iC = 1; iC <= Download.distinctCombinationsListI.Count; iC++)
        //            {
        //                estimateWorksheet.Cells[1, iC + 1].Style.WrapText = true;
        //                pointWorksheet.Cells[1, iC + 1].Style.WrapText = true;
        //                if (iU == Download.distinctUnitsList.Count)
        //                {
        //                    estimateWorksheet.Cells[1, iC + 1].Value = Download.distinctCombinationsListI[iC - 1];
        //                    pointWorksheet.Cells[1, iC + 1].Value = Download.distinctCombinationsListI[iC - 1];
        //                }
        //                estimateWorksheet.Column(iC).Width = 19;
        //                estimateWorksheet.Column(iC + 1).Width = 19;
        //                pointWorksheet.Column(iC).Width = 19;
        //                pointWorksheet.Column(iC + 1).Width = 19;
        //            }
        //        }
        //        excel_result.SaveAs(fout);
        //        for (int iU = 1; iU <= Download.distinctUnitsList.Count; iU++)
        //        {
        //            for (int iC = 1; iC <= Download.distinctCombinationsListI.Count; iC++)
        //            {
        //                string unit = Download.distinctUnitsList.ElementAt(iU - 1);
        //                string combination = Download.distinctCombinationsListI.ElementAt(iC - 1);
        //                estimateWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
        //                pointWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
        //                foreach (TrelloObjectSums aSum in sums)
        //                    if (aSum.CardUnit.Equals(unit) && aSum.LabelCombinationI.Equals(combination))
        //                    {
        //                        estimateWorksheet.Cells[iU + 1, iC + 1].Value = aSum.SumEstimate;
        //                        pointWorksheet.Cells[iU + 1, iC + 1].Value = aSum.SumPoint;
        //                    }
        //            }
        //        }
        //        excel_result.SaveAs(fout);
        //    }
        //}
    }
}
