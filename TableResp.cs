using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectClosureToolMVVM
{
    public partial class TableResp
    {
        //public static double minimumSize;
        //public static ExcelPackage Excel_result { get; set; }
        //public static double MinimumSize { get; set; }

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
        //        for (int iU = 1; iU <= Program.distinctUnitsList.Count; iU++)
        //        {
        //            estimateWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
        //            pointWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
        //            estimateWorksheet.Cells[iU + 1, 1].Value = Program.distinctUnitsList[iU - 1];
        //            pointWorksheet.Cells[iU + 1, 1].Value = Program.distinctUnitsList[iU - 1];
        //            for (int iC = 1; iC <= Program.distinctCombinationsListI.Count; iC++)
        //            {
        //                estimateWorksheet.Cells[1, iC + 1].Style.WrapText = true;
        //                pointWorksheet.Cells[1, iC + 1].Style.WrapText = true;
        //                if (iU == Program.distinctUnitsList.Count)
        //                {
        //                    estimateWorksheet.Cells[1, iC + 1].Value = Program.distinctCombinationsListI[iC - 1];
        //                    pointWorksheet.Cells[1, iC + 1].Value = Program.distinctCombinationsListI[iC - 1];
        //                }
        //                estimateWorksheet.Column(iC).Width = 19;
        //                estimateWorksheet.Column(iC + 1).Width = 19;
        //                pointWorksheet.Column(iC).Width = 19;
        //                pointWorksheet.Column(iC + 1).Width = 19;
        //            }
        //        }
        //        excel_result.SaveAs(fout);
        //        for (int iU = 1; iU <= Program.distinctUnitsList.Count; iU++)
        //        {
        //            for (int iC = 1; iC <= Program.distinctCombinationsListI.Count; iC++)
        //            {
        //                string unit = Program.distinctUnitsList.ElementAt(iU - 1);
        //                string combination = Program.distinctCombinationsListI.ElementAt(iC - 1);
        //                estimateWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
        //                pointWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
        //                foreach (TrelloObjectSums aSum in Program.sums)
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
