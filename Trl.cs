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
        public static List<TrelloObjectSums> sums = [];
        private static bool sumsListFilled = false;

        public static int iLabels = 0;

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
                    }
                    if (iOpeningBracket > iDot && iClosingBracket > iDot)
                    {
                        string s_uisq = rr.Substring(iOpeningBracket + 1, iClosingBracket - iOpeningBracket - 1).Trim();
                        double d_isq = double.Parse(s_uisq);
                        Download.currentCardPoint = d_isq;
                    }
                    string s_unit = rr[..iDot].Trim();
                    Download.currentCardUnit = s_unit;
                    Download.units.Add(Download.currentCardUnit);
                    Download.unitsListFilled = true;
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
                }
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
    }
}
