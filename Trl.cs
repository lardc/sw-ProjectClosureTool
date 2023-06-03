using System;
using System.IO;
using System.Linq;

namespace ProjectClosureToolV2
{
    public partial class Trl : TableResp
    {
        public static int iLabels = 0;
        public static int iNone = 0;

        // ввод текущей карточки
        public static string currentCardURL;
        public static string currentCardUnit;
        public static string currentCardName;
        public static double currentCardEstimate;
        public static double currentCardPoint;
        public static string[] cardLabels = new string[20];

        public static void SearchLabelsM(string rr)
        {
            cardLabels[iLabels] = rr;
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
                        try { currentCardEstimate = d_ior; }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine("Press any key");
                            Console.ReadKey();
                            return;
                        }
                    }
                    if (iOpeningBracket > iDot && iClosingBracket > iDot)
                    {
                        string s_uisq = rr.Substring(iOpeningBracket + 1, iClosingBracket - iOpeningBracket - 1).Trim();
                        double d_isq = double.Parse(s_uisq);
                        try { currentCardPoint = d_isq; }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine("Press any key");
                            Console.ReadKey();
                            return;
                        }
                    }
                    string s_unit = rr.Substring(0, iDot).Trim();
                    try 
                    { 
                        currentCardUnit = s_unit;
                        Program.units.Add(currentCardUnit);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
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
                    try { currentCardName = s_name; }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
                }
                else currentCardName = rr;
                if (currentCardEstimate == 0 && currentCardPoint == 0 && ((iDot <= Math.Min(iOpeningParenthesis, iOpeningBracket) && Math.Min(iOpeningParenthesis, iOpeningBracket) >= 0) || (Math.Max(iOpeningParenthesis, iOpeningBracket) < 0))) iNone++;
            }
        }
        public static void EMessage(string eM)
        {
            int s_err = eM.IndexOf("error", 0);
            int s_dot = eM.IndexOf(".", 0);
            string s_mess = eM.Substring(s_err, s_dot - s_err).Trim();
            Console.WriteLine(s_mess);
        }
    }
}