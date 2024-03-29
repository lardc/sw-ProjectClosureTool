﻿using System;
using System.IO;
using System.Linq;

namespace TrlConsCs
{
    public partial class Trl : TableResp
    {
        public const int iMaxParse = 1000; // максимальное к-во обрабатываемых ошибок в карточке

        public static bool parseBadgesToken;
        public static int iErrorParse;
        public static string[] parseStrErrorMessage = new string[iMaxParse];
        public static string[] parseStrErrorCardURL = new string[iMaxParse];
        public static double[] strCorrectEstimate = new double[iAllUnits]; // приближённое значение в корректной карточке
        public static double[] strCorrectPoint = new double[iAllUnits]; // точное значение в корректной карточке
        public static string[] strCorrectShotrURL = new string[iAllUnits];
        public static int iCorrectParse = 0;

        public static void ClearParsedCardErrorData()
        {
            parseBadgesToken = false;
            for (int i = 0; i < iMaxParse; i++)
            {
                parseStrErrorMessage[i] = "";
                parseStrErrorCardURL[i] = "";
            }
            iErrorParse = -1;
        }

        // Формирование таблиц оценочных и реальных значений
        public static void FillParsedValues()
        {
            if (iCurrentDepartment < 0 && iCurrentUnit >= 0)
            {
                if (iErrorParse < iMaxParse) iErrorParse++;
                parseStrErrorMessage[iErrorParse] = "Нет ярлыка стадии";
                parseStrErrorCardURL[iErrorParse] = currentShortUrl;
            }

            if (iCurrentTeam < 0 && iCurrentUnit >= 0)
            {
                if (iErrorParse < iMaxParse) iErrorParse++;
                parseStrErrorMessage[iErrorParse] = "Нет ярлыка команды";
                parseStrErrorCardURL[iErrorParse] = currentShortUrl;
            }

            if (iCurrentDepartment >= 0 && iCurrentTeam >= 0 && iCurrentUnit >= 0)
            {
                if (Curr_Estim > 0)
                {
                    FillCurrentEstimate();
                }
                if (Curr_Point > 0)
                {
                    FillCurrentPoints();
                }
                strCorrectEstimate[iCorrectParse] = Curr_Estim;
                strCorrectPoint[iCorrectParse] = Curr_Point;
                strCorrectShotrURL[iCorrectParse] = currentShortUrl;
                if (iCorrectParse < iAllUnits) iCorrectParse++;
            }

            FillParsedErrors();
            ClearCurrentUnit();
        }

        // Формирование таблицы ошибок для обработанной карточки
        public static void FillParsedErrors()
        {
            if (iLabelsEM == 0)
            {
                int i;
                for (i = 0; i <= iErrorParse; i++)
                {
                    strErrNumber[iError] = iError + 1;
                    strErrMessage[iError] = parseStrErrorMessage[i];
                    strErrCardURL[iError] = currentShortUrl;

                    iError++;
                }
            }
            ClearParsedCardErrorData();
        }

        public static void FillCurrentEstimate()
        {
            iTotal = iCurrentDepartment * (iTeams + 1) + iTeams;
            iCurrTD = iCurrentDepartment * (iTeams + 1) + iCurrentTeam;
            CurrUnit = iCurrentUnit + 1;
            Xi[0, iCurrTD, CurrUnit] += Curr_Estim;
            Xi[0, iCurrTD, 0] += Curr_Estim;
            Xi[0, iTotal, 0] += Curr_Estim;
        }

        public static void FillCurrentPoints()
        {
            iTotal = iCurrentDepartment * (iTeams + 1) + iTeams;
            iCurrTD = iCurrentDepartment * (iTeams + 1) + iCurrentTeam;
            CurrUnit = iCurrentUnit + 1;
            Xi[1, iCurrTD, CurrUnit] += Curr_Point;
            Xi[1, iCurrTD, 0] += Curr_Point;
            Xi[1, iTotal, 0] += Curr_Point;
        }

        //Очистка значений счётчиков по текущему блоку
        public static void ClearCurrentUnit()
        {
            iCurrentDepartment = -1;
            iCurrentTeam = -1;
            iCurrentUnit = -1;
            Curr_Estim = 0;
            Curr_Point = 0;
            iLabelsEM = 0; // Служебная карточка: iLabelsEM = 1
            iStartError = iError;
            badges = 0;
            cardRole = 0;
            currentShortUrl = "";
        }

        // Формирование списка стадий и команд
        public static void Search_Departments_Teams(string rr)
        {
            if (Departments.Contains(rr)) Search_Departments(rr);
            else if (Teams.Contains(rr)) Search_Teams(rr);
        }

        // Формирование списка стадий
        public static void Search_Departments(string rr)
        {
            try
            {
                if (Departments.Contains(rr))
                {
                    if (iCurrentDepartment < 0) iCurrentDepartment = Array.IndexOf(Departments, rr);
                    // уже есть стадия к текущей карточке
                    else
                    {
                        if (iErrorParse < iMaxParse) iErrorParse++;
                        parseStrErrorMessage[iErrorParse] = $"Повтор: карточке соответствует более одной стадии";
                        parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                    };
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }

        // Формирование списка команд
        public static void Search_Teams(string rr)
        {
            try
            {
                if (Teams.Contains(rr))
                {
                    if (iCurrentTeam < 0) iCurrentTeam = Array.IndexOf(Teams, rr);
                    // уже есть команда к текущей карточке
                    else
                    {
                        if (iErrorParse < iMaxParse) iErrorParse++;
                        parseStrErrorMessage[iErrorParse] = $"Повтор: карточке соответствует более одной команды";
                        parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                    };
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }

        // Формирование списка блоков
        public static void Search_Unit(string rr)
        {
            if (iAll < iAllUnits)
            {
                try
                {
                    if (All_Units.Contains(rr)) iCurrentUnit = Array.IndexOf(All_Units, rr);
                    else
                    {
                        iCurrentUnit = iAll;
                        All_Units[iCurrentUnit] = rr;
                        iAll++;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
            }
            else
            {
                if (iErrorParse < iMaxParse) iErrorParse++;
                parseStrErrorMessage[iErrorParse] = $"Количество блоков превышено ( >{iAllUnits} )";
                parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                Console.WriteLine("Количество блоков превышено");
            }
        }

        // Запись оценочных и реальных значений для обнаруженного блока
        public static void Fill_Current_Values(string rr)
        {
            int iDot = rr.IndexOf('.', 0); // позиция точки в строке rr
            int iOpeningParenthesis = rr.IndexOf('(', 0); // позиция открывающейся круглой скобки в строке rr
            int iClosingParenthesis = rr.IndexOf(')', 0); // позиция закрывающейся круглой скобки в строке rr
            int iOpeningBracket = rr.IndexOf('[', 0); // позиция открывающейся квадратной скобки в строке rr
            int iClosingBracket = rr.IndexOf(']', 0); // позиция закрывающейся квадратной скобки в строке rr

            if (iDot >= 0)
            {
                if (iOpeningParenthesis >= 0 && iClosingParenthesis < iOpeningParenthesis)
                {
                    if (iErrorParse < iMaxParse) iErrorParse++;
                    parseStrErrorMessage[iErrorParse] = $"В карточке <{rr}> нет закрывающей круглой скобки";
                    parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                }
                if (iOpeningBracket >= 0 && iClosingBracket < iOpeningBracket)
                {
                    if (iErrorParse < iMaxParse) iErrorParse++;
                    parseStrErrorMessage[iErrorParse] = $"В карточке <{rr}> нет закрывающей квадратной скобки";
                    parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                }

                if (iOpeningParenthesis >= iDot && iClosingParenthesis >= iDot)
                {
                    string s_uiro = rr.Substring(iOpeningParenthesis + 1, iClosingParenthesis - iOpeningParenthesis - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (iOpeningBracket >= iDot && iClosingBracket >= iDot)
                {
                    string s_uisq = rr.Substring(iOpeningBracket + 1, iClosingBracket - iOpeningBracket - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
                if (Curr_Estim == 0 && Curr_Point == 0 && ((iDot <= Math.Min(iOpeningParenthesis, iOpeningBracket) && Math.Min(iOpeningParenthesis, iOpeningBracket) >= 0) || (Math.Max(iOpeningParenthesis, iOpeningBracket) < 0)))
                {
                    if (iErrorParse < iMaxParse) iErrorParse++;
                    parseStrErrorMessage[iErrorParse] = $"В карточке <{rr}> нет значений";
                    parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                    iNone++;
                }
                if (iDot >= Math.Max(iOpeningParenthesis, iOpeningBracket) && Math.Max(iOpeningParenthesis, iOpeningBracket) > 0)
                {
                    if (iErrorParse < iMaxParse) iErrorParse++;
                    parseStrErrorMessage[iErrorParse] = $"В карточке <{rr}> нет имени блока";
                    parseStrErrorCardURL[iErrorParse] = currentShortUrl;
                }
                else
                {
                    string s_unit = rr.Substring(0, iDot).Trim();
                    Search_Unit(s_unit);
                }
            }
            // Обнаружение служебной карточки
            else if (rr == "Labels" || rr == "ЭМ")
            {
                Console.WriteLine($"Служебная карточка <{rr}>");
                iLabelsEM = 1;
            }
            else
            {
                if (iErrorParse < iMaxParse) iErrorParse++;
                parseStrErrorMessage[iErrorParse] = $"В карточке <{rr}> нет имени блока";
                parseStrErrorCardURL[iErrorParse] = currentShortUrl;
            }
        }

        // Запись оценочных и реальных значений в Excel-файл
        public static void FillExcel()
        {
            if (iAll == iNone)
            {
                Console.WriteLine($"В доске {API_Req.boardURL}{API_Req.boardCode} нет ни одного значения. Файл не сформирован.");
                return;
            }
            string fstrin = "json_into_xlsx_t.xltx";
            if (!File.Exists(fstrin))
            {
                Console.WriteLine("Шаблон не существует");
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
            else
            {
                FillExcel_fstrout(fstrin);
                FileInfo fin = new(fstrin);
                if (File.Exists(FillExcel_fstrout(fstrin)))
                {
                    try { File.Delete(FillExcel_fstrout(fstrin)); }
                    catch (IOException deleteError)
                    {
                        Console.WriteLine(deleteError.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
                }
                FileInfo fout = new(FillExcel_fstrout(fstrin));
                using (Excel_result = new OfficeOpenXml.ExcelPackage(fout, fin))
                {
                    Excel_result.Workbook.Properties.Author = "KM";
                    Excel_result.Workbook.Properties.Title = "Trello";
                    Excel_result.Workbook.Properties.Created = DateTime.Now;
                    MinimumSize = 6;
                    FillExcelSheets(3);
                    Excel_result.Save();
                }
            }
        }
        public static void EMessage(string eM)
        {
            int s_err = eM.IndexOf("error", 0);
            int s_dot = eM.IndexOf(".", 0);
            string s_mess = eM.Substring(s_err, s_dot - s_err).Trim();
            Console.WriteLine(s_mess);
        }
        public static string FillExcel_fstrout(string fstrin)
        {
            DateTime DTnow = DateTime.Now;
            string DTyear = DTnow.Year.ToString();
            string DTmonth = DTnow.Month.ToString();
            string DTday = DTnow.Day.ToString();
            string DThour = DTnow.Hour.ToString();
            string DTminute = DTnow.Minute.ToString();
            string DTsecond = DTnow.Second.ToString();
            string fstrout = $"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx";
            return fstrout;
        }
    }
}