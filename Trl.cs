using System;
using System.IO;
using System.Linq;

namespace TrlConsCs
{
    public partial class Trl : TableResp
    {
        public const int iMaxParse = 1000; // максимальное к-во обрабатываемых ошибок в карточке
        public static bool parseErrUnit;

        public static string parseUnitToken;
        public static string parseShortUrlToken;
        public static bool parseBadgesToken;
        public static bool parseCardRoleToken;
        public static int iErrParse;
        public static string[] parseStrErrMessage = new string[iMaxParse];
        public static string[] parseStrErrCardURL = new string[iMaxParse];

        public static void ParseClear()
        {
            parseUnitToken = "";
            parseShortUrlToken = "";
            parseBadgesToken = false;
            parseCardRoleToken = false;
            for (int i = 0; i < iMaxParse; i++)
            {
                parseStrErrMessage[i] = "";
                parseStrErrCardURL[i] = "";
            }
            iErrParse = -1;
        }

        // Формирование таблиц оценочных и реальных значений
        public static void XiParse()
        {
            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0) XiFill_Curr_Estim();
                if (Curr_Point > 0) XiFill_Curr_Point();
            }

            if (iCurr_Depart < 0 && iCurr_Unit >= 0)
            {
                if (iErrParse < iMaxParse) iErrParse++;
                parseStrErrMessage[iErrParse] = "Нет ярлыка стадии";
                parseStrErrCardURL[iErrParse] = currShortUrl;
                parseErrUnit = true;
            }

            if (iCurr_Team < 0 && iCurr_Unit >= 0)
            {
                if (iErrParse < iMaxParse) iErrParse++;
                parseStrErrMessage[iErrParse] = "Нет ярлыка команды";
                parseStrErrCardURL[iErrParse] = currShortUrl;
                parseErrUnit = true;
            }
            FillParseErr();
            Curr_Clear();
        }

        //public static void FillParseErr(string currShortUrl)
        public static void FillParseErr()
        {
            for (int i = 0; i <= iErrParse && iLabelsEM == 0; i++)
            {
                strErrNumber[iErr] = iErr+1;
                strErrMessage[iErr] = parseStrErrMessage[i];
                strErrCardURL[iErr] = parseStrErrCardURL[iErrParse];
                iErr++;
            }
            if (iErr > 0) strErrCardURL[iErr - 1] = currShortUrl;
            ParseClear();
        }

        public static void XiFill_Curr_Estim()
        {
            iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
            iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
            CurrUnit = iCurr_Unit + 1;
            Xi[0, iCurrTD, CurrUnit] += Curr_Estim;
            Xi[0, iCurrTD, 0] += Curr_Estim;
            Xi[0, iTotal, 0] += Curr_Estim;
        }

        public static void XiFill_Curr_Point()
        {
            iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
            iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
            CurrUnit = iCurr_Unit + 1;
            Xi[1, iCurrTD, CurrUnit] += Curr_Point;
            Xi[1, iCurrTD, 0] += Curr_Point;
            Xi[1, iTotal, 0] += Curr_Point;
        }

        //Очистка значений счётчиков по текущему блоку
        public static void Curr_Clear()
        {
            iCurr_Depart = -1;
            iCurr_Team = -1;
            iCurr_Unit = -1;
            Curr_Estim = 0;
            Curr_Point = 0;
            iLabelsEM = 0; // Служебная карточка
            iStartErr = iErr;
            badges = 0;
            cardRole = 0;
        }

        // Формирование списка стадий и команд
        public static void Search_Depart_Teams(string rr)
        {
            if (Departments.Contains(rr)) Search_Departments(rr);
            else if (Teams.Contains(rr)) Search_Teams(rr);
            else
            {
                if (iErrParse < iMaxParse) iErrParse++;
                parseStrErrMessage[iErrParse] = $"Ярлык <{rr}> не соответствует полям таблицы";
                parseStrErrCardURL[iErrParse] = currShortUrl;
                parseErrUnit = true;
            }
        }

        // Формирование списка стадий
        public static void Search_Departments(string rr)
        {
            try
            {
                if (Departments.Contains(rr))
                {
                    if (iCurr_Depart < 0) iCurr_Depart = Array.IndexOf(Departments, rr);
                    // уже есть стадия к текущей карточке
                    else
                    {
                        if (iErrParse < iMaxParse) iErrParse++;
                        parseStrErrMessage[iErrParse] = $"Повтор: карточке соответствует более одной стадии";
                        parseStrErrCardURL[iErrParse] = currShortUrl;
                        parseErrUnit = true;
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
                    if (iCurr_Team < 0) iCurr_Team = Array.IndexOf(Teams, rr);
                    // уже есть команда к текущей карточке
                    else
                    {
                        if (iErrParse < iMaxParse) iErrParse++;
                        parseStrErrMessage[iErrParse] = $"Повтор: карточке соответствует более одной команды";
                        parseStrErrCardURL[iErrParse] = currShortUrl;
                        parseErrUnit = true;
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
                    if (All_Units.Contains(rr)) iCurr_Unit = Array.IndexOf(All_Units, rr);
                    else
                    {
                        iCurr_Unit = iAll;
                        All_Units[iCurr_Unit] = rr;
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
                if (iErrParse < iMaxParse) iErrParse++;
                parseStrErrMessage[iErrParse] = $"Количество блоков превышено ( >{iAllUnits} )";
                parseStrErrCardURL[iErrParse] = currShortUrl;
                Console.WriteLine("Количество блоков превышено");
                parseErrUnit = true;
            }
        }

        // Запись оценочных и реальных значений для обнаруженного блока
        public static void Fill_Unit_Curr_Val(string rr)
        {
            int iDot = rr.IndexOf('.', 0);
            int iOPar = rr.IndexOf('(', 0);
            int iCPar = rr.IndexOf(')', 0);
            int iOBr = rr.IndexOf('[', 0);
            int iCBr = rr.IndexOf(']', 0);

            if (iDot >= 0)
            {
                if (iOPar >= 0 && iCPar < iOPar)
                {
                    if (iErrParse < iMaxParse) iErrParse++;
                    parseStrErrMessage[iErrParse] = $"В карточке <{rr}> нет закрывающей круглой скобки";
                    parseStrErrCardURL[iErrParse] = currShortUrl;
                    parseErrUnit = true;
                }
                if (iOBr >= 0 && iCBr < iOBr)
                {
                    if (iErrParse < iMaxParse) iErrParse++;
                    parseStrErrMessage[iErrParse] = $"В карточке <{rr}> нет закрывающей квадратной скобки";
                    parseStrErrCardURL[iErrParse] = currShortUrl;
                    parseErrUnit = true;
                }

                if (iOPar >= iDot && iCPar >= iDot)
                {
                    string s_uiro = rr.Substring(iOPar + 1, iCPar - iOPar - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (iOBr >= iDot && iCBr >= iDot)
                {
                    string s_uisq = rr.Substring(iOBr + 1, iCBr - iOBr - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
                if (Curr_Estim == 0 && Curr_Point == 0 && ((iDot <= Math.Min(iOPar, iOBr) && Math.Min(iOPar, iOBr) >= 0) || (Math.Max(iOPar, iOBr) < 0)))
                {
                    if (iErrParse < iMaxParse) iErrParse++;
                    parseStrErrMessage[iErrParse] = $"В карточке <{rr}> нет значений";
                    parseStrErrCardURL[iErrParse] = currShortUrl;
                    iNone++;
                    parseErrUnit = true;
                }
                if (iDot >= Math.Max(iOPar, iOBr) && Math.Max(iOPar, iOBr) > 0)
                {
                    if (iErrParse < iMaxParse) iErrParse++;
                    parseStrErrMessage[iErrParse] = $"В карточке <{rr}> нет имени блока";
                    parseStrErrCardURL[iErrParse] = currShortUrl;
                    parseErrUnit = true;
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
                if (iErrParse < iMaxParse) iErrParse++;
                parseStrErrMessage[iErrParse] = $"В карточке <{rr}> нет имени блока";
                parseStrErrCardURL[iErrParse] = currShortUrl;
                parseErrUnit = true;
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
                    FillExcelSheets(2);
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