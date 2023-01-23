using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrlConsCs
{
    public partial class Trl : TableResp
    {
        // Проверка наличия зафиксированных стадий, команд или блоков
        public static bool DepartTeamUnitExist()
        {
            return (iCurr_Depart >= 0 || iCurr_Team >= 0 || iCurr_Unit >= 0);
        }

        // Проверка наличия начала карточки
        public static void Check_bgs()
        {
            if (Trl.bgs == 0)
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = "Нет начала карточки";
                if (iLa < 1) { strErrCardURL[iErr] = currShortUrl; }
            }
        }

        // Проверка наличия конца карточки
        public static void Check_cRe()
        {
            if (Trl.cRe == 0)
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = "Нет конца карточки";
                if (iLa < 1) { strErrCardURL[iErr] = currShortUrl; }
            }
        }

        // Формирование таблиц оценочных и реальных значений
        public static void XiFill()
        {
            if (DepartTeamUnitExist())
            {
                Check_bgs();
                Check_cRe();
            }
            if (errUnit) 
            {
                strErrNumber[iErr++] = iErr;
                strErrCardURL[iErr] = currShortUrl;
            }

            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0) { XiFill_Curr_Estim(); }
                if (Curr_Point > 0) { XiFill_Curr_Point(); }
            }

            if (iCurr_Depart < 0 && iCurr_Unit >= 0)
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = "Нет ярлыка стадии";
                strErrCardURL[iErr] = currShortUrl;
            }

            if (iCurr_Team < 0 && iCurr_Unit >= 0)
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = "Нет ярлыка команды";
                strErrCardURL[iErr] = currShortUrl;
            }

            Curr_Clear();
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
            iLa = 0; // Служебная карточка
            errUnit = false;
            iStartErr = iErr;
            bgs = 0;
            cRe = 0;
        }

        // Формирование списка стадий и команд
        public static void Search_Depart_Teams(string rr)
        {
            if (Departments.Contains(rr)) { Search_Departments(rr); }
            else if (Teams.Contains(rr)) { Search_Teams(rr); }
            else
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = $"Ярлык <{rr}> не соответствует полям таблицы";
                strErrCardURL[iErr] = currShortUrl;
                errUnit = true;
            }
        }

        // Формирование списка стадий
        public static void Search_Departments(string rr)
        {
            try
            {
                if (Departments.Contains(rr))
                {
                    if (iCurr_Depart < 0) { iCurr_Depart = Array.IndexOf(Departments, rr); }
                    // уже есть стадия к текущей карточке
                    else
                    {
                        strErrNumber[iErr++] = iErr;
                        strErrMessage[iErr] = $"Повтор: карточке соответствует более одной стадии";
                        strErrCardURL[iErr] = currShortUrl;
                        errUnit = true;
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
                    if (iCurr_Team < 0) { iCurr_Team = Array.IndexOf(Teams, rr); }
                    // уже есть команда к текущей карточке
                    else
                    {
                        strErrNumber[iErr++] = iErr;
                        strErrMessage[iErr++] = $"Повтор: карточке соответствует более одной команды";
                        strErrCardURL[iErr++] = currShortUrl;
                        errUnit = true;
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
                    if (All_Units.Contains(rr)) { iCurr_Unit = Array.IndexOf(All_Units, rr); }
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
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = "Количество блоков превышено ( >" + iAllUnits + ")";
                strErrCardURL[iErr] = currShortUrl;
                errUnit = true;
                Console.WriteLine("Количество блоков превышено");
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
                    strErrNumber[iErr++] = iErr;
                    strErrMessage[iErr] = $"В карточке <{rr}> нет закрывающей круглой скобки";
                    strErrCardURL[iErr] = currShortUrl;
                }
                if (iOBr >= 0 && iCBr < iOBr)
                {
                    strErrNumber[iErr++] = iErr;
                    strErrMessage[iErr] = $"В карточке <{rr}> нет закрывающей квадратной скобки";
                    strErrCardURL[iErr] = currShortUrl;
                }

                if (iOPar >= iDot && iCPar >= iDot && errUnit == false)
                {
                    string s_uiro = rr.Substring(iOPar + 1, iCPar - iOPar - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (iOBr >= iDot && iCBr >= iDot && errUnit == false)
                {
                    string s_uisq = rr.Substring(iOBr + 1, iCBr - iOBr - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
                if (Curr_Estim == 0 && Curr_Point == 0 && ((iDot <= Math.Min(iOPar, iOBr) && Math.Min(iOPar, iOBr) >= 0) || (Math.Max(iOPar, iOBr) < 0)))
                {
                    strErrNumber[iErr++] = iErr;
                    strErrMessage[iErr] = $"В карточке <{rr}> нет значений";
                    strErrCardURL[iErr] = currShortUrl;
                    errUnit = true;
                    iNone++;
                }
                if (iDot >= Math.Max(iOPar, iOBr) && Math.Max(iOPar, iOBr) > 0)
                {
                    strErrNumber[iErr++] = iErr;
                    strErrMessage[iErr] = $"В карточке <{rr}> нет имени блока";
                    strErrCardURL[iErr] = currShortUrl;
                    errUnit = true;
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
                iErr = iStartErr - 1;
                errUnit = false;
                iLa = 1;
            }
            else
            {
                strErrNumber[iErr++] = iErr;
                strErrMessage[iErr] = $"В карточке <{rr}> нет имени блока";
                strErrCardURL[iErr] = currShortUrl;
                errUnit = true;
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
