using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Linq;
using OfficeOpenXml.Style;
using OfficeOpenXml;

namespace project_1
{
    public static class API_Req
    {
        public static string boardURL;
        public static string boardCode;
        public static string APIKey;
        public static string MyTrelloToken;

        public static string ReadToEnd_string;

        /// Ключ и токен для авторизации
        /// Запрос карточек доски
        /// Получение ответа в виде потока
        /// Запись ответа в память виде строки
        /// Кодировка в UTF-8
        public static void Request(string APIKey, string MyTrelloToken, string CardFilter, string CardFields, string boardCode)
        {
            System.Net.WebRequest reqGET = System.Net.WebRequest.Create(boardURL + boardCode + CardFilter + "/?key=" + APIKey + "&token=" + MyTrelloToken + CardFields);
            System.Net.WebResponse resp = reqGET.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader read_stream = new(stream);
            ReadToEnd_string = read_stream.ReadToEnd();
        }
    }

    public class TableResp
    {
        public static OfficeOpenXml.ExcelPackage excel_result;

        public static void FillExcelSheets(int WorkShtN)
        {
            for (int WorkSht = 0; WorkSht < WorkShtN; WorkSht++)
            {
                // Лист для записи оценочных значений WorkSht = 0
                // Лист для записи реальных значений WorkSht = 1
                // Лист для записи ошибок WorkSht = WorkShtN
                OfficeOpenXml.ExcelWorksheet estWorksheet = Excel_result.Workbook.Worksheets[WorkSht];
                for (int iD = 0; iD < iDep; iD++)
                {
                    for (int i = 0; i <= iAll; i++)
                    {
                        if (i == iAll)
                            estWorksheet.Cells[i + 4, 1].Value = "Total";
                        else
                            estWorksheet.Cells[i + 4, 1].Value = All_Units[i];
                        for (int j = 0; j <= iTeams; j++)
                        {
                            FillExcelCells(i, j, iD, estWorksheet, WorkSht);
                        }
                    }
                }
                estWorksheet.Cells[estWorksheet.Dimension.Address].AutoFitColumns(MinimumSize);
            }
            OfficeOpenXml.ExcelWorksheet errWorksheet = Excel_result.Workbook.Worksheets[WorkShtN];

            for (int i = 0; i <= iErr; i++) { errWorksheet.Cells[i + 2, 1].Value = strErr[i]; }
        }

        public static void FillExcelCells(int i, int j, int iD, OfficeOpenXml.ExcelWorksheet estWorksheet, int WorkSht)
        {
            int jD = iD * (iTeams + 1) + j;
            if (i == iAll)
                estWorksheet.Cells[iAll + 4, jD + 2].Formula = "=SUM(" + estWorksheet.Cells[4, jD + 2].Address + ":" + estWorksheet.Cells[iAll + 3, jD + 2].Address + ")";
            else if (j < iTeams)
            {
                estWorksheet.Cells[i + 4, jD + 2].Value = Xi[WorkSht, jD, i + 1];
                estWorksheet.Cells[i + 4, jD + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                estWorksheet.Cells[i + 4, jD + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
            }
            ExcelRange rg = estWorksheet.Cells[iAll + 5, 1, iAllUnits + 5, jD + 2];
            rg.Clear();
        }

        // Количество блоков превышено ?
        //public const int iAllUnits = 2; // максимальное к-во обрабатываемых блоков
        public const int iAllUnits = 1000; // максимальное к-во обрабатываемых блоков

        public static double[,,] Xi = new double[2, iDep * (iTeams + 1), iAllUnits + 5];
        // 2-Estimations, Points (лист отчета)
        // 35 столбцов таблицы (Total по Department + макс. к-во пар Department/Team)
        // 1000 строки All_Units + 1(Total)
        // Total - 0 строка, блоки - начинать с 1-й

        public static string[] Departments = { "Technical Solution", "Development", "Debugging", "Commissioning", "Documentation" };
        public const int iDep = 5;
        public static int iCurr_Depart = -1; // текущая стадия (0-4)

        public static string[] Teams = { "Electronics Team", "Firmware Team", "Remote Team", "Mechanics Team", "Commissioning Team", "Software Team" };
        public const int iTeams = 6;
        public static int iCurr_Team = -1; // текущая команда (0-5)

        public static string[] All_Units = new string[iAllUnits + 5]; //0 - сумма по столбцу, 1-1000 - сумма по блоку
        public static int iAll = 0;  // всего блоков обнаружено
        public static int iCurr_Unit = -1; // текущий блок

        public static string currCardName;
        public static string currShortUrl;

        public static int iTotal;
        public static int iCurrTD;
        public static int currUnit;

        public static bool errUnit = false;

        public static double curr_Estim;  // текущее оценочное значение
        public static double curr_Point;  // текущее реальное значение

        public static double minimumSize;

        public static int iErr = 0;
        public static int iStartErr = 0;

        public static int bgs;
        public static int cRe;
        public static int iLa;

        public static int iNone;

        public static string[] strErr = new string[iAllUnits + 5];

        public static ExcelPackage Excel_result { get; set; }
        public static int CurrUnit { get; set; }
        public static double Curr_Estim { get; set; }
        public static double Curr_Point { get; set; }
        public static double MinimumSize { get; set; }
    }

    public class Trl : TableResp
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
                strErr[iErr++] = "Нет начала карточки";
                if (iLa < 1) { strErr[iErr++] = currShortUrl; }
            }
        }

        // Проверка наличия конца карточки
        public static void Check_cRe()
        {
            if (Trl.cRe == 0)
            {
                strErr[iErr++] = "Нет начала карточки";
                if (iLa < 1) { strErr[iErr++] = currShortUrl; }
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
            if (errUnit) { strErr[iErr++] = currShortUrl; }

            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0) { XiFill_Curr_Estim(); }
                if (Curr_Point > 0) { XiFill_Curr_Point(); }
            }

            if (iCurr_Depart < 0 && iCurr_Unit >= 0)
            {
                strErr[iErr++] = "Нет ярлыка стадии";
                strErr[iErr++] = currShortUrl;
            }

            if (iCurr_Team < 0 && iCurr_Unit >= 0)
            {
                strErr[iErr++] = "Нет ярлыка команды";
                strErr[iErr++] = currShortUrl;
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
                strErr[iErr++] = $"Ярлык <{rr}> не соответствует полям таблицы";
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
                        strErr[iErr++] = $"Повтор: карточке соответствует более одной стадии";
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
                        strErr[iErr++] = $"Повтор: карточке соответствует более одной команды";
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
                strErr[iErr++] = "Количество блоков превышено ( >" + iAllUnits + ")";
                errUnit = true;
                Console.WriteLine("Количество блоков превышено");
            }
        }

        // Запись оценочных и реальных значений для обнаруженного блока
        public static void Fill_Unit_Curr_Val(string rr)
        {
            int s_idot = rr.IndexOf('.', 0);
            int s_ioro = rr.IndexOf('(', 0);
            int s_iorc = rr.IndexOf(')', 0);
            int s_isqo = rr.IndexOf('[', 0);
            int s_isqc = rr.IndexOf(']', 0);

            if (s_idot >= 0)
            {
                if (s_ioro >= 0 && s_iorc < s_ioro)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет закрывающей круглой скобки";
                }
                if (s_isqo >= 0 && s_isqc < s_isqo)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет закрывающей квадратной скобки";
                }

                if (s_ioro >= s_idot && s_iorc >= s_idot)
                {
                    string s_uiro = rr.Substring(s_ioro + 1, s_iorc - s_ioro - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (s_isqo >= s_idot && s_isqc >= s_idot)
                {
                    string s_uisq = rr.Substring(s_isqo + 1, s_isqc - s_isqo - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
                if (Curr_Estim == 0 && Curr_Point == 0 && ((s_idot <= Math.Min(s_ioro, s_isqo) && Math.Min(s_ioro, s_isqo) >= 0) || (Math.Max(s_ioro, s_isqo) < 0)))
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет значений";
                    errUnit = true;
                    iNone++;
                }
                if (s_idot >= Math.Max(s_ioro, s_isqo) && Math.Max(s_ioro, s_isqo) > 0)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет имени блока";
                    errUnit = true;
                }
                else
                {
                    string s_unit = rr.Substring(0, s_idot).Trim();
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
                strErr[iErr++] = $"В карточке <{rr}> нет имени блока";
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

    public class ConfProg
    {
        public string BoardCode { get; set; }
        public string APIKey { get; set; }
        public string MyTrelloToken { get; set; }
    }

    public class Program
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_badgesUtf8 = Encoding.UTF8.GetBytes("badges");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        static void Main(string[] args)
        {
            string fileName = "config.json";
            if (!File.Exists(fileName))
            {
                Console.WriteLine("Нет конфигурационного файла");
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
            string jsonString = File.ReadAllText(fileName);
            ConfProg confProg = JsonSerializer.Deserialize<ConfProg>(jsonString)!;

            Console.WriteLine($"boardCode: {confProg.BoardCode}");
            Console.WriteLine($"APIKey: {confProg.APIKey}");
            Console.WriteLine($"MyTrelloToken: {confProg.MyTrelloToken}");
            Console.WriteLine("Press any key");
            Console.ReadKey();
            API_Req.boardURL = "https://trello.com/1/boards/";
            API_Req.boardCode = confProg.BoardCode;
            API_Req.APIKey = confProg.APIKey;
            API_Req.MyTrelloToken = confProg.MyTrelloToken;
            string CardFilter = "/cards/open";
            //string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=2";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            Console.WriteLine("Start");
            Trl.Curr_Clear();

            try
            {
                API_Req.Request(API_Req.APIKey, API_Req.MyTrelloToken, CardFilter, CardFields, API_Req.boardCode);
            }
            catch (Exception e)
            {
                if (e.Message.Contains("400"))
                {
                    Trl.EMessage(e.Message);
                    Console.WriteLine("Недопустимый url доски");
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
                if (e.Message.Contains("401"))
                {
                    Trl.EMessage(e.Message);
                    Console.WriteLine("Нет доступа к доске");
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
                if (e.Message.Contains("404"))
                {
                    Trl.EMessage(e.Message);
                    Console.WriteLine("Доска не найдена");
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
            }

            try
            {
                ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.ReadToEnd_string);
                var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
                while (reader.Read())
                {
                    // Тип считанного токена
                    JsonTokenType tokenType;

                    tokenType = reader.TokenType;
                    switch (tokenType)
                    {
                        // Тип токена - начало объекта JSON
                        case JsonTokenType.StartObject:
                            break;
                        // Тип токена - название свойства
                        case JsonTokenType.PropertyName:
                            // Это токен "name"?
                            if (reader.ValueTextEquals(s_nameUtf8))
                            {
                                if (reader.GetString().StartsWith("name"))
                                {
                                    // Чтение токена
                                    reader.Read();

                                    // Блок? 
                                    if (reader.CurrentDepth.Equals(2))
                                    {
                                        // Запись оценочных и реальных значений для карточки
                                        Trl.Fill_Unit_Curr_Val(reader.GetString().ToString());
                                    }
                                    // Стадия? Команда?
                                    else if (reader.CurrentDepth.Equals(4))
                                    {
                                        Trl.Search_Depart_Teams(reader.GetString().ToString());
                                    }
                                }
                            }
                            else if (reader.ValueTextEquals(s_UrlUtf8))
                            {
                                // Это токен "shortUrl"?
                                if (reader.GetString().StartsWith("shortUrl"))
                                {
                                    // Чтение токена
                                    reader.Read();
                                    TableResp.currShortUrl = reader.GetString().ToString();
                                }
                            }
                            else if (reader.ValueTextEquals(s_badgesUtf8))
                            {
                                // Это токен "badges"?
                                if (reader.GetString().StartsWith("badges"))
                                {
                                    // Чтение токена
                                    reader.Read();
                                    TableResp.bgs++;
                                    if (TableResp.bgs > 1) { Console.WriteLine("Повторное обнаружение начала карточки"); }
                                }
                            }
                            else if (reader.ValueTextEquals(s_cardRoleUtf8))
                            {
                                // Это токен "cardRole"?
                                if (reader.GetString().StartsWith("cardRole"))
                                {
                                    // Чтение токена
                                    reader.Read();
                                    Trl.cRe++;
                                    Trl.XiFill();
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
            Trl.XiFill();
            Trl.FillExcel();
            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}