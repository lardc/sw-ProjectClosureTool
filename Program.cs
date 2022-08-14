using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Linq;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Diagnostics;

namespace project_1
{
    public static class API_Req
    {
        private static string readToEnd_string;

        public static string boardURL;
        public static string boardCode;
        public static string APIKey;
        public static string MyTrelloToken;

        public static string ReadToEnd_string { get => readToEnd_string; set => readToEnd_string = value; }

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

    public abstract class TableResp
    {
        public abstract void XiFill();

        public abstract void FillExcel();

        private static OfficeOpenXml.ExcelPackage excel_result;

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
                    }
                }
                estWorksheet.Cells[estWorksheet.Dimension.Address].AutoFitColumns(MinimumSize);
            }
            OfficeOpenXml.ExcelWorksheet errWorksheet = Excel_result.Workbook.Worksheets[WorkShtN];

            for (int i = 0; i <= iErr; i++) { errWorksheet.Cells[i + 2, 1].Value = strErr[i]; }
        }

        public const int iAllUnits = 1000; // максимальное к-во обрабатываемых блоков
        private static double[,,] xi = new double[2, iDep * (iTeams + 1), iAllUnits + 1];
        // 2-Estimations, Points (лист отчета)
        // 35 столбцов таблицы (Total по Department + макс. к-во пар Department/Team)
        // 1000 строки All_Units + 1(Total)
        // Total - 0 строка, блоки - начинать с 1-й

        private static string[] departments = { "Technical Solution", "Development", "Debugging", "Commissioning", "Documentation" };
        public const int iDep = 5;
        public static int iCurr_Depart = -1; // текущая стадия (0-4)

        private static string[] teams = { "Electronics Team", "Firmware Team", "Remote Team", "Mechanics Team", "Commissioning Team", "Software Team" };
        public const int iTeams = 6;
        public static int iCurr_Team = -1; // текущая команда (0-5)

        private static string[] all_Units = new string[iAllUnits + 1]; //0 - сумма по столбцу, 1-1000 - сумма по блоку
        public static int iAll = 0;  // всего блоков обнаружено
        public static int iCurr_Unit = -1; // текущий блок

        public static string currCardName;
        public static string currShortUrl;

        public static int iTotal;
        public static int iCurrTD;
        private static int currUnit;

        public static bool errUnit = false;

        private static double curr_Estim;  // текущее оценочное значение
        private static double curr_Point;  // текущее реальное значение

        private static double minimumSize;

        public static int iErr = 0;
        public static int iStartErr = 0;

        public static string[] strErr = new string[1000];

        public static ExcelPackage Excel_result { get => excel_result; set => excel_result = value; }
        public static double[,,] Xi { get => xi; set => xi = value; }
        public static string[] Departments { get => departments; set => departments = value; }
        public static string[] Teams { get => teams; set => teams = value; }
        public static string[] All_Units { get => all_Units; set => all_Units = value; }
        public static int CurrUnit { get => currUnit; set => currUnit = value; }
        public static double Curr_Estim { get => curr_Estim; set => curr_Estim = value; }
        public static double Curr_Point { get => curr_Point; set => curr_Point = value; }
        public static double MinimumSize { get => minimumSize; set => minimumSize = value; }
    }

    public class Trl : TableResp
    {
        // Формирование таблиц оценочных и реальных значений
        public override void XiFill()
        {
            if (errUnit) { strErr[iErr++] = currShortUrl; }
            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    CurrUnit = iCurr_Unit + 1;
                    Xi[0, iCurrTD, CurrUnit] += Curr_Estim;
                    Xi[0, iCurrTD, 0] += Curr_Estim;
                    Xi[0, iTotal, 0] += Curr_Estim;
                }
                if (Curr_Point > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    CurrUnit = iCurr_Unit + 1;
                    Xi[1, iCurrTD, CurrUnit] += Curr_Point;
                    Xi[1, iCurrTD, 0] += Curr_Point;
                    Xi[1, iTotal, 0] += Curr_Point;
                }
            }
            Curr_Clear();
        }

        //Очистка значений счётчиков по текущему блоку
        public static void Curr_Clear()
        {
            iCurr_Depart = -1;
            iCurr_Team = -1;
            iCurr_Unit = -1;
            Curr_Estim = 0;
            Curr_Point = 0;
            errUnit = false;
            iStartErr = iErr;
        }

        // Формирование списка стадий и команд
        public static void Search_Depart_Teams(string rr)
        {
            if (iCurr_Unit < 0)
            {
                //Ярлык без блока
                strErr[iErr++] = $"Ярлык <{rr}> не привязан к блоку";
                errUnit = true;
                Console.WriteLine($"iCurr_Unit == {iCurr_Unit}");
                Console.WriteLine($"Ярлык <{rr}> не привязан к блоку");
                return;
            }

            if (Departments.Contains(rr))
            {
                // if (iCurr_Depart < 1) { iCurr_Depart = Array.IndexOf(Departments, rr); }
                if (iCurr_Depart < 0) { iCurr_Depart = Array.IndexOf(Departments, rr); }
                // уже есть стадия к текущей карточке
                else
                {
                    //strErr[iErr++] = $"Повтор: карточке из блока <{All_Units[iCurr_Unit]}> соответствует более одной стадии ({rr} и {Departments[iCurr_Depart]}) ";
                    strErr[iErr++] = $"Повтор: карточке соответствует более одной стадии";
                    //strErr[iErr++] = $"Повтор: карточке соответствует более одной стадии ({rr} и {Departments[iCurr_Depart]}) ";
                    errUnit = true;
                    Console.WriteLine($"iCurr_Depart == {iCurr_Depart}");
                    Console.WriteLine($"Повтор: карточке соответствует более одной стадии");
                    //Console.WriteLine($"Повтор: карточке соответствует более одной стадии ({rr} и {Departments[iCurr_Depart]}) ");
                };
            }
            else if (Teams.Contains(rr))
            {
                if (iCurr_Team < 0) { iCurr_Team = Array.IndexOf(Teams, rr); }
                // уже есть команда к текущей карточке
                else
                {
                    strErr[iErr++] = $"Повтор: карточке соответствует более одной команды";
                    //strErr[iErr++] = $"Повтор: карточке соответствует более одной команды ({rr} и {Teams[iCurr_Team]}) ";
                    errUnit = true;
                    Console.WriteLine($"iCurr_Team == {iCurr_Team}");
                    Console.WriteLine($"Повтор: карточке соответствует более одной команды");
                    //Console.WriteLine($"Повтор: карточке соответствует более одной команды ({rr} и {Teams[iCurr_Team]}) ");
                }
            }
            else
            {
                strErr[iErr++] = $"Ярлык <{rr}> не соответствует полям таблицы";
                errUnit = true;
                Console.WriteLine($"Ярлык <{rr}> не соответствует полям таблицы");
            }
        }

        // Формирование списка блоков
        public static void Search_Unit(string rr)
        {
            if (iAll <= 1000)
            {
                //if (All_Units.Contains(rr)) { iCurr_Unit = Array.IndexOf(All_Units, rr) + 1; }
                if (All_Units.Contains(rr)) { iCurr_Unit = Array.IndexOf(All_Units, rr); }
                else
                {
                    iCurr_Unit = iAll;
                    All_Units[iCurr_Unit] = rr;
                    iAll++;
                }
            }
            else
            {
                strErr[iErr++] = "Количество блоков превышено";
                errUnit = true;
                Console.WriteLine($"iAll == {iAll}");
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
                if (s_ioro >= 0 && s_iorc == 0)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет закрывающей круглой скобки";
                    Console.WriteLine($"s_ioro == {s_ioro}");
                    Console.WriteLine($"s_iorc == {s_iorc}");
                    Console.WriteLine($"В карточке <{rr}> нет закрывающей круглой скобки");
                }
                if (s_isqo >= 0 && s_isqc == 0)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет закрывающей квадратной скобки";
                    Console.WriteLine($"s_isqo == {s_isqo}");
                    Console.WriteLine($"s_isqc == {s_isqc}");
                    Console.WriteLine($"В карточке <{rr}> нет закрывающей квадратной скобки");
                }

                if (s_ioro >= s_idot && s_iorc >= s_idot)
                //if (s_ioro >= 0 && s_iorc >= 0)
                {
                    string s_uiro = rr.Substring(s_ioro + 1, s_iorc - s_ioro - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (s_isqo >= s_idot && s_isqc >= s_idot)
                //if (s_isqo >= 0 && s_isqc >= 0)
                {
                    string s_uisq = rr.Substring(s_isqo + 1, s_isqc - s_isqo - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
                if (Curr_Estim == 0 && Curr_Point == 0 && ((s_idot <= Math.Min(s_ioro, s_isqo) && Math.Min(s_ioro, s_isqo) > 0) || (s_idot <= Math.Max(s_ioro, s_isqo) && Math.Min(s_ioro, s_isqo) == 0)))
                //if (Curr_Estim == 0 && Curr_Point == 0 && s_idot <= Math.Max(s_ioro, s_isqo))
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет значений";
                    errUnit = true;
                    Console.WriteLine($"Curr_Estim == {Curr_Estim}");
                    Console.WriteLine($"Curr_Point == {Curr_Point}");
                    Console.WriteLine($"s_idot == {s_idot}");
                    Console.WriteLine($"s_ioro == {s_ioro}");
                    Console.WriteLine($"s_isqo == {s_isqo}");
                    Console.WriteLine($"В карточке <{rr}> нет значений");
                }
                if (s_idot >= Math.Max(s_ioro, s_isqo) && Math.Max(s_ioro, s_isqo) > 0)
                {
                    strErr[iErr++] = $"В карточке <{rr}> нет имени блока";
                    errUnit = true;
                    Console.WriteLine($"s_idot == {s_idot}");
                    Console.WriteLine($"s_ioro == {s_ioro}");
                    Console.WriteLine($"s_isqo == {s_isqo}");
                    Console.WriteLine($"В карточке <{rr}> нет имени блока");
                }
                else
                {
                    string s_unit = rr.Substring(0, s_idot).Trim();
                    Search_Unit(s_unit);
                }
            }
            else if (rr == "Labels" || rr == "ЭМ")
            {
                Console.WriteLine($"Служебная карточка <{rr}>");
                iErr = iStartErr - 1;
                errUnit = false;

            }
            else
            {
                strErr[iErr++] = $"В карточке <{rr}> нет имени блока";
                errUnit = true;
                Console.WriteLine($"В карточке <{rr}> нет имени блока");
            }
        }

        // Запись оценочных и реальных значений в Excel-файл
        public override void FillExcel()
        {
            if (iAll == 0)
            {
                Console.WriteLine($"В доске {API_Req.boardURL}{API_Req.boardCode} нет ни одного значения. Файл не сформирован.");
                Console.WriteLine("Press any key");
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
                DateTime DTnow = DateTime.Now;
                string DTyear = DTnow.Year.ToString();
                string DTmonth = DTnow.Month.ToString();
                string DTday = DTnow.Day.ToString();
                string DThour = DTnow.Hour.ToString();
                string DTminute = DTnow.Minute.ToString();
                string DTsecond = DTnow.Second.ToString();
                string fstrout = $"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx";
                FileInfo fin = new(fstrin);
                if (File.Exists(fstrout))
                {
                    try { File.Delete(fstrout); }
                    catch (IOException deleteError)
                    {
                        Console.WriteLine(deleteError.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
                }
                FileInfo fout = new(fstrout);
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
    }
    public class Program
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");

        public static void StartExtractor()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "project 1.exe";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "dXURQTbH";
            Process.Start(startInfo);
        }

        static void Main(string[] args)
        {
            //Console.Read();
            Console.WriteLine("Start");
            Trl.Curr_Clear();
            API_Req.boardURL = "https://trello.com/1/boards/";
            API_Req.boardCode = "dXURQTbH";
            string APIKey = "4b02fbde8c00369dc53e25222e864941";
            string MyTrelloToken = "717ed29e99fcd032275052b563319915f7ce0ec975c5a2abcd965ddd2cf91b07";
            string CardFilter = "/cards/open";
            //string CardFields = "&fiels=id,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,url&limit=2";
            string CardFields = "&fiels=id,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,url&limit=1000";

            if (args.Length == 0) { }
            else if (args.Length == 1)
            {
                API_Req.boardCode = args[0];
                Console.WriteLine(API_Req.boardCode);
            }
            else if (args.Length == 3)
            {
                API_Req.boardCode = args[0];
                Console.WriteLine(API_Req.boardCode);
                APIKey = args[1];
                MyTrelloToken = args[2];
            }
            else
            {
                Console.WriteLine("Некорректное количество аргументов");
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }

            try
            {
                API_Req.Request(APIKey, MyTrelloToken, CardFilter, CardFields, API_Req.boardCode);
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

            //Console.WriteLine(API_Req.ReadToEnd_string); // что прочли - отладка

            ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.ReadToEnd_string);
            var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);

            // Тип считанного токена
            JsonTokenType tokenType;

            var xiFillAbstr = new Trl();
            var FillExcelSheetsSt = new Trl();

            while (reader.Read())
            {
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
                                    // Запись оценочных и реальных значений для текущей карточки
                                    Trl.Fill_Unit_Curr_Val(reader.GetString().ToString());
                                    // Формирование таблиц оценочных и реальных значений для текущей карточки
                                    xiFillAbstr.XiFill();

                                    TableResp.currCardName = reader.GetString().ToString();
                                }
                                // Стадия? Команда?
                                else
                                {
                                    Trl.Search_Depart_Teams(reader.GetString().ToString());
                                }
                            }
                        }
                        else if (reader.ValueTextEquals(s_UrlUtf8))
                        {
                            if (reader.GetString().StartsWith("shortUrl"))
                            {
                                // Чтение токена
                                reader.Read();
                                TableResp.currShortUrl = reader.GetString().ToString();
                            }
                        }
                        break;
                }
            }
            xiFillAbstr.XiFill();
            FillExcelSheetsSt.FillExcel();
            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}