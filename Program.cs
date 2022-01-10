using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace project_1
{
    public class Trl
    {
        const int iAllUnit = 1000; // максимальное к-во обрабатываемых блоков
        static double[,,] Xi = new double[2, iDep * (iTeams + 1), iAllUnit + 1];
        // 2-Estimations, Points (лист отчета)
        // 35 столбцов таблицы (Total по Department + макс. к-во пар Department/Team)
        // 1000 строки All_Unit + 1(Total)
        // Total - 0 строка, блоки - начинать с 1-й

        public static string[] Department = { "Technical Solution", "Development", "Debugging", "Commissioning", "Documentation" };
        const int iDep = 5;
        static int iCurr_Depart = -1; // текущая стадия (0-4)

        public static string[] Teams = { "Electronics Team", "Firmware Team", "Remote Team", "Mechanics Team", "Commissioning Team", "Software Team" };
        const int iTeams = 6;
        static int iCurr_Team = -1; // текущая команда (0-5)

        static string[] All_Unit = new string[iAllUnit]; //0 - сумма по столбцу, 1-1000 - сумма по блоку
        static int iAll = 0;  // всего блоков обнаружено
        static int iCurr_Unit = -1; // текущий блок

        static int iTotal;
        static int iCurrTD;
        static int CurrUnit;

        static double Curr_Estim;  // тек. оценочное значение
        static double Curr_Point;  // тек. реальное значение

        // Формирование таблиц оценочных и реальных значений
        public static void xiFill()
        {
            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    CurrUnit = iCurr_Unit + 1;
                    Xi[0, iCurrTD, CurrUnit] += Curr_Estim;
                    Xi[0, iTotal, CurrUnit] += Curr_Estim;
                    Xi[0, iCurrTD, 0] += Curr_Estim;
                    Xi[0, iTotal, 0] += Curr_Estim;
                }
                if (Curr_Point > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    CurrUnit = iCurr_Unit + 1;
                    Xi[1, iCurrTD, CurrUnit] += Curr_Point;
                    Xi[1, iTotal, CurrUnit] += Curr_Point;
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
        }

        // Формирование списка стадий и команд
        public static void Search_Depart_Teams(string rr)
        {
            if (iCurr_Unit < 0) return;
            if (Department.Contains(rr))
            {
                if (iCurr_Depart < 0) { iCurr_Depart = Array.IndexOf(Department, rr); }
                // уже есть стадия к текущей карточке
                else
                {
                    Console.WriteLine($"Повтор: карточке из блока {All_Unit[iCurr_Unit] } соответствует более одной стадии ({rr} и {Department[iCurr_Depart]}) ");
                    Curr_Clear();
                };
            }
            else if (Teams.Contains(rr))
            {
                if (iCurr_Team < 0) { iCurr_Team = Array.IndexOf(Teams, rr); }
                // уже есть команда к текущей карточке
                else
                {
                    Console.WriteLine($"Повтор: карточке из блока {All_Unit[iCurr_Unit] } соответствует более одной команды ({rr} и {Teams[iCurr_Team]}) ");
                    Curr_Clear();
                }
            }
        }

        // Формирование списка блоков
        public static void Search_Unit(string rr)
        {
            if (iAll <= 1000)
            {
                if (All_Unit.Contains(rr)) { iCurr_Unit = Array.IndexOf(All_Unit, rr); }
                else
                {
                    iCurr_Unit = iAll;
                    All_Unit[iCurr_Unit] = rr;
                    iAll++;
                }
            }
            else
            {
                Console.WriteLine("Количество блоков превышено");
            }
        }

        // Запись оценочных и реальных значений для обнаруженного блока
        public static void Fill_Unit_Curr_Val(string rr)
        {
            int s_idot;
            s_idot = rr.IndexOf('.', 0);
            int s_ioro = rr.IndexOf('(', 0);
            int s_iorc = rr.IndexOf(')', 0);
            int s_isqo = rr.IndexOf('[', 0);
            int s_isqc = rr.IndexOf(']', 0);

            if (s_idot >= 0)
            {
                string s_unit = rr.Substring(0, s_idot).Trim();
                Search_Unit(s_unit);

                if (s_ioro >= 0 && s_iorc >= 0)
                {
                    string s_uiro = rr.Substring(s_ioro + 1, s_iorc - s_ioro - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                }
                if (s_isqo >= 0 && s_isqc >= 0)
                {
                    string s_uisq = rr.Substring(s_isqo + 1, s_isqc - s_isqo - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                }
            }
        }

        // Запись оценочных и реальных значений в Excel-файл
        public static void FillExcel()
        {
            string fstrin = "json_into_xlsx_t.xltx";
            if (!File.Exists(fstrin))
            {
                Console.WriteLine("Шаблон не существует");
                Console.ReadKey();
                return;
            }
            else
            {
                string fstrout = "json_into_xlsx.xlsx";
                FileInfo fin = new FileInfo(fstrin);
                if (File.Exists(fstrout))
                {
                    try { File.Delete(fstrout); }
                    catch (IOException deleteError)
                    {
                        Console.WriteLine(deleteError.Message);
                        Console.ReadKey();
                        return;
                    }
                }
                FileInfo fout = new FileInfo(fstrout);
                using (var excel_result = new OfficeOpenXml.ExcelPackage(fout, fin))
                {
                    excel_result.Workbook.Properties.Author = "KM";
                    excel_result.Workbook.Properties.Title = "Trello";
                    excel_result.Workbook.Properties.Created = DateTime.Now;
                    double minimumSize = 5;
                    for (int WorkSht = 0; WorkSht <= 1; WorkSht++)
                    {
                        // Лист для записи оценочных значений WorkSht = 0
                        // Лист для записи реальных значений WorkSht = 1
                        ExcelWorksheet estWorksheet = excel_result.Workbook.Worksheets[WorkSht];
                        for (int iD = 0; iD < iDep; iD++)
                        {
                            for (int i = 0; i <= iAll; i++)
                            {
                                if (i == iAll)
                                    estWorksheet.Cells[i + 4, 1].Value = "Total";
                                else
                                    estWorksheet.Cells[i + 4, 1].Value = All_Unit[i];
                                for (int j = 0; j <= iTeams; j++)
                                {
                                    int jD = iD * (iTeams + 1) + j;
                                    if (i == iAll)
                                        estWorksheet.Cells[i + 4, jD + 2].Value = Xi[WorkSht, jD, 0];
                                    else
                                    {
                                        estWorksheet.Cells[i + 4, jD + 2].Value = Xi[WorkSht, jD, i + 1];
                                        if (j < iTeams)
                                        {
                                            estWorksheet.Cells[i + 4, jD + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            estWorksheet.Cells[i + 4, jD + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                        }
                                    }
                                }
                            }
                        }
                        estWorksheet.Cells[estWorksheet.Dimension.Address].AutoFitColumns(minimumSize);
                    }
                    excel_result.Save();
                }
            }
        }
    }
    public class Program
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");

        static void Main(String[] args)
        {
            Trl.Curr_Clear();
            /// Ключ и токен для авторизации
            /// Запрос карточек доски
            /// Получение ответа в виде потока
            /// Запись ответа в память виде строки
            /// Кодировка в UTF-8
            string APIKey = "4b02fbde8c00369dc53e25222e864941";
            string MyTrelloToken = "717ed29e99fcd032275052b563319915f7ce0ec975c5a2abcd965ddd2cf91b07";
            System.Net.WebRequest reqGET = System.Net.WebRequest.Create("https://trello.com/1/boards/dXURQTbH/cards/?key=" + APIKey + "&token=" + MyTrelloToken);
            System.Net.WebResponse resp = reqGET.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader read_stream = new System.IO.StreamReader(stream);
            string readToEnd_string = read_stream.ReadToEnd();
            ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(readToEnd_string);

            var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
            // Тип считанного токена
            JsonTokenType tokenType;

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
                                    // Формирование таблиц оценочных и реальных значений для предыдущей карточки
                                    Trl.xiFill();
                                    // Запись оценочных и реальных значений для текущей карточки
                                    Trl.Fill_Unit_Curr_Val(reader.GetString().ToString());
                                }
                                // Стадия? Команда?
                                else
                                {
                                    Trl.Search_Depart_Teams(reader.GetString().ToString());
                                }
                            }
                        }
                        break;
                }
            }
            Trl.xiFill();
            Trl.FillExcel();
        }
    }
}