using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.IO;
using System.Threading.Tasks;
using RestSharp;
using System.Text;
using System.Text.Json;
using Newtonsoft.Json;
using System.Linq;
using OfficeOpenXml;

namespace project_1
{
    public class Trl
    {
        const int iAllUnit = 1000; //max к-во обрабатываемых блоков
        static double[,,] Xi = new double[2, iDep * (iTeams + 1) + 1, iAllUnit + 1];
        // лист, столбец (ярлыки), строка (блок, Total - 0 строка, Unit - с 1-й

        public static string[] Departments = { "Technical Solution", "Development", "Debugging", "Commissioning", "Documentation" };
        const int iDep = 5;
        static int iCurr_Depart = -1; // текущий Department (0-4)

        public static string[] Teams = { "Electronics Team", "Firmware Team", "Remote Team", "Mechanics Team", "Commissioning Team", "Software Team" };
        const int iTeams = 6;
        static int iCurr_Team = -1; // текущий Team (0-5)

        static string[] All_Unit = new string[iAllUnit]; //0 - сумма по столбцу, 1-1000 - сумма по блоку
        static int iAll = 0;  // всего блоков обнаружено
        static int iCurr_Unit = -1; // текущий блок

        static int iTotal;
        static int iCurrTD;
        static int CurrUnit;

        static double Curr_Estim;  // тек. оценочное значение
        static double Curr_Point;  // тек. реальное значение

        string[] Sheets = { "Estimations", "Points" };

        static string[,,] sDept = new string[2, iTeams + 1, iAllUnit + 1];

        public static void Fill()
        {
            if (iCurr_Depart >= 0 && iCurr_Team >= 0 && iCurr_Unit >= 0)
            {
                if (Curr_Estim > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    Console.WriteLine($"iTotal = {iTotal}");
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    Console.WriteLine($"iCurrTD = {iCurrTD}");
                    CurrUnit = iCurr_Unit + 1;
                    Console.WriteLine($"CurrUnit = {CurrUnit}");
                    Xi[0, iCurrTD, CurrUnit] += Curr_Estim;
                    Console.WriteLine($"Xi[0,{iCurrTD},{CurrUnit}] = {Xi[0, iCurrTD, CurrUnit]}");
                    Xi[0, iTotal, CurrUnit] += Curr_Estim;
                    Console.WriteLine($"Xi[0,{iTotal},{CurrUnit}] = {Xi[0, iTotal, CurrUnit]}");
                    Xi[0, iCurrTD, 0] += Curr_Estim;
                    Console.WriteLine($"Xi[0,{iCurrTD},0] = {Xi[0, iCurrTD, 0]}");
                    Xi[0, iTotal, 0] += Curr_Estim;
                    Console.WriteLine($"Xi[0,{iTotal},0] = {Xi[0, iTotal, 0]}");
                }
                if (Curr_Point > 0)
                {
                    iTotal = iCurr_Depart * (iTeams + 1) + iTeams;
                    Console.WriteLine($"iTotal = {iTotal}");
                    iCurrTD = iCurr_Depart * (iTeams + 1) + iCurr_Team;
                    Console.WriteLine($"iCurrTD = {iCurrTD}");
                    CurrUnit = iCurr_Unit + 1;
                    Console.WriteLine($"CurrUnit = {CurrUnit}");
                    Xi[1, iCurrTD, CurrUnit] += Curr_Point;
                    Console.WriteLine($"Xi[1,{iCurrTD},{CurrUnit}] = {Xi[0, iCurrTD, CurrUnit]}");
                    Xi[1, iTotal, CurrUnit] += Curr_Point;
                    Console.WriteLine($"Xi[1,{iTotal},{CurrUnit}] = {Xi[0, iTotal, CurrUnit]}");
                    Xi[1, iCurrTD, 0] += Curr_Point;
                    Console.WriteLine($"Xi[0,{iCurrTD},0] = {Xi[0, iCurrTD, 0]}");
                    Xi[1, iTotal, 0] += Curr_Point;
                    Console.WriteLine($"Xi[0,{iTotal},0] = {Xi[0, iTotal, 0]}");
                }
            }
            Curr_Clear();
        }

        public static void Curr_Clear()
        {
            iCurr_Depart = -1;
            iCurr_Team = -1;
            iCurr_Unit = -1;
            Curr_Estim = 0;
            Curr_Point = 0;
        }
        public static void Search_Depart_Teams(string rr)
        //формируем список Department и Team
        {
            if (Departments.Contains(rr)) { iCurr_Depart = Array.IndexOf(Departments, rr); }
            else if (Teams.Contains(rr)) { iCurr_Team = Array.IndexOf(Teams, rr); }
        }

        public static void Search_Unit(string rr)
        //формируем список Unit
        {
            if (All_Unit.Contains(rr)) { iCurr_Unit = Array.IndexOf(All_Unit, rr); }
            else
            {
                iCurr_Unit = iAll;
                All_Unit[iCurr_Unit] = rr;
                iAll++;
            }
        }

        public static void Fill_Unit_Curr_Val(string rr)
        {
            int s_ln, s_idot;
            s_ln = rr.Length;
            s_idot = rr.IndexOf('.', 0);
            int s_ioro = rr.IndexOf('(', 0);
            int s_iorc = rr.IndexOf(')', 0);
            int s_isqo = rr.IndexOf('[', 0);
            int s_isqc = rr.IndexOf(']', 0);

            if (s_idot >= 0)
            {
                string s_unit = rr.Substring(0, s_idot).Trim();
                Search_Unit(s_unit);

                Console.WriteLine($"{s_unit} = {iCurr_Unit}   <<{rr}>>");

                if (s_ioro >= 0 && s_iorc >= 0)
                {
                    string s_uiro = rr.Substring(s_ioro + 1, s_iorc - s_ioro - 1).Trim();
                    double d_ior = double.Parse(s_uiro);
                    Curr_Estim = d_ior;
                    Console.WriteLine($" Curr_Estim()  = {Curr_Estim}");
                }
                if (s_isqo >= 0 && s_isqc >= 0)
                {
                    string s_uisq = rr.Substring(s_isqo + 1, s_isqc - s_isqo - 1).Trim();
                    double d_isq = double.Parse(s_uisq);
                    Curr_Point = d_isq;
                    Console.WriteLine($" Curr_Point[]  = {Curr_Point}");
                }
            }
            else
            {
                Console.WriteLine($"length  {s_ln} = {rr}, индекс '.' = {s_idot} это НЕ Unit");
            };
        }
        public static void TableDept(int iDept)
        {
            for (int i = 0; i < iAll; i++)
            {
                for (int j = 0; j <= iTeams; j++)
                {
                    sDept[0, j, i] = $"{ Xi[0, iDept * (iTeams + 1) + j, i]}";
                    Console.Write($"{ sDept[0, j, i] }\t");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
            for (int i = 0; i < iAll; i++)
            {
                for (int j = 0; j <= iTeams; j++)
                {
                    sDept[1, j, i] = $"{ Xi[1, iDept * (iTeams + 1) + j, i]}";
                    Console.Write($"{ sDept[0, j, i] }\t");
                }
                Console.WriteLine();
            }
        }

        public static void FillExcel()
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Properties.Author = "KM";
                excelPackage.Workbook.Properties.Title = "Trello";
                excelPackage.Workbook.Properties.Created = DateTime.Now;
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Estimations");

                for (int iD = 0; iD < iDep; iD++)
                {
                    for (int i = 0; i <= iAll; i++)
                    {
                        if (i == iAll)
                            worksheet.Cells[i + 3, 1].Value = "Total";
                        else
                            worksheet.Cells[i + 3, 1].Value = All_Unit[i];
                        for (int j = 0; j <= iTeams; j++)
                        {
                            int jD = iD * (iTeams + 1) + j;
                            if (j == 0 && i == 0)
                                worksheet.Cells[i + 1, jD + 2].Value = Departments[iD];
                            if (j == iTeams && i == 0)
                                worksheet.Cells[i + 2, jD + 2].Value = "Total";
                            if (j < iTeams && i == 0)
                                worksheet.Cells[i + 2, jD + 2].Value = Teams[j];
                            if (i == iAll)
                                worksheet.Cells[i + 3, jD + 2].Value = Xi[0, jD, 0];
                            else
                                worksheet.Cells[i + 3, jD + 2].Value = Xi[0, jD, i + 1];
                        }
                    }

                }

                worksheet = excelPackage.Workbook.Worksheets.Add("Points");

                for (int iD = 0; iD < iDep; iD++)
                {
                    for (int i = 0; i <= iAll; i++)
                    {
                        if (i == iAll)
                            worksheet.Cells[i + 3, 1].Value = "Total";
                        else
                            worksheet.Cells[i + 3, 1].Value = All_Unit[i];
                        for (int j = 0; j <= iTeams; j++)
                        {
                            int jD = iD * (iTeams + 1) + j;
                            if (j == 0 && i == 0)
                                worksheet.Cells[i + 1, jD + 2].Value = Departments[iD];
                            if (j == iTeams && i == 0)
                                worksheet.Cells[i + 2, jD + 2].Value = "Total";
                            if (j < iTeams && i == 0)
                                worksheet.Cells[i + 2, jD + 2].Value = Teams[j];
                            if (i == iAll)
                                worksheet.Cells[i + 3, jD + 2].Value = Xi[1, jD, 0];
                            else
                                worksheet.Cells[i + 3, jD + 2].Value = Xi[1, jD, i + 1];
                        }
                    }
                }
                excelPackage.SaveAs(new FileInfo("EPPlusExcelFromTrello.xlsx"));
            }
        }
    }

    public class Data
    {
        public String id { get; set; }
        public bool display { get; set; }
        public bool entities { get; set; }
        public String fields { get; set; }
        public bool member { get; set; }
    }

    public class Program
    {
        private static object args;
        private static Stream createStream;
        private static String fileName;
        Data data = new Data();

        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static ReadOnlySpan<byte> Utf8Bom => new byte[] { 0xEF, 0xBB, 0xBF };

        static void Main(String[] args)
        {
            TrelloRequestAsync("https://trello.com/b/I5Yr4tAu/software-management-automatization");



            Trl.Curr_Clear();
            string fileName = "response_MME_cards.json";
            ReadOnlySpan<byte> jsonReadOnlySpan = File.ReadAllBytes(fileName);

            if (jsonReadOnlySpan.StartsWith(Utf8Bom))
            {
                jsonReadOnlySpan = jsonReadOnlySpan.Slice(Utf8Bom.Length);
            }

            int count = 0;
            int total = 0;

            var reader = new Utf8JsonReader(jsonReadOnlySpan);

            while (reader.Read())
            {
                JsonTokenType tokenType = reader.TokenType;

                switch (tokenType)
                {
                    case JsonTokenType.StartObject:
                        total++;
                        break;
                    case JsonTokenType.PropertyName:
                        if (reader.ValueTextEquals(s_nameUtf8))
                        {
                            if (reader.GetString().StartsWith("name"))
                            {
                                count++;
                                reader.Read();
                                if (reader.CurrentDepth.Equals(2))
                                { // значения () []
                                    Trl.Fill();
                                    Trl.Curr_Clear();
                                    Trl.Fill_Unit_Curr_Val(reader.GetString().ToString());
                                }
                                else
                                {
                                    Trl.Search_Depart_Teams(reader.GetString().ToString());
                                    string rs = reader.GetString().ToString();
                                    if (Array.Exists(Trl.Departments, element => element == rs))
                                    {
                                        Console.WriteLine($"{rs} = {Array.IndexOf(Trl.Departments, rs)}");
                                    }
                                    if (Array.Exists(Trl.Teams, element => element == rs))
                                    {
                                        Console.WriteLine($"{rs} = {Array.IndexOf(Trl.Teams, rs)}");
                                    }
                                }
                            }
                        }
                        break;
                }
            }

            Trl.TableDept(2);
            Trl.FillExcel();
        }

        static async Task<String> TrelloRequestAsync(String url)
        {
            Console.WriteLine("var tcs");

            var tcs = new TaskCompletionSource<String>();
            try
            {
                Console.WriteLine("try");

                var client = new RestClient(url);

                Console.WriteLine("var client");

                client.GetAsync(new RestRequest(), (response, handle) =>
                {
                    Console.WriteLine("GetAsync");

                    if ((int)response.StatusCode >= 400)
                    {
                        tcs.SetException(new Exception(response.StatusDescription));
                        //Console.Write(">=400");
                    }
                    else
                    {
                        tcs.SetResult(response.Content);
                        //Console.Write("<400");
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception");

                tcs.SetException(ex);
            }
            

            //fileName = @"C:\Users\mb17k\Desktop\project\project1_NET5\response_cards.json";
            //using FileStream openStream = System.IO.File.OpenRead(fileName);

            //String JsonResponseString = openStream.ToString();
            //Data deserializedData = JsonConvert.DeserializeObject<Data>(JsonResponseString);

            //catch (InvalidOperationException e)
            //{

            //}
            //catch (HttpRequestException e)
            //{

            //}
            //catch (TaskCanceledException e)
            //{

            //}

            return await tcs.Task;
        }

        private static Task PostAsJsonAsync(String url, String id)
        {
            throw new NotImplementedException();
        }

        
    }
}
