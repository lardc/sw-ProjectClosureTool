using System;
using System.IO;
using System.Text;
using System.Text.Json;
using TrlConsCs;

namespace ProjectClosureToolV2
{
    public class Program
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_badgesUtf8 = Encoding.UTF8.GetBytes("badges");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        public static bool labelsListFilled;

        public static void NameToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                // Чтение токена
                reader.Read();

                // Блок? 
                if (reader.CurrentDepth.Equals(2))
                {
                    // Запись оценочных и реальных значений для карточки
                    Trl.Fill_Current_Values(reader.GetString().ToString());
                }
                // Стадия? Команда?
                else if (reader.CurrentDepth.Equals(4))
                {
                    Trl.Search_Departments_Teams(reader.GetString().ToString());
                }
            }
        }

        public static void ShortUrlToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("shortUrl"))
            {
                // Чтение токена
                reader.Read();
                TableResp.currentShortUrl = reader.GetString().ToString();
            }
        }

        public static void BadgesToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("badges"))
            {
                // Чтение токена
                reader.Read();
                TableResp.badges++;
            }
        }

        public static void CardRoleToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("cardRole"))
            {
                // Чтение токена
                reader.Read();
                TableResp.cardRole++;
                Trl.FillParsedValues();
            }
        }

        public static void Help()
        {
            Console.WriteLine("kt - ввод ключа и токена");
            Console.WriteLine("board - ввод кода доски");
            Console.WriteLine("labels - нумерованный список всех имеющихся на доске ярлыков");
            Console.WriteLine("q - выход");
        }

        // Получение ключа и токена
        public static void KeyToken()
        {
            Console.Write("Введите ключ пользователя >");
            try { API_Req.APIKey = Console.ReadLine(); }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
            Console.Write("Введите токен пользователя >");
            try { API_Req.myTrelloToken = Console.ReadLine(); }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }

        // Получение кода для работы с конкретной доской
        public static void Board()
        {
            labelsListFilled = false;
            Console.Write("Введите код доски >");
            try { API_Req.boardCode = Console.ReadLine(); }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"boardCode: {API_Req.boardCode}");
            Console.WriteLine("Press any key");

            Console.ReadKey();
            API_Req.boardURL = "https://trello.com/1/boards/";
            string CardFilter = "/cards/open";
            //string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=2";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            Console.WriteLine("Start");
            Trl.ClearCurrentUnit();

            try
            {
                API_Req.Request(API_Req.APIKey, API_Req.myTrelloToken, CardFilter, CardFields, API_Req.boardCode);
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
                    Console.Write("Нет доступа к доске \nkt - ввод ключа и токена\n");
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
                ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
                var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
                Trl.ClearParsedCardErrorData();
                while (reader.Read())
                {
                    // Тип считанного токена
                    JsonTokenType tokenType;

                    tokenType = reader.TokenType;
                    switch (tokenType)
                    {
                        // Тип токена - начало объекта JSON
                        case JsonTokenType.StartObject:
                            {
                                break;
                            }
                        // Тип токена - название свойства
                        case JsonTokenType.PropertyName:
                            // Это токен "badges"?
                            if (reader.ValueTextEquals(s_badgesUtf8))
                            {
                                try { BadgesToken(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                                if (Trl.parseBadgesToken == true)
                                {
                                    if (Trl.iErrorParse < Trl.iMaxParse) Trl.iErrorParse++;
                                    Trl.parseStrErrorMessage[Trl.iErrorParse] = "Нет конца карточки";
                                    Trl.parseStrErrorCardURL[Trl.iErrorParse] = TableResp.currentShortUrl;
                                    Trl.FillParsedValues();
                                    Trl.parseBadgesToken = true;
                                }
                                else { Trl.parseBadgesToken = true; }
                            }
                            // Это токен "name"?
                            else if (reader.ValueTextEquals(s_nameUtf8))
                            {
                                try { NameToken(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                            }
                            // Это токен "shortUrl"?
                            else if (reader.ValueTextEquals(s_UrlUtf8))
                            {
                                try { ShortUrlToken(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                            }
                            // Это токен "cardRole"? 
                            else if (reader.ValueTextEquals(s_cardRoleUtf8))
                            {
                                try { CardRoleToken(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                                Trl.FillParsedValues();
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

            Trl.FillParsedValues();
            Trl.FillExcel();

            labelsListFilled = true;
        }

        // Нумерованный список всех имеющихся на доске ярлыков
        public static void LabelsList()
        {
            if (labelsListFilled == false)
            {
                Console.WriteLine("Ярлыков нет");
                Console.WriteLine("Press any key");
                Console.ReadKey();
            }
            else
            {
                for (int i = 0; i < Trl.iLabel; i++)
                {
                    Console.WriteLine($"{i + 1}. {Trl.label[i]}");
                }
            }
        }

        static void Main()
        {
            labelsListFilled = false;
            Console.WriteLine("Введите команду");
            string sc = "";
            while (sc != "q")
            {
                Console.Write(" >");
                sc = Console.ReadLine();
                switch (sc)
                {
                    case "help":
                        Help();
                        break;
                    case "kt":
                        KeyToken();
                        break;
                    case "board":
                        Board();
                        break;
                    case "labels":
                        LabelsList();
                        break;
                    case "q":
                        Console.WriteLine("q - выход");
                        break;
                    default:
                        Console.Write("Команда не распознана \nhelp - перечень доступных команд\n");
                        break;
                }
            }

            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}