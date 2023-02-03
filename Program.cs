﻿using System;
using System.IO;
using System.Text;
using System.Text.Json;
using TrlConsCs;

namespace project_1
{
    public class Program
    {
        //public const int iMaxParse = 100; // максимальное к-во обрабатываемых ошибок в карточке

        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_badgesUtf8 = Encoding.UTF8.GetBytes("badges");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

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
                    Trl.Fill_Unit_Curr_Val(reader.GetString().ToString());
                    Trl.parseUnitToken = reader.GetString();
                }
                // Стадия? Команда?
                else if (reader.CurrentDepth.Equals(4))
                {
                    Trl.Search_Depart_Teams(reader.GetString().ToString());
                }
                Trl.parseShortUrlToken = TableResp.currShortUrl;
            }
        }

        public static void ShortUrlToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("shortUrl"))
            {
                // Чтение токена
                reader.Read();
                TableResp.currShortUrl = reader.GetString().ToString();
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
                Trl.XiParse();
            }
        }

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

            Console.WriteLine($"boardCode: {args[0]}");
            Console.WriteLine($"APIKey: {confProg.APIKey}");
            Console.WriteLine($"myTrelloToken: {confProg.myTrelloToken}");
            Console.WriteLine("Press any key");
            Console.ReadKey();
            API_Req.boardURL = "https://trello.com/1/boards/";
            API_Req.boardCode = args[0];
            API_Req.APIKey = confProg.APIKey;
            API_Req.myTrelloToken = confProg.myTrelloToken;
            string CardFilter = "/cards/open";
            //string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=2";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            Console.WriteLine("Start");
            Trl.Curr_Clear();

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
                ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
                var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
                Trl.ParseClear();
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
                                    if (Trl.iErrParse < Trl.iMaxParse) Trl.iErrParse++;
                                    Trl.parseStrErrMessage[Trl.iErrParse] = "Нет конца карточки";
                                    Trl.parseStrErrCardURL[Trl.iErrParse] = TableResp.currShortUrl;
                                    Trl.XiParse();
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
                                Trl.parseShortUrlToken = reader.GetString().ToString();
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
                                Trl.parseCardRoleToken = true;
                                Trl.XiParse();
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
            Trl.XiParse();
            Trl.FillExcel();
            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}