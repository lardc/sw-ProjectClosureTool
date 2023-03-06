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

        private static bool labelsListFilled;
        private static bool ignoredLabelsListFilled;

        private static int nIgnore;
        private static int iInputIgnore;
        private static int iDeleteIgnore;

        private static bool isLabel;
        private static bool isIgnored;

        public static void NameToken(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                // Чтение токена
                reader.Read();

                // Блок? 
                if (reader.CurrentDepth.Equals(2))
                    // Запись оценочных и реальных значений для карточки
                    Trl.Fill_Current_Values(reader.GetString().ToString());
                // Стадия? Команда?
                else if (reader.CurrentDepth.Equals(4))
                    Trl.Search_Departments_Teams(reader.GetString().ToString());
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
            Console.WriteLine("addI - создание списка игнорируемых ярлыков");
            Console.WriteLine("listI - просмотр списка игнорируемых ярлыков");
            Console.WriteLine("deleteI - удаление игнорируемых ярлыков по номеру");
            Console.WriteLine("clearI - очистка списка игнорируемых ярлыков");
            Console.WriteLine("fillExcel - формирование excel-файла");
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

        public static void Board()
        {
            Trl.ClearCurrentUnit();
            Trl.ClearValues();
            IgnoredLabelsList(nIgnore);

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
                                break;
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
                                if (Trl.parseBadgesToken)
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
            labelsListFilled = true;
        }

        // Получение кода для работы с конкретной доской
        public static void ReadBoard()
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

            ClearIgnoredLabels();

            Console.WriteLine($"boardCode: <{API_Req.boardCode}>");
            Console.WriteLine("Press any key");

            Console.ReadKey();
            API_Req.boardURL = "https://trello.com/1/boards/";
            string CardFilter = "/cards/open";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            Console.WriteLine("Start");
            

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
        }

        public static void ReadLabels()
        {
            try 
            {
                ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
                var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
                Trl.ClearParsedCardErrorData();
                ClearLabels();
                while (reader.Read())
                {
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
                            if (reader.ValueTextEquals(s_nameUtf8))
                            {
                                if (reader.GetString().StartsWith("name"))
                                try
                                {
                                    reader.Read(); 
                                    if (reader.CurrentDepth.Equals(4)) Trl.Input_Labels(reader.GetString().ToString());
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
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
            labelsListFilled = true;
        }

        // Нумерованный список всех имеющихся на доске ярлыков
        public static void LabelsList()
        {
            if (!labelsListFilled)
            {
                Console.WriteLine("Ярлыков нет. Выполните команду ввода кода доски.");
                Console.WriteLine("Press any key");
                Console.ReadKey();
            }
            else
            {
                for (int i = 0; i < Trl.iLabels; i++)
                {
                    CheckIgnored(Trl.labels[i]);
                    if (!isIgnored) Console.WriteLine($"{i + 1}. {Trl.labels[i]}");
                }
            }
        }

        // Создание списка игнорируемых ярлыков
        public static void AddIgnore()
        {
            LabelsList();
            Console.Write("Введите номер игнорируемого ярлыка >");
            try { iInputIgnore = int.Parse(Console.ReadLine()); }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
            while (iInputIgnore != 0)
            {
                CheckLabels(Trl.labels[iInputIgnore - 1]);
                if (!isLabel || iInputIgnore > Trl.iLabels)
                    Console.WriteLine($"Нет ярлыка с номером <{iInputIgnore}>. Введите другой номер.");
                else
                { 
                    if (isIgnored)
                        Console.WriteLine($"Ярлык <{Trl.labels[iInputIgnore - 1]}> уже содержится в списке игнорируемых");
                    else
                    {
                        Trl.ignoredLabels[nIgnore++] = Trl.labels[iInputIgnore - 1];
                        ignoredLabelsListFilled = true;
                    }
                }
                
                Console.Write("Введите номер игнорируемого ярлыка >");
                try { iInputIgnore = int.Parse(Console.ReadLine()); }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
            }
        }

        // Просмотр списка игнорируемых ярлыков
        public static void IgnoredLabelsList(int n)
        {
            if (!ignoredLabelsListFilled)
                Console.WriteLine("Игнорируемых ярлыков нет");
            else
                for (int i = 0; i < n; i++)
                {
                    CheckLabels(Trl.ignoredLabels[i]);
                    if (!isLabel) Console.WriteLine($"Игнорируемый ярлык <{Trl.ignoredLabels[i]}> отсутствует в списке ярлыков. Перезагрузите доску.");
                    else Console.WriteLine($"{i + 1}. {Trl.ignoredLabels[i]}");
                }    
        }

        // Удаление игнорируемых ярлыков по номеру
        public static void DeleteIgnoredLabel()
        {
            if (!ignoredLabelsListFilled)
                Console.WriteLine("Игнорируемых ярлыков нет");
            else
            {
                Console.Write("Удаление ярлыков по номеру \nДля выхода введите 0\n");
                IgnoredLabelsList(nIgnore);
                Console.Write("Введите номер удаляемого из списка игнорируемых ярлыка >");
                try { iDeleteIgnore = int.Parse(Console.ReadLine()); }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                    return;
                }
                while (iDeleteIgnore != 0)
                {
                    if (iDeleteIgnore > nIgnore)
                    {
                        Console.WriteLine($"Нет игнорируемого ярлыка с номером <{iDeleteIgnore}>. Введите другой номер.");
                    }
                    else
                    {
                        for (int i = iDeleteIgnore - 1; i < nIgnore - 1; i++)
                        {
                            try { Trl.ignoredLabels[i] = Trl.ignoredLabels[i + 1]; }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                                Console.WriteLine("Press any key");
                                Console.ReadKey();
                                return;
                            }
                        }
                        Trl.ignoredLabels[--nIgnore] = "";
                    }
                    if (nIgnore < 1)
                    { 
                        ignoredLabelsListFilled = false;
                        iDeleteIgnore = 0;
                        IgnoredLabelsList(nIgnore);
                    }
                    else
                    {
                        IgnoredLabelsList(nIgnore);
                        Console.Write("Введите номер удаляемого из списка игнорируемых ярлыка >");
                        try { iDeleteIgnore = int.Parse(Console.ReadLine()); }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine("Press any key");
                            Console.ReadKey();
                            return;
                        }
                    }
                }
            }
        }

        // Очистка списка игнорируемых ярлыков
        public static void ClearIgnoredLabels()
        {
            for (int i = 0; i < TableResp.iAllUnits; i++) Trl.ignoredLabels[i] = "";
            nIgnore = 0;
            ignoredLabelsListFilled = false;
        }

        public static void ClearLabels()
        {
            for (int i = 0; i < TableResp.iAllUnits; i++) Trl.labels[i] = "";
            Trl.iLabels = 0;
            TableResp.iAll = 0;
        }

        public static bool CheckLabels(string rr)
        {
            if (Trl.labels.Contains(rr)) isLabel = true; else isLabel = false;
            return isLabel;
        }

        public static bool CheckIgnored(string rr)
        {
            if (Trl.ignoredLabels.Contains(rr)) isIgnored = true; else isIgnored = false;
            return isIgnored;
        }

        static void Main()
        {
            labelsListFilled = false;
            Console.Write("help - перечень доступных команд \nВведите команду\n");
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
                        ClearIgnoredLabels();
                        ClearLabels();
                        ReadBoard();
                        ReadLabels();
                        break;
                    case "labels":
                        Console.WriteLine("Перечень ярлыков:");
                        LabelsList();
                        break;
                    case "addI":
                        Console.Write("Создание списка игнорируемых ярлыков \nДля выхода введите 0\n");
                        AddIgnore();
                        break;
                    case "listI":
                        Console.WriteLine("Перечень игнорируемых ярлыков:");
                        IgnoredLabelsList(nIgnore);
                        break;
                    case "deleteI":
                        DeleteIgnoredLabel();
                        break;
                    case "clearI":
                        ClearIgnoredLabels();
                        break;
                    case "fillExcel":
                        if (!labelsListFilled) Console.WriteLine("Код доски не введён");
                        else 
                        { 
                            Console.WriteLine("Формирование excel-файла");
                            Board();
                            Trl.FillExcel();
                        }
                        break;
                    case "q":
                        Console.WriteLine("q - выход");
                        break;
                    default:
                        Console.Write("Команда не распознана");
                        break;
                }
            }

            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}